"""Unit tests for the pure helpers in generate_dashboard.py."""
import datetime
import json

import pytest

import generate_dashboard as gd


# ---------------------------------------------------------------- dates
@pytest.mark.parametrize("value,expected", [
    (None, None),
    ("", None),
    ("no year here", None),
    ("2024", 2024.5),                       # bare year -> mid-year
    (2024, 2024.5),
    ("2024 Fall", 2024.70),
    ("Fall 2024", 2024.70),
    ("2025 Spring", 2025.12),
    ("summer 2023", 2023.45),
    ("2024-03", 2024 + 2 / 12),
    ("2024/03", 2024 + 2 / 12),
    ("2024-03-15", 2024 + 2 / 12),
    (datetime.date(2024, 3, 15), 2024 + 2 / 12),
    (datetime.datetime(2022, 1, 2), 2022.0),
])
def test_to_decimal_year(value, expected):
    got = gd.to_decimal_year(value)
    if expected is None:
        assert got is None
    else:
        assert got == pytest.approx(expected, abs=1e-6)


# ---------------------------------------------------------------- projects
def test_parse_projects_leads_and_separators():
    got = gd.parse_projects("Multijet BSM search (lead); Outer tracker, HSCP")
    assert got == [
        {"name": "Multijet BSM search", "lead": True},
        {"name": "Outer tracker", "lead": False},
        {"name": "HSCP", "lead": False},
    ]


def test_parse_projects_newlines_case_and_empty():
    assert gd.parse_projects("") == []
    got = gd.parse_projects("A (LEAD)\nB")
    assert got[0] == {"name": "A", "lead": True}
    assert got[1] == {"name": "B", "lead": False}


# ---------------------------------------------------------------- misc helpers
def test_add_stamp():
    assert gd.add_stamp("dashboard.html", "2026-07-30") == "dashboard_2026-07-30.html"
    assert gd.add_stamp("a/b/x.pdf", "2026-01-01") == "a/b/x_2026-01-01.pdf"


@pytest.mark.parametrize("value,expected", [
    ("$50,000", 50000.0),
    ("$1,200,000", 1200000.0),
    ("", 0.0),
    ("N/A", 0.0),
    (None, 0.0),
    ("750000", 750000.0),
])
def test_parse_money(value, expected):
    assert gd.parse_money(value) == expected


def test_clean_normalises_numbers_and_none():
    assert gd._clean(800.0) == "800"
    assert gd._clean(3.5) == "3.5"
    assert gd._clean(None) == ""
    assert gd._clean("  x  ") == "x"


def test_theme_key():
    assert gd._theme_key("Collider (CMS)") == "collider-cms"
    assert gd._theme_key("") == "other"


# ---------------------------------------------------------------- rank logic
@pytest.mark.parametrize("rank,student,postdoc,faculty", [
    ("Grad student", True, False, False),
    ("Undergrad", True, False, False),
    ("Postdoc", False, True, False),
    ("Faculty", False, False, True),
    ("Research staff", False, False, True),
    ("Professor", False, False, True),
    ("", False, False, False),
])
def test_rank_predicates(rank, student, postdoc, faculty):
    p = {"rank": rank}
    assert gd.is_student(p) is student
    assert gd.is_postdoc(p) is postdoc
    assert gd.is_faculty(p) is faculty


# ---------------------------------------------------------------- rendering
def _minimal_person(name="A", theme="T"):
    return {
        "name": name, "rank": "Grad student", "theme": theme, "site": "X",
        "funding": "", "program_year": "Y1", "effort": "", "file": "a.xlsx",
        "projects": [], "start_year": 2024.5, "group_year": 2024.5,
        "milestones": {}, "presentations": [], "publications": [],
        "applications": [], "awards": [], "roles": [], "outreach": [],
        "events": [], "grants": [],
        "committee_members": [], "committee_meetings": [], "progress_reports": [],
    }


def test_build_model_themes_and_kpis():
    model = gd.build_model([_minimal_person("A", "T1"), _minimal_person("B", "T2")],
                           "Title")
    assert model["group"]["title"] == "Title"
    assert [t["label"] for t in model["themes"]] == ["T1", "T2"]
    # distinct colors per theme
    assert len({t["color"] for t in model["themes"]}) == 2
    assert model["kpis"]["people"] == 2
    assert model["kpis"]["students"] == 2
    assert model["kpis"]["faculty"] == 0


def test_render_html_embeds_data_and_escapes_script():
    person = _minimal_person()
    person["funding"] = "money for </script> research"
    model = gd.build_model([person], "T")
    html = gd.render_html(model)
    assert "/*__DATA__*/null" not in html
    # the raw closing tag must not appear inside the data payload
    payload = html.split("const DATA = ", 1)[1].split(";\n", 1)[0]
    assert "</script>" not in payload
    assert json.loads(payload.replace("<\\/", "</"))["kpis"]["people"] == 1
