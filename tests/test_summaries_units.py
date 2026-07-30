"""Unit tests for the aggregation helpers in generate_summaries.py."""
import pytest

import generate_summaries as gs


def test_year_extraction():
    assert gs._year("2025-06") == "2025"
    assert gs._year("June 2025") == "2025"
    assert gs._year("") == ""
    assert gs._year(None) == ""


@pytest.mark.parametrize("venue,series", [
    ("USMCC 2025", "USMCC"),
    ("USMCC", "USMCC"),
    ("IMCC 2026", "IMCC"),
    ("DPF 2026,", "DPF"),
    ("2025", "2025"),          # nothing left after stripping -> keep original
])
def test_venue_series(venue, series):
    assert gs._venue_series(venue) == series


def test_sorted_join_dedup_and_blanks():
    assert gs._sorted_join(["b", "a", "b", "", "a"]) == "a; b"
    assert gs._sorted_join([]) == ""


def test_is_school_presentation():
    assert gs._is_school({"type": "Lecture", "venue": "anything"}) is True
    assert gs._is_school({"type": "Poster", "venue": "US Particle Accelerator School"}) is True
    assert gs._is_school({"type": "Poster", "venue": "APS April Meeting"}) is False


def test_event_is_school_uses_type_and_name():
    assert gs._event_is_school("CMSDAS", "Summer school") is True
    assert gs._event_is_school("CERN Academy 2025", "Conference") is True
    assert gs._event_is_school("CMS Week", "Collaboration meeting") is False


def _people():
    return [
        {"name": "Alice",
         "events": [
             {"event": "USMCC 2025", "type": "Conference", "location": "Chicago, IL",
              "start": "2025-06"},
             {"event": "CMSDAS", "type": "Summer school", "location": "FNAL",
              "start": "2023-01"},
         ],
         "presentations": [
             # same series + same year as the event above -> must dedup
             {"venue": "USMCC 2025", "type": "Poster", "location": "Chicago, IL",
              "date": "2025-06"},
             # same series, different year -> counted
             {"venue": "USMCC 2026", "type": "Contributed talk",
              "location": "Austin, TX", "date": "2026-06"},
         ]},
        {"name": "Bob",
         "events": [],
         "presentations": [
             {"venue": "USMCC 2025", "type": "Contributed talk",
              "location": "Chicago, IL", "date": "2025-06"},
         ]},
    ]


def test_attendance_merge_and_dedup():
    records = gs.collect_attendance(_people())
    conf = gs._aggregate(records, want_schools=False)
    assert set(conf.keys()) == {"usmcc"}
    usmcc = conf["usmcc"]
    # Alice 2025 (deduped across event+presentation), Alice 2026, Bob 2025
    assert usmcc["count"] == 3
    assert sorted(set(usmcc["attendees"])) == ["Alice", "Bob"]
    assert sorted(set(usmcc["years"])) == ["2025", "2026"]

    schools = gs._aggregate(records, want_schools=True)
    assert set(schools.keys()) == {"cmsdas"}
    assert schools["cmsdas"]["count"] == 1


def test_out_path_stamping():
    saved = gs.STAMP
    try:
        gs.STAMP = ""
        assert gs.out_path("dir", "x.csv").endswith("x.csv")
        gs.STAMP = "2026-07-30"
        assert gs.out_path("dir", "x.csv").endswith("x_2026-07-30.csv")
    finally:
        gs.STAMP = saved
