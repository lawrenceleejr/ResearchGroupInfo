# Research Group Info — a living dashboard for your group

A lightweight, **free, no-server** way for any faculty member to keep a live
picture of their research group. Every member (students, postdocs, and the
faculty/PI) maintains one simple spreadsheet — their **Research Record** — and a
Python script turns the whole folder into:

- a single **interactive HTML dashboard** (one file, opens in any browser), and
- a set of **summary CSVs** (presentations, publications, conferences, schools,
  grants) you can drop into reports, reviews, or renewals.

No database, no accounts, no hosting. Spreadsheets your group already knows how
to use, in → a shareable snapshot of everything the group is doing, out.

> **Faculty:** this is built to be a general tool. Fork or clone this repo,
> hand the template to your group, and you have a group dashboard in an
> afternoon. Nothing here is specific to one group.

---

## The workflow

1. **Put the template on a shared Google Drive.** Upload
   [`template/FirstnameLastname_Info_Template.xlsx`](template/FirstnameLastname_Info_Template.xlsx)
   into a Google Drive folder shared with your whole group (open it in Google
   Sheets, or keep it as `.xlsx` — both work).

2. **Each member keeps their own copy up to date.** Everyone makes a copy named
   for themselves — `JaneDoe_Info` — and keeps it current: new talks, papers,
   milestones, grants, conferences, and schools as they happen. The
   dropdown menus keep everyone's entries consistent. This is a *living*
   database: members update it whenever something changes, not once a year.

3. **Download the folder and run the script for a snapshot.** Whenever you want
   a current picture (group meeting, annual review, proposal, site visit),
   download the Drive folder of `.xlsx` files and run:

   ```bash
   pip install -r requirements.txt
   python generate_dashboard.py <folder> -o dashboard.html --title "My Group"
   python generate_summaries.py <folder> -o summaries/
   ```

   Open `dashboard.html` in any browser. Everything is embedded — no server, no
   internet needed. Re-run any time for a fresh snapshot.

> Google Sheets export: *File → Download → Microsoft Excel (.xlsx)*. Downloading
> the whole Drive folder as a zip and unzipping it gives you the input folder in
> one step.

---

## The template (11 tabs)

| Tab | What it captures |
|-----|------------------|
| **Overview** | Identity & timeline (incl. **rank**: undergrad / grad / postdoc / **faculty** / research staff), milestones & exams, thesis topic & projects |
| **Committee** | Thesis committee members, meeting log, annual progress reports |
| **Roles** | Teaching, service, collaboration roles, mentoring |
| **Presentations** | Talks, posters, seminars |
| **Conferences & Schools** | Events **attended** — conferences, workshops, summer/winter schools (whether or not you presented) |
| **Outreach & education** | Public engagement, teaching, mentoring |
| **Publications** | Papers and notes |
| **Applications** | Individual fellowships, grants, travel, programs (requested vs. awarded) |
| **Grants & Funding** | The group's **funded** grant portfolio — usually the faculty/PI's grants (PI / Co-PI / senior personnel, amounts, status) |
| **Awards** | Prizes and honors |
| **Notes** | Free-form |

Dropdowns (rank, research theme, milestone status, application/grant status,
event type, …) keep the data consistent so the script can aggregate it cleanly.

Regenerate the template yourself with `python build_template.py` (it adds the
faculty option and the two newest tabs to the base workbook).

---

## The dashboard

Modeled on a clean, Tufte-inspired design; every section hides itself when
there's no data for it:

- **KPI strip** — people, students, postdocs, **faculty/staff**, themes, sites,
  quals/comps passed, publications, presentations, awards, **active grants**,
  and **total funding on the books**.
- **Theme filter chips** — narrow every section to one research theme.
- **Student cohort timeline** and **postdoc tenure** — with completion /
  career-stage bands. (Faculty appear in the roster, not the timelines.)
- **People ledger** — sortable table.
- **Milestones & exams** — qual / comp / committee / author / defense.
- **Who works on what** — project matrix (large square = project lead).
- **Grants & funding** — the funded portfolio, by status.
- **Applications & fellowship pipeline** — individual applications, by status.
- **Publications**, **Presentations & awards**, **Conferences & schools**,
  **Roles & service**, **Outreach & education** — aggregated across the group.
- Light/dark theme toggle (follows the OS preference by default).

Research themes and their colors are assigned automatically from whatever
appears in the data, so the dashboard adapts to any group.

## The summary CSVs

`generate_summaries.py` writes group-wide CSVs from the same folder:

| File | Contents |
|------|----------|
| `presentations.csv` | Every talk / poster / seminar, one row each |
| `publications.csv` | Every paper / note, one row each |
| `conferences.csv` | Unique conferences attended — attendees, years, locations, counts |
| `schools.csv` | Unique summer/winter schools attended — attendees, years, locations |
| `grants.csv` | The funded grant portfolio — holder, agency, role, amount, period, status |

Conferences and schools are drawn from **both** the *Conferences & Schools*
attendance tab and the *Presentations* tab (each presentation implies attending
its venue), then merged: entries typed/named like a *school* / *summer school* /
*academy* / *tutorial* become schools, the rest conferences. Events are grouped
by series (a trailing year is stripped), so `USMCC 2025` and `USMCC 2026`
collapse into one `USMCC` row, and the same person+event+year is counted once.

---

## Try it now

Sample records and pre-built output are included:

```bash
python make_sample_data.py                                   # writes sample_data/*.xlsx
python generate_dashboard.py sample_data -o dashboard.html   # -> dashboard.html
python generate_summaries.py sample_data -o summaries/       # -> summaries/*.csv
```

The committed [`dashboard.html`](dashboard.html) and [`summaries/`](summaries/)
are that demo output.

## Files

```
template/FirstnameLastname_Info_Template.xlsx   the master spreadsheet (11 tabs)
build_template.py                               (re)builds the template
generate_dashboard.py                           a folder of records -> HTML dashboard
generate_summaries.py                           a folder of records -> summary CSVs
make_sample_data.py                             writes illustrative sample records
sample_data/                                    generated sample records
dashboard.html                                  demo dashboard (from sample_data)
summaries/                                      demo CSVs (from sample_data)
requirements.txt                                openpyxl
```

## Notes

- The only dependency is `openpyxl`; the generated dashboard has **zero**
  dependencies and works offline by double-clicking.
- The parser finds fields by their **labels**, not fixed cell coordinates, so
  inserted rows are fine — only the tab and label names need to stay intact.
- Dates written as `2024`, `2024-03`, or `2024 Fall` are all understood.
- Example (`ex`) rows and blank rows are skipped automatically.
