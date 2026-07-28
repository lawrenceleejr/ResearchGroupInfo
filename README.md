# ResearchGroupInfo

A lightweight system for keeping track of a research group. Each group member
maintains one spreadsheet (a "Research Record") describing their timeline,
milestones, projects, publications, presentations, applications, awards, and
service. A Python script reads a whole directory of these spreadsheets and
renders a single, self-contained, **interactive static HTML dashboard**.

## How it works

1. **The template** — [`template/FirstnameLastname_Info_Template.xlsx`](template/FirstnameLastname_Info_Template.xlsx)
   is the master form. Upload it to Google Sheets (or hand out copies), rename
   it `FirstnameLastname_Info.xlsx`, and ask each member to keep it current.
   It has nine tabs:

   | Tab | Contents |
   |-----|----------|
   | **Overview** | Identity & timeline, milestones/exams, thesis topic & projects |
   | **Committee** | Committee members, meeting log, annual progress reports |
   | **Roles** | Teaching, service, collaboration roles, mentoring |
   | **Presentations** | Talks, posters, seminars |
   | **Outreach & education** | Public engagement, teaching, mentoring |
   | **Publications** | Papers and notes |
   | **Applications** | Fellowships, grants, travel, programs (requested vs. awarded) |
   | **Awards** | Prizes and honors |
   | **Notes** | Free-form |

   Several fields use dropdowns (rank, research theme, milestone status,
   application status, …) so the data stays consistent across the group.

2. **Collect the records** — download everyone's sheet back to `.xlsx`
   (Google Sheets: *File → Download → Microsoft Excel*) into one directory.

3. **Generate the dashboard**:

   ```bash
   pip install -r requirements.txt
   python generate_dashboard.py <input_dir> -o dashboard.html --title "My Group — Dashboard"
   ```

   Open `dashboard.html` in any browser. Everything is embedded — no server,
   no external dependencies, no internet needed.

## The dashboard

Modeled on the reference design, the dashboard renders (each section is hidden
automatically when there's no data for it):

- **KPI strip** — people, students, postdocs, themes, sites, quals/comps
  passed, publications, presentations, awards, funding awarded.
- **Theme filter chips** — narrow every section to one research theme.
- **Student cohort timeline** — program start → present, with the 5–6 yr
  completion band.
- **Postdoc tenure** — years since arrival, with a ~3 yr career-stage band.
- **People ledger** — sortable table (click any column header).
- **Milestones & exams** — qual / comp / committee / author / defense status.
- **Who works on what** — project matrix (large square = project lead).
- **Applications & funding pipeline** — kanban board grouped by status.
- **Publications**, **Presentations & awards**, **Roles & service**,
  **Outreach & education** — tables aggregated across the group.
- Light/dark theme toggle (respects the OS preference by default).

Research themes and their colors are assigned automatically from whatever
themes appear in the data, so the dashboard adapts to any group.

## Summary CSVs

`generate_summaries.py` writes group-wide summary CSVs from the same inputs:

```bash
python generate_summaries.py <input_dir> -o summaries/
```

| File | Contents |
|------|----------|
| `presentations.csv` | Every talk / poster / seminar, one row each |
| `publications.csv` | Every paper / note, one row each |
| `conferences.csv` | Unique conferences attended, with attendees, years, and presentation counts |
| `schools.csv` | Unique schools / lecture courses attended, with attendees |

Conferences and schools are **derived from the Presentations tab** (the template
has no separate attendance field): each presentation implies attendance at its
venue. Entries typed `Lecture` — or whose venue name contains *school* /
*academy* / *tutorial* — are counted as schools; everything else is a
conference. Venues are grouped by series (a trailing year is stripped), so
`USMCC 2025` and `USMCC 2026` collapse into one `USMCC` row spanning both years.

A pre-built demo lives in [`summaries/`](summaries/).

## Trying it out

Sample records and a pre-built demo are included:

```bash
python make_sample_data.py                       # writes sample_data/*.xlsx
python generate_dashboard.py sample_data -o dashboard.html
```

The committed [`dashboard.html`](dashboard.html) is that demo output.

## Files

```
template/FirstnameLastname_Info_Template.xlsx   the master spreadsheet
generate_dashboard.py                           reads a dir of records → HTML
generate_summaries.py                           reads a dir of records → summary CSVs
make_sample_data.py                             writes illustrative sample records
sample_data/                                    generated sample records
dashboard.html                                  demo output (from sample_data)
summaries/                                       demo CSV output (from sample_data)
requirements.txt                                openpyxl
```

## Notes on parsing

The parser locates fields by their labels (not fixed cell coordinates), so it
tolerates inserted rows and only needs the tab/label names to stay intact.
Dates written as `2024`, `2024-03`, or `2024 Fall` are all understood for the
timelines. Example ("ex") rows and blank rows are skipped automatically.
