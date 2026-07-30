# Research Group Info — a living dashboard for your group

A lightweight, **free, no-server** way for any faculty member to keep a live
picture of their research group. Every member (students, postdocs, and the
faculty/PI) maintains one simple spreadsheet — their **Research Record** — and
the tool turns the whole folder into:

- a single **interactive HTML dashboard** — timelines, milestones, grants,
  publications, presentations, and more, and
- a set of **summary CSVs** (presentations, publications, conferences, schools,
  grants) you can drop into reports, reviews, or renewals.

No database, no accounts, no software to install. The spreadsheets your group
already knows how to use, in → a shareable snapshot of everything the group is
doing, out.

> **Faculty:** this is built to be a general tool. Point it at your group's
> Drive folder and you have a group dashboard in an afternoon. Nothing here is
> specific to one group.

There are two ways to run it. **Most people should use the first one.**

---

## ⭐ Recommended: run it inside Google Drive (no command line)

This is the easiest way and needs **nothing installed** — it runs on Google's
own servers against your Drive folder. Set it up once (about 5 minutes) and the
dashboard rebuilds itself on a schedule or at the click of a button.

1. **Put the template on a shared Google Drive.** Upload
   [`template/FirstnameLastname_Info_Template.xlsx`](template/FirstnameLastname_Info_Template.xlsx)
   into a Drive folder shared with your group, and **open it in Google Sheets**
   (so it becomes a Google Sheet).

2. **Each member keeps their own copy up to date.** This is all a group member
   ever has to do — no code, no downloads, no setup:

   > **For group members (copy/paste this to your group):**
   > 1. Open the template in the shared Drive folder.
   > 2. **File → Make a copy**, into the same folder.
   > 3. Rename the copy to your name, e.g. `JaneDoe_Info`.
   > 4. Fill in the **Overview** tab — that's the only required one. Add talks,
   >    papers, grants, conferences, etc. on the other tabs as you have them.
   > 5. Keep it updated whenever something changes. That's it.

   The dropdown menus keep everyone's entries consistent. This is a *living*
   database: members update it whenever something changes, not once a year.

3. **Set up the Apps Script once.** Follow
   [`apps_script/README.md`](apps_script/README.md) — in a Sheet inside the
   folder you open **Extensions → Apps Script**, paste in two files, and click
   **Run**. No IDs to configure (it finds the folder on its own). Each run
   writes **three outputs** into your Drive folder:
   - a **Google Doc report** (`group_report_2026-07-30_1430`) — opens right in
     Drive with a double-click,
   - a **Group Summaries** Google Sheet (presentations, publications,
     conferences, schools, grants tables) — also opens right in Drive, and
   - a **timestamped interactive HTML dashboard**
     (`dashboard_2026-07-30_1430.html`).

   It can run on a **daily schedule**, from a **menu button** in that Sheet,
   and/or serve a **live web-app URL** the whole group can bookmark.
   Timestamped outputs are kept, so a history builds up automatically.

Members never touch anything but their spreadsheet. Only the one-time setup
above is done by you (or anyone comfortable clicking through a Google prompt).

### How do I actually *see* the outputs?

- **The Doc report and the Group Summaries sheet open right in Drive** — just
  double-click them, like any Google Doc or Sheet. For most everyday questions
  ("what has the group presented this year?", "what grants are active?") these
  are all you need, and they're the easiest things to share.

- **The interactive HTML dashboard** is the richest view (filters, timelines,
  sortable tables), but **Google Drive does not render HTML pages** —
  double-clicking only offers a download. Two ways to view it properly:
  - **Best — the live web-app URL.** Deploying the Apps Script as a **web app**
    (one step, covered in `apps_script/README.md`) gives you a link that renders
    the current dashboard right in the browser, always up to date. Bookmark it
    and share it with the group.
  - **For a past snapshot** — download that timestamped `.html` from Drive
    (right-click → **Download**) and double-click it; each file is completely
    self-contained and opens in any browser with no internet.

So: **double-click the Doc/Sheet for a quick look, use the web-app URL for the
interactive dashboard, and the timestamped files are your archive.**

---

## Advanced: run it from the command line

If you're comfortable in a terminal (or want to script it), you can run the
Python version on any machine with Python. This is optional — the Google Drive
method above produces the same dashboard.

```bash
pip install -r requirements.txt
python generate_dashboard.py <folder> -o dashboard.html --title "My Group"
python generate_summaries.py <folder> -o summaries/
```

`<folder>` is a directory of the members' records as `.xlsx` files (download the
Drive folder: select it → **Download**, then unzip). Outputs are **date-stamped**
by default — `dashboard_2026-07-30.html`, `presentations_2026-07-30.csv` — so
re-running builds a history. (Add `--no-datestamp` for stable filenames.) The
HTML is self-contained; just open it in a browser.

**PDF.** Add `--pdf` to also export a PDF (great for reports and reviews):

```bash
python generate_dashboard.py <folder> --pdf
```

It prints through a headless browser so the PDF matches the page. One-time
setup: `pip install playwright && playwright install chromium`. No Python handy?
Open the HTML and use your browser's **Print → Save as PDF** — a print
stylesheet is included so it paginates cleanly.

---

## The template (11 tabs)

Only **Overview** is required; every other tab is an optional log you fill in as
you have things to add.

| Tab | What it captures |
|-----|------------------|
| **Overview** *(required)* | Identity & timeline (incl. **rank**: undergrad / grad / postdoc / **faculty** / research staff), milestones & exams, thesis topic & projects |
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
event type, …) keep the data consistent so the tool can aggregate it cleanly.

Regenerate the template yourself with `python build_template.py` (it adds the
faculty option and the two newest tabs to the base workbook).

---

## The dashboard

A clean, uncluttered layout; every section hides itself when there's no data
for it:

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

## The summary tables

Five group-wide tables — in Drive mode they're the tabs of the **Group
Summaries** Google Sheet; at the command line, `generate_summaries.py` writes
them as CSVs:

| Table | Contents |
|-------|----------|
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

## Try it now (command line)

Sample records and pre-built output are included:

```bash
python make_sample_data.py                                   # writes sample_data/*.xlsx
python generate_dashboard.py sample_data --pdf               # -> dashboard_<date>.html + .pdf
python generate_summaries.py sample_data -o summaries/       # -> summaries/*_<date>.csv
```

The committed [`dashboard.html`](dashboard.html), [`dashboard.pdf`](dashboard.pdf),
and [`summaries/`](summaries/) are that demo output, generated with
`--no-datestamp` so the links here stay stable.

## Files

```
template/FirstnameLastname_Info_Template.xlsx   the master spreadsheet (11 tabs)
apps_script/                                    ⭐ Google Apps Script (no command line)
build_template.py                               (re)builds the template
generate_dashboard.py                           a folder of records -> HTML dashboard (+ optional PDF)
generate_summaries.py                           a folder of records -> summary CSVs
make_sample_data.py                             writes illustrative sample records
sample_data/                                    generated sample records
dashboard.html                                  demo dashboard (from sample_data)
summaries/                                      demo CSVs (from sample_data)
requirements.txt                                openpyxl (playwright optional, for --pdf)
```

## Notes

- The Google Drive method needs nothing installed. For the command-line method
  the only required dependency is `openpyxl`; the generated dashboard has
  **zero** dependencies and works offline by double-clicking. PDF export needs
  `playwright` (optional).
- The parser finds fields by their **labels**, not fixed cell coordinates, so
  inserted rows are fine — only the tab and label names need to stay intact.
- Dates written as `2024`, `2024-03`, or `2024 Fall` are all understood.
- Example (`ex`) rows and blank rows are skipped automatically.
