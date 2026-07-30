# Run it inside Google Drive — no command line (recommended)

This is the **recommended way to run the tool**. It's a **Google Apps Script**
that runs on Google's servers against your Drive folder, so nobody in the group
ever installs or runs anything locally. Members just keep their spreadsheets
current; the dashboard regenerates itself.

It reads every Google Sheet in a folder, parses the 11 tabs, and writes **three
outputs** back into Drive:

| Output | What it is | Opens directly in Drive? |
|--------|-----------|--------------------------|
| `group_report_YYYY-MM-DD_HHMM` | **Google Doc report** — key numbers, people, milestones, grants, publications, presentations, conferences/schools, awards, roles, outreach | ✅ double-click |
| **Group Summaries** | **Google Sheet** with the five summary tables (presentations, publications, conferences, schools, grants), updated in place | ✅ double-click |
| `dashboard_YYYY-MM-DD_HHMM.html` | the **interactive dashboard** snapshot | ❌ view via web-app URL or download |

Timestamped outputs are kept, so a history builds up automatically. Everything
matches the command-line version (verified against the same data).

## What's here

| File | Role |
|------|------|
| `Code.gs` | the logic (parsing, aggregation, rendering) |
| `dashboard_template.html` | the HTML/CSS/JS shell (generated from the Python template via `python generate_dashboard.py --dump-template`) |

## Setup

The complete click-by-click setup lives in the
[main README](../README.md#-recommended-run-it-inside-google-drive-no-command-line)
— in short: create a "Dashboard control" Sheet inside the records folder, open
**Extensions → Apps Script**, paste in `Code.gs` and `dashboard_template.html`
(as an HTML file named exactly `dashboard_template`), and **▶ Run**
`generateAll` once.

One extra note for non-standard setups: the script finds the folder on its own
because it lives in a Sheet **inside** it. If you'd rather keep the script
somewhere else, paste your folder's id into `FOLDER_ID` at the top of `Code.gs`
(the long string in the folder's URL after `/folders/`). And if a member's
record is an `.xlsx` upload rather than a Google Sheet, open it and do
**File → Save as Google Sheets** so it counts.

## How you'll view it

- **The Doc report and the Group Summaries sheet open right in Drive** — just
  double-click them. For most everyday "how's the group doing" questions, the
  Doc report is the easiest thing to look at and share.

- **The interactive HTML dashboard** is the richest view (filters, timelines,
  sortable tables), but **Google Drive does not render HTML pages** —
  double-clicking only offers a download. Two ways to see it properly:
  - **Live URL (web app) — recommended.** **Deploy → New deployment → Web
    app**, execute as *you*, access *anyone in your org* (or *anyone with the
    link*). You get a URL that renders the current dashboard on every visit —
    bookmark it and share it with the group.
  - **A past snapshot.** Download a timestamped `.html` from Drive
    (right-click → **Download**) and open it in a browser. Each file is
    self-contained and opens with no internet.

## Keep it fresh automatically — pick any

- **Daily trigger.** Run `installDailyTrigger` once. All three outputs refresh
  every night (~4am).

- **A menu in your control Sheet.** Reload the "Dashboard control" sheet and a
  **Group Dashboard** menu appears — *Regenerate everything*, or any single
  output, any time.

The web-app URL always shows the latest, so you don't need to regenerate files
just to view current data — the files are your dated archive.

## Updating the template later

When you improve the template (a new tab, a new dropdown option), members'
already-filled records don't break — the parser reads by tab/label names and
treats every tab except Overview as optional. To actually push the new pieces
into everyone's file, keep the template Sheet in the folder (its name must
contain "template") and click **Group Dashboard → Update member records to
latest template**. For each member Sheet it:

- copies in any **missing tabs** (formatting and dropdowns included), and
- **refreshes the dropdown lists** on tabs they already have,

and never touches anything a member has written. It's safe to run repeatedly.
The corresponding template policy: **only add tabs/labels/options — never
rename, remove, or reorder them.**

## PDF

The dashboard draws itself with JavaScript, so a faithful PDF needs a real
browser to run that code:

- **Zero-install:** open the web-app URL (or a downloaded snapshot) and use the
  browser's **Print → Save as PDF**. A print stylesheet is included, so it
  paginates cleanly and forces a light theme.
- **Automated:** run `python generate_dashboard.py <folder> --pdf` on any
  machine with Python (see the top-level README). Good for a laptop or a
  scheduled job.

Apps Script's own `Blob.getAs(PDF)` can't run the page's JavaScript, so it would
export a blank page — that's why PDF stays a browser/print step rather than a
server-side one.

## Keeping the template in sync

`dashboard_template.html` is generated from the single source of truth in
`generate_dashboard.py`. If you change the dashboard's look there, regenerate it
with:

```bash
python generate_dashboard.py --dump-template apps_script/dashboard_template.html
```

and paste the new contents into the `dashboard_template` HTML file in your
Apps Script project.
