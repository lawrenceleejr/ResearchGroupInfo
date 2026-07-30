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

## Setup (about 5 minutes, once — you only do this)

You need to do this **once**. Group members never see any of it.

1. **Make one Drive folder** and put the members' records in it as *Google
   Sheets*. (If someone uploads an `.xlsx`, open it and do **File → Save as
   Google Sheets** so it counts.) Put one more empty Google Sheet in the folder
   and call it e.g. **"Dashboard control".**

2. **Open the script editor.** In that "Dashboard control" sheet, click
   **Extensions → Apps Script**. A code editor opens in a new tab.

3. **Paste in the two files.**
   - Select everything in the `Code.gs` panel and replace it with the contents
     of this folder's `Code.gs`.
   - Click the **＋** next to *Files* → **HTML**, name it exactly
     **`dashboard_template`**, and paste in the contents of
     `dashboard_template.html`. Click the save icon.

4. **Run it once.** Pick `generateAll` in the toolbar dropdown and click
   **▶ Run**. Approve the Google permission prompt. That's it — the Doc report,
   the Group Summaries sheet, and a timestamped dashboard HTML appear in your
   folder.

Because the script lives in a Sheet **inside** the folder, it finds the folder
on its own — there's nothing to configure. (If you'd rather keep the script
somewhere else, paste your folder's id into `FOLDER_ID` at the top of `Code.gs`;
it's the long string in the folder's URL after `/folders/`.)

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
