# Run it inside Google Drive — no command line (recommended)

This is the **recommended way to run the tool**. It's a **Google Apps Script**
that runs on Google's servers against your Drive folder, so nobody in the group
ever installs or runs anything locally. Members just keep their spreadsheets
current; the dashboard regenerates itself.

It reads every Google Sheet in a folder, parses the 11 tabs, and writes a
**timestamped** `dashboard_YYYY-MM-DD_HHMM.html` back into Drive — older
snapshots are kept, so a history builds up. It produces the same dashboard as
the command-line version (verified against the same data).

## What's here

| File | Role |
|------|------|
| `Code.gs` | the logic (parsing, aggregation, rendering) |
| `dashboard_template.html` | the HTML/CSS/JS shell (generated from the Python template via `python generate_dashboard.py --dump-template`) |

## Setup (about 5 minutes, once)

1. **Collect the records in one Drive folder** as *Google Sheets*. When members
   upload an `.xlsx`, open it once in Google Sheets (or turn on Drive Settings →
   *Convert uploaded files to Google Docs editor format*) so it becomes a Sheet.

2. **Create the script.** Go to [script.google.com](https://script.google.com) →
   **New project**. Then:
   - Paste `Code.gs` over the default `Code.gs`.
   - Add an **HTML file** (＋ → HTML) named exactly **`dashboard_template`**, and
     paste in the contents of `dashboard_template.html`.

3. **Point it at your folder.** In `Code.gs`, set `FOLDER_ID` to your Drive
   folder's id — the long string in the folder URL
   (`drive.google.com/drive/folders/`**`THIS_PART`**). Optionally set `TITLE`.

4. **Run once.** Select `generateDashboard` in the toolbar and click **Run**.
   Approve the permission prompt (it needs access to your Drive). A timestamped
   `dashboard_YYYY-MM-DD_HHMM.html` appears in the folder.

## How you'll view it

**Google Drive does not render HTML pages** — double-clicking the generated file
just offers a download, not the running dashboard. So set up the web app, which
is the everyday way to look at the dashboard:

- **Live URL (web app) — do this.** **Deploy → New deployment → Web app**,
  execute as *you*, access *anyone in your org* (or *anyone with the link*). You
  get a URL that renders the current dashboard on every visit — bookmark it and
  share it with the group. Nothing to download, always up to date.

- **A past snapshot.** To look at an archived day, download that timestamped
  `.html` from Drive (right-click → **Download**) and open it in a browser. Each
  file is self-contained and opens with no internet.

## Keep it fresh automatically — pick any

- **Daily trigger.** Run `installDailyTrigger` once. A fresh timestamped
  snapshot is written every night (~4am).

- **A button in a Sheet.** If you bind the script to a Google Sheet (Extensions →
  Apps Script from a Sheet), a **Group Dashboard → Regenerate now** menu appears.

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
