/**
 * Research Group Info — Google Apps Script edition.
 *
 * Runs entirely on Google's servers against a Drive folder of member
 * "Research Record" spreadsheets. No command line, nothing to install.
 *
 * It reads every Google Sheet in a folder, parses the 11 tabs, and writes back
 * into the folder (older snapshots are kept, so a history builds up):
 *   - dashboard_YYYY-MM-DD_HHMM.html — the interactive dashboard snapshot
 *   - group_report_YYYY-MM-DD_HHMM  — a Google Doc report (opens right in Drive)
 *   - "Group Summaries"             — a Google Sheet of the five summary tables,
 *                                     updated in place (opens right in Drive)
 * It can run from a menu, on a daily time-driven trigger, or be published as a
 * web app so the whole group can bookmark a live URL.
 *
 * VIEWING: the Doc report and the summary Sheet open directly in Google Drive.
 * The interactive HTML does not (Drive doesn't render HTML pages) — view it via
 * the deployed web-app URL, or download a snapshot and open it in a browser.
 *
 * SETUP (see apps_script/README.md for the click-by-click version)
 *   1. Put the members' records in one Drive folder, each as a Google Sheet.
 *   2. In a Google Sheet inside that folder: Extensions → Apps Script.
 *   3. Paste this into Code.gs, and add an HTML file named exactly
 *      `dashboard_template` with the contents of dashboard_template.html.
 *   4. Run `generateDashboard` once and grant permissions. (The folder is
 *      detected automatically; or set FOLDER_ID below to be explicit.)
 *
 * OPTIONAL
 *   - `installDailyTrigger()` regenerates the dashboard every night.
 *   - Deploy → New deployment → Web app for a live, always-current URL.
 *
 * PDF: the dashboard draws itself with JavaScript, so a faithful PDF needs a
 *   browser. Open the web-app URL (or a downloaded snapshot) and use the
 *   browser's Print → Save as PDF, or run generate_dashboard.py --pdf.
 */

// Leave FOLDER_ID blank to auto-detect (the folder that holds this script's
// Sheet). Or paste your Drive folder's id — the long string in its URL:
//   drive.google.com/drive/folders/THIS_PART
var FOLDER_ID = '';
var OUTPUT_FOLDER_ID = '';            // blank = write into the same folder
var TITLE = 'Research Group — Dashboard';
var TEMPLATE_FILE = 'dashboard_template';   // name of the HTML file in this project

var PALETTE = ['#9c5b46', '#3f6079', '#5b7a5b', '#9c7b2e',
               '#7a5b8a', '#3f8a8a', '#a5563f', '#6b7a3f'];
var SCHOOL_TABS = ['Conferences & Schools', 'Conferences and Schools', 'Conferences'];
var GRANT_TABS = ['Grants & Funding', 'Grants and Funding', 'Grants'];

/* ============================ entry points ============================ */

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Group Dashboard')
      .addItem('Regenerate everything', 'generateAll')
      .addSeparator()
      .addItem('Dashboard (HTML snapshot)', 'generateDashboard')
      .addItem('Doc report (opens in Drive)', 'generateDocReport')
      .addItem('Summary Sheet (opens in Drive)', 'generateSummarySheet')
      .addSeparator()
      .addItem('Update member records to latest template', 'updateMemberRecords')
      .addToUi();
  } catch (e) { /* not a bound script */ }
}

/** One click: HTML snapshot + Doc report + summary Sheet. */
function generateAll() {
  var urls = {
    dashboard: generateDashboard(),
    report: generateDocReport(),
    summaries: generateSummarySheet(),
  };
  Logger.log('Done: %s', JSON.stringify(urls));
  return urls;
}

/* ===================== template updates (additive only) ===================== */

// Dropdowns and layout live in this top-left region of each tab; validation
// refresh is confined to it so nothing else is touched.
var DV_REGION_ROWS = 60;
var DV_REGION_COLS = 12;

/**
 * Upgrade every member record to the latest template, without touching data.
 *
 * Strictly additive: for each member Sheet in the folder it (a) copies over any
 * tab that exists in the template but is missing from the member's file
 * (formatting and dropdowns included), and (b) refreshes the dropdown rules on
 * tabs the member already has, so new options (e.g. a new rank or status)
 * appear. Data rows, filled cells, and anything the member wrote are never
 * modified, moved, or deleted — which is also why the template policy is
 * "only add tabs/labels/options; never rename or remove them".
 *
 * The master is the Google Sheet in the folder whose name contains "template".
 * Safe to run repeatedly; run it after you change the template.
 */
function updateMemberRecords() {
  var folderId = resolveFolderId_();
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var template = null, members = [];
  while (files.hasNext()) {
    var f = files.next();
    if (/template/i.test(f.getName())) { if (!template) template = f; continue; }
    if (/dashboard|summar|report/i.test(f.getName())) continue;
    members.push(f);
  }
  if (!template) {
    throw new Error('No template found: keep a Google Sheet whose name ' +
                    'contains "template" in the folder.');
  }
  members.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });

  var tmpl = SpreadsheetApp.openById(template.getId());
  var tmplSheets = tmpl.getSheets();
  var report = [];
  members.forEach(function (f) {
    var ss = SpreadsheetApp.openById(f.getId());
    var added = [], refreshed = 0;
    tmplSheets.forEach(function (ts) {
      var name = ts.getName();
      var target = ss.getSheetByName(name);
      if (!target) {
        ts.copyTo(ss).setName(name);
        added.push(name);
      } else {
        var rows = Math.min(DV_REGION_ROWS, ts.getMaxRows(), target.getMaxRows());
        var cols = Math.min(DV_REGION_COLS, ts.getMaxColumns(), target.getMaxColumns());
        if (rows < 1 || cols < 1) return;
        var rules = ts.getRange(1, 1, rows, cols).getDataValidations();
        target.getRange(1, 1, rows, cols).setDataValidations(rules);
        refreshed++;
      }
    });
    var line = f.getName() + ' — added: ' +
               (added.length ? added.join(', ') : 'none') +
               '; dropdowns refreshed on ' + refreshed + ' tab(s)';
    report.push(line);
    Logger.log(line);
  });
  Logger.log('Updated %s member record(s) from "%s".', members.length, template.getName());
  return report;
}

function generateDashboard() {
  var folderId = resolveFolderId_();
  var model = buildModel_(readPeople_(folderId), TITLE);
  var html = renderHtml_(model);
  var out = OUTPUT_FOLDER_ID ? DriveApp.getFolderById(OUTPUT_FOLDER_ID)
                             : DriveApp.getFolderById(folderId);
  // Timestamped filename (date + time) so every run is kept and a history
  // builds up in Drive. Only an exact same-minute name is replaced.
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  var name = 'dashboard_' + stamp + '.html';
  var existing = out.getFilesByName(name);
  while (existing.hasNext()) existing.next().setTrashed(true);
  var file = out.createFile(name, html, MimeType.HTML);
  Logger.log('Wrote %s (%s people) → %s',
             file.getName(), model.people.length, file.getUrl());
  return file.getUrl();
}

/** Live web-app endpoint: renders the current snapshot on each request. */
function doGet() {
  var model = buildModel_(readPeople_(resolveFolderId_()), TITLE);
  return HtmlService.createHtmlOutput(renderHtml_(model))
    .setTitle(TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Work out which Drive folder holds the records:
 *   1. the FOLDER_ID constant if you set one, else
 *   2. a value remembered from a previous run, else
 *   3. the folder that contains this script's Google Sheet (bound scripts).
 * The resolved id is remembered so daily triggers and the web app work too.
 */
function resolveFolderId_() {
  var props = PropertiesService.getScriptProperties();
  if (FOLDER_ID && FOLDER_ID.indexOf('PUT_YOUR') === -1) {
    props.setProperty('FOLDER_ID', FOLDER_ID);
    return FOLDER_ID;
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var parents = DriveApp.getFileById(ss.getId()).getParents();
      if (parents.hasNext()) {
        var id = parents.next().getId();
        props.setProperty('FOLDER_ID', id);
        return id;
      }
    }
  } catch (e) { /* not a bound script or no Drive access yet */ }
  var saved = props.getProperty('FOLDER_ID');
  if (saved) return saved;
  throw new Error(
    'Could not find the records folder. Either run this from a Google Sheet ' +
    'inside that folder, or set FOLDER_ID at the top of Code.gs to the folder id.');
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'generateDashboard' || h === 'generateAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('generateAll').timeBased().everyDays(1).atHour(4).create();
  Logger.log('Daily trigger installed (≈4am): HTML snapshot + Doc report + summary Sheet.');
}

/* ============================ helpers ============================ */

function cleanStr_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  if (typeof v === 'number') {
    return (v % 1 === 0) ? String(Math.round(v)) : String(v);
  }
  return String(v).trim();
}

var TERMS = {winter: 0.05, spring: 0.12, summer: 0.45, fall: 0.70, autumn: 0.70};

function toDecimalYear_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + v.getMonth() / 12.0;
  }
  var s = cleanStr_(v);
  var ym = s.match(/((?:19|20)\d{2})/);
  if (!ym) return null;
  var year = parseInt(ym[1], 10);
  var mm = s.match(/(?:19|20)\d{2}[-\/](\d{1,2})/);
  var frac = null;
  if (mm) { var m = parseInt(mm[1], 10); if (m >= 1 && m <= 12) frac = (m - 1) / 12.0; }
  if (frac === null) {
    var low = s.toLowerCase();
    for (var t in TERMS) { if (low.indexOf(t) !== -1) { frac = TERMS[t]; break; } }
  }
  return year + (frac === null ? 0.5 : frac);
}

function parseMoney_(v) {
  var s = cleanStr_(v).replace(/[^0-9.]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function findLabelRow_(values, text, col) {
  col = col || 0;
  var target = String(text).trim().toLowerCase();
  for (var i = 0; i < values.length; i++) {
    if (cleanStr_(values[i][col]).toLowerCase() === target) return i;
  }
  return -1;
}

function readTable_(values, headerRow, ncols, stopLabels) {
  var stops = {};
  (stopLabels || []).forEach(function (s) { stops[s.toLowerCase()] = true; });
  var out = [];
  for (var r = headerRow + 1; r < values.length; r++) {
    var first = cleanStr_(values[r][0]).toLowerCase();
    if (stops[first]) break;
    var vals = [];
    for (var c = 0; c < ncols; c++) vals.push(cleanStr_(values[r][c]));
    if (first === 'ex') continue;
    var empty = true;
    for (var k = 1; k < vals.length; k++) if (vals[k] !== '') { empty = false; break; }
    if (empty) continue;
    out.push(vals);
  }
  return out;
}

function parseProjects_(raw) {
  var out = [];
  if (!raw) return out;
  raw.split(/[\n;,]+/).forEach(function (part) {
    var name = part.trim();
    if (!name) return;
    var lead = false;
    var m = name.match(/\(\s*lead\s*\)/i);
    if (m) { lead = true; name = (name.slice(0, m.index) + name.slice(m.index + m[0].length)).trim(); }
    name = name.replace(/^[\s\-–·]+|[\s\-–·]+$/g, '');
    if (name) out.push({name: name, lead: lead});
  });
  return out;
}

function themeKey_(theme) {
  var k = String(theme || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  return k || 'other';
}

function sheetByNames_(ss, names) {
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}

function values_(sheet) { return sheet ? sheet.getDataRange().getValues() : null; }

/* ============================ parsing ============================ */

function readPeople_(folderId) {
  var folder = DriveApp.getFolderById(folderId || resolveFolderId_());
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var names = [];
  while (files.hasNext()) names.push(files.next());
  names.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  var people = [];
  names.forEach(function (f) {
    // Skip the tool's own outputs and the blank template sitting in the folder.
    if (/dashboard|template|summar|report/i.test(f.getName())) return;
    try {
      people.push(parseSpreadsheet_(SpreadsheetApp.openById(f.getId()), f.getName()));
    } catch (e) {
      Logger.log('SKIP %s: %s', f.getName(), e);
    }
  });
  // Nudge if members left records as un-converted Excel uploads (skipped above).
  var xlsx = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  var nx = 0;
  while (xlsx.hasNext()) { xlsx.next(); nx++; }
  if (nx) Logger.log('Note: %s Excel (.xlsx) file(s) ignored — open each in ' +
                     'Google Sheets (File → Save as Google Sheets) so it is counted.', nx);
  if (!people.length) throw new Error(
    'No Google Sheet records found in the folder. Make sure each member\'s file ' +
    'is a Google Sheet (not an .xlsx upload).');
  return people;
}

function parseSpreadsheet_(ss, fileName) {
  var p = {file: fileName};

  var ov = ss.getSheetByName('Overview');
  if (ov) {
    var v = values_(ov);
    var get = function (label) {
      var r = findLabelRow_(v, label);
      return r >= 0 ? cleanStr_(v[r][1]) : '';
    };
    p.rank = get('Rank');
    p.name = get('Full name');
    p.theme = get('Research theme / area');
    p.program_start = get('Program start (year / term)');
    p.group_start = get('Started with the group');
    p.program_year = get('Current program year');
    p.grad_term = get('Expected graduation term');
    p.effort = get('Appointment / effort');
    p.site = get('Primary site / base');
    p.funding = get('Primary funding source');
    p.thesis_title = get('Working thesis title / topic');
    p.projects_raw = get('Primary project(s)');
    p.focus = get('Current focus / status');
    p.milestones = {};
    [['qual', 'Qualifying exam'], ['comp', 'Comprehensive exam'],
     ['committee', 'Thesis committee formed'],
     ['author', 'Author qualification (if applicable)'],
     ['defense', 'Thesis defense']].forEach(function (pair) {
      var r = findLabelRow_(v, pair[1]);
      p.milestones[pair[0]] = r >= 0
        ? {status: cleanStr_(v[r][1]), date: cleanStr_(v[r][2]), loc: cleanStr_(v[r][3]), notes: cleanStr_(v[r][4])}
        : {status: '', date: '', loc: '', notes: ''};
    });
  }
  if (!p.name) p.name = fileName.replace(/\.[^.]+$/, '').replace(/[_\- ]*info.*$/i, '') || fileName;
  p.projects = parseProjects_(p.projects_raw || '');
  p.start_year = toDecimalYear_(p.program_start);
  p.group_year = toDecimalYear_(p.group_start) || p.start_year;

  var readSheet = function (sheetNames, headerFinder, ncols, mapFn, stops) {
    var sh = (typeof sheetNames === 'string') ? ss.getSheetByName(sheetNames) : sheetByNames_(ss, sheetNames);
    if (!sh) return [];
    var v = values_(sh);
    var h = headerFinder(v);
    if (h < 0) return [];
    return readTable_(v, h, ncols, stops).map(mapFn);
  };
  var byHash = function (v) { return findLabelRow_(v, '#'); };

  // Committee (three sub-tables)
  p.committee_members = []; p.committee_meetings = []; p.progress_reports = [];
  var com = ss.getSheetByName('Committee');
  if (com) {
    var cv = values_(com);
    var rMem = findLabelRow_(cv, 'Committee members');
    var rMeet = findLabelRow_(cv, 'Committee meetings');
    var rProg = findLabelRow_(cv, 'Annual progress reports');
    if (rMem >= 0) readTable_(cv, rMem + 1, 5, ['Committee meetings']).forEach(function (x) {
      p.committee_members.push({name: x[1], role: x[2], dept: x[3], notes: x[4]}); });
    if (rMeet >= 0) readTable_(cv, rMeet + 1, 5, ['Annual progress reports']).forEach(function (x) {
      p.committee_meetings.push({meeting: x[0], date: x[1], loc: x[2], focus: x[3], notes: x[4]}); });
    if (rProg >= 0) readTable_(cv, rProg + 1, 4, []).forEach(function (x) {
      p.progress_reports.push({date: x[1], period: x[2], notes: x[3]}); });
  }

  p.roles = readSheet('Roles', byHash, 7, function (x) {
    return {role: x[1], category: x[2], org: x[3], start: x[4], end: x[5], notes: x[6]}; });
  p.presentations = readSheet('Presentations', byHash, 9, function (x) {
    return {type: x[1], venue: x[2], location: x[3], date: x[4], title: x[5],
            invited: x[6], travel: x[7], link: x[8]}; });
  p.outreach = readSheet(['Outreach & education', 'Outreach'], byHash, 9, function (x) {
    return {activity: x[1], type: x[2], audience: x[3], date: x[4], role: x[5],
            hours: x[6], notes: x[7], url: x[8]}; });
  p.publications = readSheet('Publications', byHash, 7, function (x) {
    return {title: x[1], role: x[2], journal: x[3], status: x[4], date: x[5], doi: x[6]}; });
  p.applications = readSheet('Applications', byHash, 8, function (x) {
    return {name: x[1], type: x[2], status: x[3], requested: x[4], awarded: x[5],
            date: x[6], notes: x[7]}; });
  p.awards = readSheet('Awards', byHash, 6, function (x) {
    return {award: x[1], org: x[2], date: x[3], amount: x[4], notes: x[5]}; });
  p.events = readSheet(SCHOOL_TABS, byHash, 9, function (x) {
    return {type: x[1], event: x[2], location: x[3], start: x[4], end: x[5],
            role: x[6], presented: x[7], notes: x[8]}; });
  p.grants = readSheet(GRANT_TABS, byHash, 9, function (x) {
    return {title: x[1], agency: x[2], role: x[3], amount: x[4], start: x[5],
            end: x[6], status: x[7], notes: x[8]}; });
  return p;
}

/* ============================ model ============================ */

function isPostdoc_(p) { return /postdoc/i.test(p.rank || ''); }
function isFaculty_(p) { return /faculty|professor|staff/i.test(p.rank || '') || /\bpi\b/i.test(p.rank || ''); }
function isStudent_(p) {
  if (isPostdoc_(p) || isFaculty_(p)) return false;
  return /grad|student|undergrad/i.test(p.rank || '');
}

function buildModel_(people, title) {
  var themes = {}, order = 0;
  people.forEach(function (p) {
    var theme = p.theme || 'Unspecified';
    var key = themeKey_(theme);
    p.key = key;
    if (!themes[key]) { themes[key] = {key: key, label: theme, color: PALETTE[order % PALETTE.length]}; order++; }
  });

  var projectKey = {}, membership = {};
  people.forEach(function (p) {
    var pm = {};
    (p.projects || []).forEach(function (pr) {
      if (!(pr.name in projectKey)) projectKey[pr.name] = p.key;
      pm[pr.name] = pr.lead ? 'lead' : 'member';
    });
    if (Object.keys(pm).length) membership[p.name] = pm;
  });
  var projects = Object.keys(projectKey).map(function (n) { return {name: n, key: projectKey[n]}; });

  var flat = function (field) {
    var rows = [];
    people.forEach(function (p) {
      (p[field] || []).forEach(function (rec) {
        var item = {};
        for (var k in rec) item[k] = rec[k];
        item.who = p.name; item.key = p.key;
        rows.push(item);
      });
    });
    return rows;
  };

  var now = new Date();
  var nowDec = now.getFullYear() + now.getMonth() / 12.0 + (now.getDate() / 31.0) / 12.0;

  var passed = function (key) {
    var c = 0;
    people.forEach(function (p) {
      var st = ((p.milestones || {})[key] || {}).status || '';
      if (/^(passed|yes)$/i.test(st)) c++;
    });
    return c;
  };

  var apps = flat('applications');
  var grants = flat('grants');
  var activeGrants = 0, grantTotal = 0;
  grants.forEach(function (g) {
    if (/^(active|awarded)$/i.test(g.status || '')) { activeGrants++; grantTotal += parseMoney_(g.amount); }
  });
  var sites = {};
  people.forEach(function (p) { if (p.site) sites[p.site] = true; });
  var sum = function (field) { var n = 0; people.forEach(function (p) { n += (p[field] || []).length; }); return n; };

  return {
    group: {title: title, generated: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')},
    now: Math.round(nowDec * 1000) / 1000,
    themes: Object.keys(themes).map(function (k) { return themes[k]; }),
    people: people,
    projects: projects,
    membership: membership,
    applications: apps,
    grants: grants,
    events: flat('events'),
    presentations: flat('presentations'),
    awards: flat('awards'),
    publications: flat('publications'),
    roles: flat('roles'),
    outreach: flat('outreach'),
    kpis: {
      people: people.length,
      students: people.filter(isStudent_).length,
      postdocs: people.filter(isPostdoc_).length,
      faculty: people.filter(isFaculty_).length,
      themes: Object.keys(themes).length,
      sites: Object.keys(sites).length,
      quals: passed('qual'),
      comps: passed('comp'),
      publications: sum('publications'),
      presentations: sum('presentations'),
      awards: sum('awards'),
      awarded: apps.filter(function (a) { return /^awarded$/i.test(a.status || ''); }).length,
      grants: activeGrants,
      grant_total: grantTotal
    }
  };
}

/* ============================ summaries ============================ */
/* A port of generate_summaries.py: the same five tables, as 2-D arrays whose
 * first row is the header. Pure data — testable without any Google API. */

var SCHOOL_RE_ = /\b(school|academy|tutorial)\b/i;

function yearOf_(s) {
  var m = String(s || '').match(/(?:19|20)\d{2}/);
  return m ? m[0] : '';
}

function venueSeries_(v) {
  var s = String(v || '').trim();
  var out = s.replace(/[\s,'’]*(?:19|20)\d{2}\s*$/, '')
             .replace(/^[\s,\-–]+|[\s,\-–]+$/g, '');
  return out || s;
}

function sortedJoin_(arr) {
  return arr.filter(function (v, i) { return v && arr.indexOf(v) === i; })
            .sort().join('; ');
}

function collectAttendance_(people) {
  var recs = [];
  people.forEach(function (p) {
    (p.events || []).forEach(function (r) {
      var name = (r.event || '').trim();
      if (!name) return;
      recs.push({person: p.name, name: name, year: yearOf_(r.start),
                 location: r.location || '', descriptor: r.type || '',
                 isSchool: SCHOOL_RE_.test((r.type || '') + ' ' + name)});
    });
    (p.presentations || []).forEach(function (r) {
      var venue = (r.venue || '').trim();
      if (!venue) return;
      recs.push({person: p.name, name: venue, year: yearOf_(r.date),
                 location: r.location || '', descriptor: r.type || '',
                 isSchool: (r.type || '').trim().toLowerCase() === 'lecture' ||
                           SCHOOL_RE_.test(venue)});
    });
  });
  return recs;
}

function aggregateEvents_(records, wantSchools) {
  var agg = {};
  records.forEach(function (rec) {
    if (rec.isSchool !== wantSchools) return;
    var key = venueSeries_(rec.name).toLowerCase();
    if (!agg[key]) agg[key] = {name: venueSeries_(rec.name), attendees: [],
                               years: [], locations: [], types: [], seen: {}, count: 0};
    var e = agg[key];
    var dedup = rec.person.toLowerCase() + '|' + rec.year;
    if (e.seen[dedup]) return;
    e.seen[dedup] = true;
    e.attendees.push(rec.person);
    e.years.push(rec.year);
    e.locations.push(rec.location);
    e.types.push(rec.descriptor);
    e.count++;
  });
  var rows = Object.keys(agg).map(function (k) { return agg[k]; });
  rows.sort(function (a, b) {
    return (b.count - a.count) || a.name.toLowerCase().localeCompare(b.name.toLowerCase());
  });
  return rows;
}

function summaryTables_(people) {
  var pres = [['Person', 'Type', 'Venue / meeting', 'Location', 'Date', 'Title',
               'Invited?', 'External travel funding', 'Link']];
  var pubs = [['Person', 'Title', 'Role', 'Journal / collaboration', 'Status',
               'Date', 'DOI / link']];
  var grants = [['Holder', 'Grant / award title', 'Agency / sponsor', 'Role',
                 'Amount', 'Start', 'End', 'Status', 'Grant no. / notes']];
  people.forEach(function (p) {
    (p.presentations || []).forEach(function (r) {
      pres.push([p.name, r.type, r.venue, r.location, r.date, r.title,
                 r.invited, r.travel, r.link]);
    });
    (p.publications || []).forEach(function (r) {
      pubs.push([p.name, r.title, r.role, r.journal, r.status, r.date, r.doi]);
    });
    (p.grants || []).forEach(function (r) {
      grants.push([p.name, r.title, r.agency, r.role, r.amount, r.start,
                   r.end, r.status, r.notes]);
    });
  });
  var records = collectAttendance_(people);
  var conf = [['Conference / meeting', 'Attendees', '# Attendees', 'Years',
               'Locations', '# Attendances', 'Roles / contributions']];
  aggregateEvents_(records, false).forEach(function (e) {
    var uniq = e.attendees.filter(function (v, i) { return e.attendees.indexOf(v) === i; });
    conf.push([e.name, uniq.slice().sort().join('; '), uniq.length,
               sortedJoin_(e.years), sortedJoin_(e.locations), e.count,
               sortedJoin_(e.types)]);
  });
  var schools = [['School / course', 'Attendees', '# Attendees', 'Years',
                  'Locations', '# Attendances']];
  aggregateEvents_(records, true).forEach(function (e) {
    var uniq = e.attendees.filter(function (v, i) { return e.attendees.indexOf(v) === i; });
    schools.push([e.name, uniq.slice().sort().join('; '), uniq.length,
                  sortedJoin_(e.years), sortedJoin_(e.locations), e.count]);
  });
  return {presentations: pres, publications: pubs, conferences: conf,
          schools: schools, grants: grants};
}

/* ============================ summary Sheet ============================ */

/**
 * Write/update the "Group Summaries" Google Sheet in the folder — one tab per
 * summary table. Updated in place so the group has a single stable, natively
 * viewable Sheet (the timestamped HTML/Doc outputs are the dated archive);
 * cell A1's note records when each tab was last refreshed.
 */
function generateSummarySheet() {
  var folderId = resolveFolderId_();
  var people = readPeople_(folderId);
  var tables = summaryTables_(people);
  var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID || folderId);
  var name = 'Group Summaries';
  var it = folder.getFilesByName(name);
  var ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.openById(it.next().getId());
  } else {
    ss = SpreadsheetApp.create(name);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  [['Presentations', tables.presentations],
   ['Publications', tables.publications],
   ['Conferences', tables.conferences],
   ['Schools', tables.schools],
   ['Grants', tables.grants]].forEach(function (pair) {
    writeTab_(ss, pair[0], pair[1], stamp);
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  Logger.log('Updated "%s" → %s', name, ss.getUrl());
  return ss.getUrl();
}

function writeTab_(ss, name, data, stamp) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  var ncols = data[0].length;
  var rect = data.map(function (row) {
    var r = row.slice(0, ncols).map(function (c) { return c == null ? '' : c; });
    while (r.length < ncols) r.push('');
    return r;
  });
  sh.getRange(1, 1, rect.length, ncols).setValues(rect);
  sh.getRange(1, 1, 1, ncols).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, ncols);
  sh.getRange(1, 1).setNote('Generated ' + stamp + ' from ' + (data.length - 1) + ' row(s).');
}

/* ============================ Doc report ============================ */

function money_(n) {
  return '$' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function docTable_(body, headers, rows) {
  var data = [headers].concat(rows).map(function (row) {
    return row.map(function (c) { return String(c == null ? '' : c); });
  });
  var t = body.appendTable(data);
  t.setBorderWidth(0.5);
  for (var i = 0; i < headers.length; i++) {
    t.getRow(0).getCell(i).editAsText().setBold(true);
  }
  return t;
}

function docSection_(body, title, headers, rows) {
  if (!rows.length) return;
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  docTable_(body, headers, rows);
}

/**
 * Write a timestamped Google Doc report into the folder. Unlike the HTML
 * dashboard, a Doc opens directly in Google Drive with a double click.
 */
function generateDocReport() {
  var folderId = resolveFolderId_();
  var people = readPeople_(folderId);
  var model = buildModel_(people, TITLE);
  var tables = summaryTables_(people);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');

  var doc = DocumentApp.create('group_report_' + stamp);
  var body = doc.getBody();
  var k = model.kpis;

  body.appendParagraph(TITLE).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var sub = body.appendParagraph(
    k.people + ' members · ' + k.students + ' students, ' + k.postdocs +
    ' postdocs' + (k.faculty ? ', ' + k.faculty + ' faculty/staff' : '') +
    ' · generated ' + model.group.generated);
  sub.editAsText().setItalic(true);

  var kpiRows = [
    ['People', k.people], ['Students', k.students], ['Postdocs', k.postdocs],
    ['Faculty / staff', k.faculty], ['Research themes', k.themes], ['Sites', k.sites],
    ['Quals passed', k.quals], ['Comps passed', k.comps],
    ['Publications', k.publications], ['Presentations', k.presentations],
    ['Awards', k.awards], ['Active grants', k.grants],
    ['Funding on the books', k.grant_total ? money_(k.grant_total) : 0],
  ].filter(function (r) { return r[1]; });
  docSection_(body, 'Key numbers', ['Measure', 'Value'], kpiRows);

  docSection_(body, 'People', ['Name', 'Rank', 'Theme', 'Year', 'Site', 'Funding'],
    model.people.map(function (p) {
      return [p.name, p.rank, p.theme, p.program_year || '—', p.site || '—', p.funding || '—'];
    }));

  var students = model.people.filter(isStudent_);
  docSection_(body, 'Milestones & exams',
    ['Student', 'Qual', 'Comp', 'Committee', 'Author', 'Defense'],
    students.map(function (p) {
      var g = function (key) { return ((p.milestones || {})[key] || {}).status || '—'; };
      return [p.name, g('qual'), g('comp'), g('committee'), g('author'), g('defense')];
    }));

  docSection_(body, 'Grants & funding', tables.grants[0], tables.grants.slice(1));
  docSection_(body, 'Applications & fellowships',
    ['Person', 'Name', 'Type', 'Status', 'Date'],
    model.applications.map(function (a) {
      return [a.who, a.name, a.type, a.status, a.date];
    }));
  docSection_(body, 'Publications', tables.publications[0], tables.publications.slice(1));
  docSection_(body, 'Presentations',
    ['Person', 'Type', 'Venue', 'Date', 'Title'],
    model.presentations.map(function (t) {
      return [t.who, t.type, t.venue, t.date, t.title];
    }));
  docSection_(body, 'Conferences attended', tables.conferences[0], tables.conferences.slice(1));
  docSection_(body, 'Schools attended', tables.schools[0], tables.schools.slice(1));
  docSection_(body, 'Awards & honors',
    ['Person', 'Award', 'Organization', 'Date', 'Amount'],
    model.awards.map(function (a) {
      return [a.who, a.award, a.org, a.date, a.amount];
    }));
  docSection_(body, 'Roles & service',
    ['Person', 'Role', 'Category', 'Organization', 'Start', 'End'],
    model.roles.map(function (r) {
      return [r.who, r.role, r.category, r.org, r.start, r.end];
    }));
  docSection_(body, 'Outreach & education',
    ['Person', 'Activity', 'Type', 'Audience / venue', 'Date', 'Role'],
    model.outreach.map(function (o) {
      return [o.who, o.activity, o.type, o.audience, o.date, o.role];
    }));

  var foot = body.appendParagraph(
    'Generated ' + model.group.generated + ' from ' + people.length +
    ' Research Record spreadsheet(s). The interactive dashboard has more detail; ' +
    'this report is the directly-viewable snapshot.');
  foot.editAsText().setItalic(true);

  doc.saveAndClose();
  var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID || folderId);
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  Logger.log('Wrote %s → %s', doc.getName(), doc.getUrl());
  return doc.getUrl();
}

/* ============================ render ============================ */

function renderHtml_(model) {
  var tpl = HtmlService.createHtmlOutputFromFile(TEMPLATE_FILE).getContent();
  var json = JSON.stringify(model).replace(/<\//g, '<\\/');
  return tpl.replace('/*__DATA__*/null', json);
}
