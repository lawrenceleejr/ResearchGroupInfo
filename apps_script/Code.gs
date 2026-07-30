/**
 * Research Group Info — Google Apps Script edition.
 *
 * Runs entirely on Google's servers against a Drive folder of member
 * "Research Record" spreadsheets. No command line, nothing to install.
 *
 * It reproduces generate_dashboard.py: it reads every Google Sheet in a folder,
 * parses the 11 tabs, and writes a timestamped `dashboard_YYYY-MM-DD_HHMM.html`
 * back into the folder (older snapshots are kept, so a history builds up). It can
 * run from a menu, on a daily time-driven trigger, or be published as a web app
 * so the whole group can bookmark a live URL.
 *
 * VIEWING: Google Drive does not render HTML pages, so double-clicking the
 * generated file only offers a download. To view the dashboard live, deploy this
 * as a Web app (see README) and use that URL; to view an archived snapshot,
 * download that timestamped .html and open it in a browser.
 *
 * SETUP
 *   1. Put the members' records in one Drive folder (as Google Sheets — open
 *      each .xlsx in Google Sheets once, or upload with "Convert uploads" on).
 *   2. Create an Apps Script project (script.google.com → New project).
 *   3. Add two files: this Code.gs, and an HTML file named exactly
 *      `dashboard_template` whose contents are apps_script/dashboard_template.html.
 *   4. Set FOLDER_ID below to the Drive folder's id (the long string in its URL).
 *   5. Run `generateDashboard` once and grant permissions.
 *
 * OPTIONAL
 *   - `installDailyTrigger()` regenerates the dashboard every night.
 *   - Deploy → New deployment → Web app to get a live, always-current URL
 *     (doGet renders on demand).
 *
 * PDF: the dashboard draws itself with JavaScript, so a faithful PDF needs a
 *   browser. Open dashboard.html and use the browser's Print → Save as PDF
 *   (a print stylesheet is included), or run generate_dashboard.py --pdf.
 */

var FOLDER_ID = 'PUT_YOUR_DRIVE_FOLDER_ID_HERE';
var OUTPUT_FOLDER_ID = '';            // blank = write into FOLDER_ID
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
      .addItem('Regenerate now', 'generateDashboard')
      .addToUi();
  } catch (e) { /* not a bound script */ }
}

function generateDashboard() {
  var model = buildModel_(readPeople_(), TITLE);
  var html = renderHtml_(model);
  var out = OUTPUT_FOLDER_ID ? DriveApp.getFolderById(OUTPUT_FOLDER_ID)
                             : DriveApp.getFolderById(FOLDER_ID);
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
  var model = buildModel_(readPeople_(), TITLE);
  return HtmlService.createHtmlOutput(renderHtml_(model))
    .setTitle(TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'generateDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('generateDashboard').timeBased().everyDays(1).atHour(4).create();
  Logger.log('Daily trigger installed (≈4am).');
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

function readPeople_() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var names = [];
  while (files.hasNext()) names.push(files.next());
  names.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  var people = [];
  names.forEach(function (f) {
    if (f.getName().toLowerCase().indexOf('dashboard') !== -1) return;
    try {
      people.push(parseSpreadsheet_(SpreadsheetApp.openById(f.getId()), f.getName()));
    } catch (e) {
      Logger.log('SKIP %s: %s', f.getName(), e);
    }
  });
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

/* ============================ render ============================ */

function renderHtml_(model) {
  var tpl = HtmlService.createHtmlOutputFromFile(TEMPLATE_FILE).getContent();
  var json = JSON.stringify(model).replace(/<\//g, '<\\/');
  return tpl.replace('/*__DATA__*/null', json);
}
