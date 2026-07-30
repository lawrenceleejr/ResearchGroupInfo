#!/usr/bin/env node
/**
 * Tests for the Apps Script edition (apps_script/Code.gs).
 *
 * Loads Code.gs with stubbed Google Apps Script APIs and runs it against the
 * fixture produced by dump_fixture.py, asserting:
 *   - parsing/KPIs/summary tables are IDENTICAL to the Python edition
 *   - outputs get timestamped names; folder auto-resolution persists
 *   - template/output files in the folder are not parsed as members
 *   - updateMemberRecords() adds exactly the missing tabs and never writes data
 *
 * Usage: node tests/gas/gas.test.js [fixtureDir]   (default: tests/gas/.fixture)
 */
'use strict';
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const ROOT = path.resolve(__dirname, '..', '..');
const FIXTURE = process.argv[2] || path.join(__dirname, '.fixture');
const sheets = JSON.parse(fs.readFileSync(path.join(FIXTURE, 'sheets.json'), 'utf8'));
const expected = JSON.parse(fs.readFileSync(path.join(FIXTURE, 'expected.json'), 'utf8'));

// Distractors that must be ignored by readPeople_/updateMemberRecords:
sheets['FirstnameLastname_Info_Template.xlsx'] = sheets['LawrenceLee_Info.xlsx'];
sheets['Group Summaries'] = sheets['LawrenceLee_Info.xlsx'];
// Simulate one member on an OLD template (missing the two newest tabs):
const OLD_MEMBER = 'AdamBrown_Info.xlsx';
const adam = JSON.parse(JSON.stringify(sheets[OLD_MEMBER]));
delete adam['Grants & Funding'];
delete adam['Conferences & Schools'];
sheets[OLD_MEMBER] = adam;

/* ------------------------- Google API stubs ------------------------- */
const calls = [];
function FakeRange(v) { this._v = v; }
FakeRange.prototype.getValues = function () { return this._v; };

function FakeTab(ssName, tabName, vals) { this._ss = ssName; this._n = tabName; this._v = vals; }
FakeTab.prototype.getName = function () { return this._n; };
FakeTab.prototype.getMaxRows = function () { return 1000; };
FakeTab.prototype.getMaxColumns = function () { return 12; };
FakeTab.prototype.getDataRange = function () { return new FakeRange(this._v); };
FakeTab.prototype.getRange = function (r, c, nr, nc) {
  const self = this;
  return {
    getDataValidations: () => Array.from({length: nr}, () => Array(nc).fill(null)),
    setDataValidations: (rules) => {
      assert.strictEqual(rules.length, nr, 'DV row dims');
      assert.strictEqual(rules[0].length, nc, 'DV col dims');
      calls.push('setDV ' + self._ss + '!' + self._n);
    },
    setValues: () => { throw new Error('DATA WRITE on ' + self._ss + '!' + self._n); },
    setFontWeight: () => {}, setNote: () => {},
  };
};
FakeTab.prototype.copyTo = function (destSS) {
  const self = this;
  return {setName: (n) => { calls.push('copyTab "' + n + '" -> ' + destSS._name); destSS._s[n] = self._v; }};
};

function FakeSS(name) { this._name = name; this._s = sheets[name]; this._id = name; }
FakeSS.prototype.getSheetByName = function (n) { return this._s[n] ? new FakeTab(this._name, n, this._s[n]) : null; };
FakeSS.prototype.getSheets = function () {
  const self = this;
  return Object.keys(this._s).map((n) => new FakeTab(self._name, n, self._s[n]));
};
FakeSS.prototype.getId = function () { return this._id; };

const fileList = Object.keys(sheets).map((n) => ({getName: () => n, getId: () => n}));
const iterOf = (a) => { let i = 0; return {hasNext: () => i < a.length, next: () => a[i++]}; };
const created = [];

global.MimeType = {GOOGLE_SHEETS: 'gs', HTML: 'html', MICROSOFT_EXCEL: 'xlsx'};
global.DriveApp = {
  getFolderById: () => ({
    getFilesByType: (t) => (t === 'gs' ? iterOf(fileList) : iterOf([])),
    getFilesByName: () => iterOf([]),
    createFile: (n) => { created.push(n); return {getName: () => n, getUrl: () => 'fake://' + n}; },
  }),
  getFileById: () => ({getParents: () => iterOf([{getId: () => 'FOLDER123'}]), moveTo: () => {}}),
};
const props = {};
global.PropertiesService = {getScriptProperties: () => ({
  getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; }})};
global.Logger = {log: () => {}};
global.Session = {getScriptTimeZone: () => 'UTC'};
const pad = (n) => String(n).padStart(2, '0');
global.Utilities = {formatDate: (d, tz, f) => f
  .replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1))
  .replace('dd', pad(d.getDate())).replace('HH', pad(d.getHours()))
  .replace('mm', pad(d.getMinutes()))};
global.HtmlService = {createHtmlOutputFromFile: (f) => ({
  getContent: () => fs.readFileSync(path.join(ROOT, 'apps_script', f + '.html'), 'utf8')})};
const summaryTabs = {};
function fakeSummaryTab(name) {
  return {clear: () => {}, setFrozenRows: () => {}, autoResizeColumns: () => {},
    getRange: (r, c, nr, nc) => ({
      setValues: (v) => {
        assert.strictEqual(v.length, nr);
        v.forEach((row) => assert.strictEqual(row.length, nc, 'ragged tab ' + name));
        summaryTabs[name] = v;
      },
      setFontWeight: () => {}, setNote: () => {}})};
}
const madeTabs = {};
const summarySS = {
  insertSheet: (n) => { madeTabs[n] = fakeSummaryTab(n); return madeTabs[n]; },
  getSheetByName: (n) => madeTabs[n] || null,
  getSheets: () => Object.values(madeTabs),
  deleteSheet: () => {}, getId: () => 'SS1', getUrl: () => 'fake://summaries',
};
global.SpreadsheetApp = {
  openById: (id) => new FakeSS(id),
  getActiveSpreadsheet: () => new FakeSS('LawrenceLee_Info.xlsx'),
  create: () => summarySS,
  getUi: () => { throw new Error('no ui'); },
};
let docTables = 0;
const fakeText = {setBold: () => fakeText, setItalic: () => fakeText};
const fakePara = {setHeading: () => fakePara, editAsText: () => fakeText};
global.DocumentApp = {
  ParagraphHeading: {HEADING1: 1, HEADING2: 2, NORMAL: 0},
  create: (name) => {
    created.push(name);
    return {getBody: () => ({
      appendParagraph: () => fakePara,
      appendTable: (data) => {
        const w = data[0].length;
        data.forEach((r) => { assert.strictEqual(r.length, w, 'ragged doc table');
          r.forEach((c) => assert.strictEqual(typeof c, 'string')); });
        docTables++;
        return {setBorderWidth: () => {}, getRow: () => ({getCell: () => ({editAsText: () => fakeText})})};
      },
    }), saveAndClose: () => {}, getId: () => 'DOC1', getName: () => name,
        getUrl: () => 'fake://doc'};
  },
};

/* ------------------------------ load ------------------------------ */
// Indirect eval => non-strict global scope, so Code.gs's function declarations
// land on globalThis and are callable below.
(0, eval)(fs.readFileSync(path.join(ROOT, 'apps_script', 'Code.gs'), 'utf8'));
/* global resolveFolderId_, readPeople_, buildModel_, summaryTables_,
          renderHtml_, generateDashboard, generateDocReport,
          generateSummarySheet, updateMemberRecords, parseSpreadsheet_ */

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('FAIL  ' + name + '\n      ' + e.message); process.exitCode = 1; }
}
const str2d = (t) => t.map((r) => r.map((c) => String(c == null ? '' : c)));

/* ------------------------------ tests ------------------------------ */
console.log('gas.test.js');

test('folder auto-resolution persists', () => {
  assert.strictEqual(resolveFolderId_(), 'FOLDER123');
  assert.strictEqual(props.FOLDER_ID, 'FOLDER123');
});

const people = readPeople_('F');
test('readPeople_ skips template/output files', () => {
  assert.strictEqual(people.length, 6);
  assert.ok(!people.some((p) => /firstname|summar/i.test(p.name)));
});

test('people names match the Python edition', () => {
  assert.deepStrictEqual(people.map((p) => p.name).sort(), expected.people);
});

const model = buildModel_(people, 'T');
test('KPIs identical to the Python edition', () => {
  assert.deepStrictEqual(model.kpis, expected.kpis);
});

test('themes identical to the Python edition', () => {
  assert.deepStrictEqual(model.themes.map((t) => t.label), expected.themes);
});

const tables = summaryTables_(people);
for (const name of ['presentations', 'publications', 'conferences', 'schools', 'grants']) {
  test(`summary table "${name}" identical to the Python CSV`, () => {
    assert.deepStrictEqual(str2d(tables[name]), expected.tables[name]);
  });
}

test('renderHtml_ embeds valid JSON', () => {
  const html = renderHtml_(model);
  assert.ok(!html.includes('/*__DATA__*/null'));
  const payload = html.split('const DATA = ')[1].split(';\n')[0];
  assert.strictEqual(JSON.parse(payload.replace(/<\\\//g, '</')).kpis.people, 6);
});

test('generateDashboard writes a timestamped html', () => {
  generateDashboard();
  assert.ok(/^dashboard_\d{4}-\d{2}-\d{2}_\d{4}\.html$/.test(created.at(-1)), created.at(-1));
});

test('generateDocReport writes a timestamped Doc with tables', () => {
  generateDocReport();
  assert.ok(/^group_report_\d{4}-\d{2}-\d{2}_\d{4}$/.test(created.at(-1)), created.at(-1));
  assert.ok(docTables >= 10, 'doc tables: ' + docTables);
});

test('generateSummarySheet writes the five tabs', () => {
  generateSummarySheet();
  assert.deepStrictEqual(Object.keys(summaryTabs).sort(),
    ['Conferences', 'Grants', 'Presentations', 'Publications', 'Schools']);
  assert.deepStrictEqual(str2d(summaryTabs.Conferences), expected.tables.conferences);
});

test('updateMemberRecords adds exactly the missing tabs, touches no data', () => {
  calls.length = 0;
  updateMemberRecords();
  const copies = calls.filter((c) => c.startsWith('copyTab'));
  assert.deepStrictEqual(copies.sort(), [
    'copyTab "Conferences & Schools" -> ' + OLD_MEMBER,
    'copyTab "Grants & Funding" -> ' + OLD_MEMBER,
  ]);
  // dropdown refresh on every pre-existing tab: 5 members x 11 + old member x 9
  assert.strictEqual(calls.filter((c) => c.startsWith('setDV')).length, 64);
});

test('updated old member parses with the new tabs', () => {
  const p = parseSpreadsheet_(new FakeSS(OLD_MEMBER), OLD_MEMBER);
  assert.ok(p.grants.length > 0);
  assert.ok(p.events.length > 0);
});

console.log(process.exitCode ? 'FAILED' : `all ${passed} tests passed`);
