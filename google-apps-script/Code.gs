/**
 * In-Out Register: Google Sheet backup
 *
 * Paste this whole file into Extensions > Apps Script of a new Google Sheet,
 * change SECRET below, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Full steps are in README.md.
 *
 * The phone sends every IN / OUT here. Raw data is kept in two hidden tabs
 * (_punches, _employees); the visible tabs are rebuilt from it on every sync:
 *   Dashboard  - month picker, totals, per-employee summary, chart, who is out
 *   Daily      - one row per employee per day
 *   Time Out   - every OUT and the IN that followed it
 *   Punch Log  - every single IN / OUT
 *   Employees  - the employee list
 */

const SECRET = 'change-this-password';

const TAB = { dash: 'Dashboard', daily: 'Daily', away: 'Time Out', log: 'Punch Log', emps: 'Employees', rawP: '_punches', rawE: '_employees' };
const RAW_P = ['id', 'empId', 'type', 'ts', 'reason', 'note', 'updatedAt'];
const RAW_E = ['id', 'name', 'dept', 'active', 'createdAt', 'updatedAt'];
const ALL_MONTHS = 'All months';
const NO_REASON = 'No reason';

const C = {
  brand: '#1b2a4e', brandDark: '#14203d', brandSoft: '#e7eaf2',
  text: '#1a2340', muted: '#6e6a60', line: '#e7dfcf', band: '#faf6ee', white: '#fffdf8',
  ink: '#1f7a4d', inSoft: '#e5f2ea', out: '#b4532a', outSoft: '#f8e9df',
  warn: '#9a6a12', warnSoft: '#f8efd9', cream: '#f6f1e7', creamText: '#e9e1cf',
};
const FONT = 'Roboto';

/* ---------------- Web app entry points ---------------- */

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  return handle_(body);
}

function handle_(req) {
  if (SECRET === 'change-this-password') return json_({ ok: false, error: 'Change SECRET in the script first, then deploy again' });
  if (String(req.secret || '') !== SECRET) return json_({ ok: false, error: 'Wrong password' });
  try {
    switch (req.action) {
      case 'ping': return json_({ ok: true, sheet: SpreadsheetApp.getActive().getName() });
      case 'sync': return json_(sync_(req));
      case 'restore': return json_(restore_());
      default: return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Sync & restore ---------------- */

function sync_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    if (req.timeZone && ss.getSpreadsheetTimeZone() !== req.timeZone) ss.setSpreadsheetTimeZone(req.timeZone);
    if (Array.isArray(req.reasons)) PropertiesService.getScriptProperties().setProperty('reasons', JSON.stringify(req.reasons));
    const emps = upsert_(rawSheet_(ss, TAB.rawE, RAW_E), RAW_E, req.employees || []);
    const punches = upsert_(rawSheet_(ss, TAB.rawP, RAW_P), RAW_P, req.punches || []);
    rebuild_(ss, emps, punches);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function restore_() {
  const ss = SpreadsheetApp.getActive();
  const reasons = JSON.parse(PropertiesService.getScriptProperties().getProperty('reasons') || '[]');
  return {
    ok: true,
    reasons: reasons,
    employees: readRaw_(rawSheet_(ss, TAB.rawE, RAW_E), RAW_E).map(e => Object.assign(e, { active: e.active === true || e.active === 'TRUE' })),
    punches: readRaw_(rawSheet_(ss, TAB.rawP, RAW_P), RAW_P),
  };
}

/** Rebuilds every visible tab. Run it from the editor to restyle the sheet by hand. */
function rebuildAll() {
  const ss = SpreadsheetApp.getActive();
  rebuild_(ss, readRaw_(rawSheet_(ss, TAB.rawE, RAW_E), RAW_E), readRaw_(rawSheet_(ss, TAB.rawP, RAW_P), RAW_P));
}

/* ---------------- Raw storage ---------------- */

function rawSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function readRaw_(sh, cols) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; cols.forEach((c, i) => { o[c] = r[i]; }); return o; });
}

// Newer updatedAt wins. Records marked deleted are removed. Returns the full list afterwards.
function upsert_(sh, cols, incoming) {
  const rows = readRaw_(sh, cols);
  if (!incoming.length) return rows;
  const index = new Map(rows.map((r, i) => [String(r.id), i]));
  const drop = new Set();
  for (const inc of incoming) {
    if (!inc || !inc.id) continue;
    const id = String(inc.id);
    const at = index.get(id);
    if (at !== undefined && Number(rows[at].updatedAt || 0) > Number(inc.updatedAt || 0)) continue;
    if (inc.deleted) { if (at !== undefined) drop.add(at); continue; }
    const rec = {};
    cols.forEach(c => { rec[c] = inc[c] === undefined || inc[c] === null ? '' : inc[c]; });
    if (at !== undefined) rows[at] = rec;
    else { index.set(id, rows.length); rows.push(rec); }
  }
  const out = rows.filter((_, i) => !drop.has(i));
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (out.length) {
    ensureSize_(sh, out.length + 1, cols.length);
    sh.getRange(2, 1, out.length, cols.length).setValues(out.map(r => cols.map(c => r[c])));
  }
  return out;
}

/* ---------------- Calculations ---------------- */

function computeDays_(emps, punches, tz) {
  const byEmp = new Map(emps.map(e => [String(e.id), e]));
  const groups = new Map();
  for (const p of punches) {
    const ts = Number(p.ts);
    if (!ts || (p.type !== 'in' && p.type !== 'out')) continue;
    const d = new Date(ts);
    const key = p.empId + '|' + Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ id: p.id, type: p.type, ts: ts, date: d, reason: String(p.reason || ''), note: String(p.note || '') });
  }
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const days = [];
  groups.forEach((list, key) => {
    list.sort((a, b) => a.ts - b.ts);
    const parts = key.split('|');
    const emp = byEmp.get(parts[0]) || { id: parts[0], name: 'Unknown', dept: '' };
    const st = { key: parts[1], emp: emp, list: list, firstIn: null, leftAt: null, away: [], awayMin: 0, insideMin: 0, openIn: null, last: list[list.length - 1] };
    for (let i = 0; i < list.length; i++) {
      const a = list[i], b = list[i + 1];
      if (a.type === 'in' && !st.firstIn) st.firstIn = a.date;
      if (b && a.type === b.type) continue;
      if (a.type === 'out') {
        if (b) { const m = (b.ts - a.ts) / 60000; st.away.push({ out: a, back: b, min: m }); st.awayMin += m; }
        else st.leftAt = a.date;
      } else if (b) {
        st.insideMin += (b.ts - a.ts) / 60000;
      } else {
        st.openIn = a;
      }
    }
    st.isToday = parts[1] === today;
    st.month = Utilities.formatDate(list[0].date, tz, 'MMMM yyyy');
    days.push(st);
  });
  days.sort((a, b) => (a.key < b.key ? 1 : a.key > b.key ? -1 : String(a.emp.name).localeCompare(String(b.emp.name))));
  return days;
}

const dur_ = min => Math.round(min) / 1440; // minutes as a Sheets duration
const fmtMin_ = m => { m = Math.round(m); return m < 60 ? m + 'm' : Math.floor(m / 60) + 'h ' + ('0' + (m % 60)).slice(-2) + 'm'; };

/* ---------------- Building the visible tabs ---------------- */

function rebuild_(ss, emps, punches) {
  const tz = ss.getSpreadsheetTimeZone();
  const days = computeDays_(emps, punches, tz);
  const empMap = new Map(emps.map(e => [String(e.id), e]));
  buildDaily_(ss, days);
  buildAway_(ss, days);
  buildLog_(ss, punches, empMap);
  buildEmployees_(ss, emps);
  buildDashboard_(ss, emps, days, tz);
  orderTabs_(ss);
}

function buildDaily_(ss, days) {
  const head = ['Date', 'Day', 'Employee', 'Department', 'First IN', 'Left at', 'Times out', 'Time out', 'Time inside', 'Reasons', 'Flags', 'Month'];
  const rows = days.map(d => {
    const reasons = {};
    d.away.forEach(a => { const r = a.out.reason || NO_REASON; reasons[r] = (reasons[r] || 0) + a.min; });
    const flags = [];
    if (d.openIn && !d.isToday) flags.push('No OUT after last IN');
    if (d.isToday && d.last.type === 'out') flags.push('Out now');
    return [
      d.list[0].date, d.list[0].date, d.emp.name, d.emp.dept || '',
      d.firstIn || '', d.leftAt || '', d.away.length, dur_(d.awayMin), dur_(d.insideMin),
      Object.keys(reasons).map(r => r + ' ' + fmtMin_(reasons[r])).join(', '), flags.join(', '), d.month,
    ];
  });
  const sh = table_(ss, TAB.daily, head, rows, {
    widths: [120, 60, 190, 140, 90, 90, 90, 90, 100, 280, 170, 120],
    formats: ['dd mmm yyyy', 'ddd', '@', '@', 'h:mm am/pm', 'h:mm am/pm', '0', '[h]:mm', '[h]:mm', '@', '@', '@'],
    center: [2, 5, 6, 7, 8, 9],
  });
  sh.hideColumns(12);
  if (rows.length) {
    const n = rows.length;
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(2, 8, n, 1)])
        .setGradientMinpointWithValue(C.white, SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMaxpointWithValue('#e9a27f', SpreadsheetApp.InterpolationType.PERCENTILE, '95').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(2, 11, n, 1)])
        .whenTextContains('No OUT').setFontColor(C.warn).setBackground(C.warnSoft).setBold(true).build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(2, 11, n, 1)])
        .whenTextContains('Out now').setFontColor(C.out).setBackground(C.outSoft).setBold(true).build(),
    ];
    sh.setConditionalFormatRules(rules);
    sh.getRange(2, 8, n, 1).setFontWeight('bold');
  }
}

function buildAway_(ss, days) {
  const head = ['Date', 'Employee', 'Department', 'Out', 'Back', 'Duration', 'Reason', 'Note'];
  const items = [];
  days.forEach(d => d.away.forEach(a => items.push({ d: d, a: a })));
  items.sort((x, y) => y.a.out.ts - x.a.out.ts);
  const rows = items.map(x => [
    x.a.out.date, x.d.emp.name, x.d.emp.dept || '', x.a.out.date, x.a.back.date, dur_(x.a.min),
    x.a.out.reason || NO_REASON, [x.a.out.note, x.a.back.note].filter(String).join('; '),
  ]);
  const sh = table_(ss, TAB.away, head, rows, {
    widths: [120, 190, 140, 95, 95, 90, 150, 280],
    formats: ['dd mmm yyyy', '@', '@', 'h:mm am/pm', 'h:mm am/pm', '[h]:mm', '@', '@'],
    center: [4, 5, 6],
  });
  if (rows.length) {
    const n = rows.length;
    sh.getRange(2, 6, n, 1).setFontWeight('bold');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(2, 6, n, 1)])
        .setGradientMinpointWithValue(C.white, SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMaxpointWithValue('#e9a27f', SpreadsheetApp.InterpolationType.PERCENTILE, '95').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(2, 7, n, 1)])
        .whenTextEqualTo(NO_REASON).setFontColor(C.muted).setItalic(true).build(),
    ]);
  }
}

function buildLog_(ss, punches, empMap) {
  const head = ['Date', 'Time', 'Employee', 'Department', 'IN / OUT', 'Reason', 'Note'];
  const sorted = punches.filter(p => Number(p.ts)).sort((a, b) => Number(b.ts) - Number(a.ts));
  const rows = sorted.map(p => {
    const e = empMap.get(String(p.empId)) || { name: 'Unknown', dept: '' };
    const d = new Date(Number(p.ts));
    return [d, d, e.name, e.dept || '', p.type === 'in' ? 'IN' : 'OUT', p.reason || '', p.note || ''];
  });
  const sh = table_(ss, TAB.log, head, rows, {
    widths: [120, 95, 190, 140, 90, 150, 280],
    formats: ['dd mmm yyyy', 'h:mm am/pm', '@', '@', '@', '@', '@'],
    center: [2, 5],
  });
  if (rows.length) {
    const r = sh.getRange(2, 5, rows.length, 1);
    r.setFontWeight('bold');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().setRanges([r]).whenTextEqualTo('IN').setBackground(C.inSoft).setFontColor(C.ink).build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([r]).whenTextEqualTo('OUT').setBackground(C.outSoft).setFontColor(C.out).build(),
    ]);
  }
}

function buildEmployees_(ss, emps) {
  const head = ['Name', 'Department', 'Status', 'Added on'];
  const list = emps.slice().sort((a, b) => {
    const aa = a.active === true || a.active === 'TRUE', bb = b.active === true || b.active === 'TRUE';
    return aa !== bb ? (aa ? -1 : 1) : String(a.name).localeCompare(String(b.name));
  });
  const rows = list.map(e => [
    e.name, e.dept || '', e.active === true || e.active === 'TRUE' ? 'Active' : 'Removed',
    Number(e.createdAt) ? new Date(Number(e.createdAt)) : '',
  ]);
  const sh = table_(ss, TAB.emps, head, rows, { widths: [220, 180, 110, 130], formats: ['@', '@', '@', 'dd mmm yyyy'], center: [3, 4] });
  if (rows.length) {
    const r = sh.getRange(2, 3, rows.length, 1);
    r.setFontWeight('bold');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().setRanges([r]).whenTextEqualTo('Active').setBackground(C.inSoft).setFontColor(C.ink).build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([r]).whenTextEqualTo('Removed').setBackground('#efeadf').setFontColor(C.muted).build(),
    ]);
  }
}

/* Dashboard layout (columns B to I, A is a margin):
   2-3   title band
   5     month picker
   7-9   four KPI tiles
   11    "Employee summary" heading, 12 table header, 13.. rows, then a total row
   then  "Out right now" list and the chart */
function buildDashboard_(ss, emps, days, tz) {
  let sh = ss.getSheetByName(TAB.dash);
  const prevMonth = sh ? String(sh.getRange('C5').getValue() || '') : '';
  sh = resetSheet_(ss, TAB.dash);

  const months = [];
  days.forEach(d => { if (months.indexOf(d.month) < 0) months.push(d.month); });
  const thisMonth = Utilities.formatDate(new Date(), tz, 'MMMM yyyy');
  if (months.indexOf(thisMonth) < 0) months.unshift(thisMonth);
  const options = months.concat([ALL_MONTHS]);
  const month = options.indexOf(prevMonth) >= 0 ? prevMonth : thisMonth;

  const withData = new Set(days.map(d => String(d.emp.id)));
  const people = emps.filter(e => e.active === true || e.active === 'TRUE' || withData.has(String(e.id)))
    .sort((a, b) => String(a.name).localeCompare(String(b.name)));
  const n = Math.max(people.length, 1);
  const top = 13, totalRow = top + n;
  const outNowRow = totalRow + 3;

  ensureSize_(sh, Math.max(outNowRow + 40, 60), 12);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily(FONT).setFontColor(C.text).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 24);
  [200, 150, 110, 100, 110, 130, 120, 100].forEach((w, i) => sh.setColumnWidth(i + 2, w));
  sh.setColumnWidth(10, 24);
  sh.setColumnWidth(11, 24);
  sh.hideColumns(12); // L5 holds the criteria used by every formula

  // Title band
  sh.getRange('A1:J1').setBackground(C.brand);
  sh.getRange('A2:J3').setBackground(C.brand);
  sh.setRowHeight(1, 10);
  sh.setRowHeight(2, 42);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(4, 14);
  sh.getRange('B2').setValue('In-Out Register').setFontSize(22).setFontWeight('bold').setFontColor(C.white);
  sh.getRange('B3').setValue('Last updated ' + Utilities.formatDate(new Date(), tz, "d MMM yyyy 'at' h:mm a"))
    .setFontSize(10).setFontColor(C.creamText);
  sh.getRange('I2:I3').merge().setValue(ss.getSpreadsheetTimeZone()).setFontSize(9).setFontColor(C.creamText).setHorizontalAlignment('right');

  // Month picker
  sh.setRowHeight(5, 34);
  sh.getRange('B5').setValue('Showing').setFontColor(C.muted).setFontWeight('bold').setHorizontalAlignment('right');
  const pick = sh.getRange('C5:D5').merge();
  pick.setNumberFormat('@').setValue(month).setFontWeight('bold').setFontSize(12).setFontColor(C.brandDark).setBackground(C.brandSoft)
    .setHorizontalAlignment('center').setBorder(true, true, true, true, false, false, C.brand, SpreadsheetApp.BorderStyle.SOLID);
  pick.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).setAllowInvalid(false).build());
  sh.getRange('E5:G5').merge().setValue('Change the month here; every number below updates.').setFontSize(9).setFontColor(C.muted).setFontStyle('italic');
  sh.getRange('L5').setFormula('=IF($C$5="' + ALL_MONTHS + '","*",$C$5)');
  const crit = '$L$5';
  const D = "'" + TAB.daily + "'!";
  const col = c => D + '$' + c + '$2:$' + c;

  // KPI tiles
  sh.setRowHeight(6, 14);
  sh.setRowHeight(7, 22);
  sh.setRowHeight(8, 40);
  sh.setRowHeight(9, 8);
  const tiles = [
    ['B', 'C', 'Employees', '=' + people.filter(e => e.active === true || e.active === 'TRUE').length, '0'],
    ['D', 'E', 'Times out', '=SUMIFS(' + col('G') + ',' + col('L') + ',' + crit + ')', '0'],
    ['F', 'G', 'Total time out', '=SUMIFS(' + col('H') + ',' + col('L') + ',' + crit + ')', '[h]:mm'],
    ['H', 'I', 'Average out per day', '=IFERROR(SUMIFS(' + col('H') + ',' + col('L') + ',' + crit + ')/COUNTIFS(' + col('L') + ',' + crit + '),0)', '[h]:mm'],
  ];
  tiles.forEach(t => {
    const label = sh.getRange(t[0] + '7:' + t[1] + '7').merge();
    const value = sh.getRange(t[0] + '8:' + t[1] + '8').merge();
    const whole = sh.getRange(t[0] + '7:' + t[1] + '9');
    whole.setBackground(C.band);
    sh.getRange(t[0] + '9:' + t[1] + '9').setBackground(C.brand);
    label.setValue(t[2].toUpperCase()).setFontSize(9).setFontWeight('bold').setFontColor(C.muted).setHorizontalAlignment('center');
    value.setFormula(t[3]).setNumberFormat(t[4]).setFontSize(22).setFontWeight('bold').setFontColor(C.brandDark).setHorizontalAlignment('center');
  });

  // Employee summary
  sh.setRowHeight(10, 18);
  sh.getRange('B11').setValue('Employee summary').setFontSize(13).setFontWeight('bold').setFontColor(C.brandDark);
  const head = ['Employee', 'Department', 'Days present', 'Times out', 'Time out', 'Average per day', 'Time inside', 'Hours out'];
  sh.getRange(12, 2, 1, head.length).setValues([head])
    .setBackground(C.brand).setFontColor(C.white).setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setWrap(true);
  sh.getRange(12, 2).setHorizontalAlignment('left');
  sh.setRowHeight(12, 32);

  const rows = [];
  if (people.length) {
    people.forEach((e, i) => {
      const r = top + i;
      const who = '$B' + r;
      rows.push([
        e.name, e.dept || '',
        '=COUNTIFS(' + col('C') + ',' + who + ',' + col('L') + ',' + crit + ')',
        '=SUMIFS(' + col('G') + ',' + col('C') + ',' + who + ',' + col('L') + ',' + crit + ')',
        '=SUMIFS(' + col('H') + ',' + col('C') + ',' + who + ',' + col('L') + ',' + crit + ')',
        '=IFERROR(F' + r + '/D' + r + ',0)',
        '=SUMIFS(' + col('I') + ',' + col('C') + ',' + who + ',' + col('L') + ',' + crit + ')',
        '=ROUND(F' + r + '*24,2)',
      ]);
    });
  } else {
    rows.push(['No employees yet', '', '', '', '', '', '', '']);
  }
  const body = sh.getRange(top, 2, rows.length, head.length);
  body.setValues(rows).setFontSize(10);
  sh.getRange(top, 4, rows.length, 2).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(top, 6, rows.length, 3).setNumberFormat('[h]:mm').setHorizontalAlignment('center');
  sh.getRange(top, 9, rows.length, 1).setNumberFormat('0.00').setHorizontalAlignment('center').setFontColor(C.muted);
  sh.getRange(top, 6, rows.length, 1).setFontWeight('bold');
  sh.getRange(top, 2, rows.length, 1).setFontWeight('bold');
  for (let i = 0; i < rows.length; i++) {
    sh.setRowHeight(top + i, 26);
    if (i % 2) sh.getRange(top + i, 2, 1, head.length).setBackground(C.band);
  }
  body.setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);

  // Total row
  const tr = sh.getRange(totalRow, 2, 1, head.length);
  tr.setValues([[
    'Total', '',
    '=SUM(D' + top + ':D' + (totalRow - 1) + ')', '=SUM(E' + top + ':E' + (totalRow - 1) + ')',
    '=SUM(F' + top + ':F' + (totalRow - 1) + ')', '=IFERROR(F' + totalRow + '/D' + totalRow + ',0)',
    '=SUM(H' + top + ':H' + (totalRow - 1) + ')', '=SUM(I' + top + ':I' + (totalRow - 1) + ')',
  ]]).setFontWeight('bold').setBackground(C.brandSoft).setFontColor(C.brandDark)
    .setBorder(true, null, true, null, null, null, C.brand, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(totalRow, 4, 1, 2).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(totalRow, 6, 1, 3).setNumberFormat('[h]:mm').setHorizontalAlignment('center');
  sh.getRange(totalRow, 9).setNumberFormat('0.00').setHorizontalAlignment('center');
  sh.setRowHeight(totalRow, 28);

  if (people.length) {
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(top, 6, people.length, 1)])
        .setGradientMinpointWithValue(C.white, SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMaxpointWithValue('#e9a27f', SpreadsheetApp.InterpolationType.MAX, '').build(),
    ]);
  }

  // Out right now (as of this update)
  const outNow = days.filter(d => d.isToday && d.last.type === 'out')
    .sort((a, b) => a.last.ts - b.last.ts);
  sh.getRange(outNowRow, 2).setValue('Out right now').setFontSize(13).setFontWeight('bold').setFontColor(C.brandDark);
  sh.getRange(outNowRow, 3, 1, 3).merge().setValue('As of the last update. Includes anyone who has left for the day.')
    .setFontSize(9).setFontColor(C.muted).setFontStyle('italic');
  const oh = sh.getRange(outNowRow + 1, 2, 1, 4);
  oh.setValues([['Employee', 'Out since', 'Reason', 'Department']]).setBackground(C.out).setFontColor(C.white).setFontWeight('bold').setFontSize(10);
  let outRows = outNow.map(d => [d.emp.name, d.last.date, d.last.reason || NO_REASON, d.emp.dept || '']);
  if (!outRows.length) outRows = [['Nobody is out', '', '', '']];
  const ob = sh.getRange(outNowRow + 2, 2, outRows.length, 4);
  ob.setValues(outRows).setFontSize(10).setBackground(C.outSoft)
    .setBorder(null, null, true, null, null, true, '#ecc9b5', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(outNowRow + 2, 3, outRows.length, 1).setNumberFormat('h:mm am/pm').setHorizontalAlignment('center');
  sh.getRange(outNowRow + 2, 2, outRows.length, 1).setFontWeight('bold');

  // Chart of hours out per employee for the chosen month
  if (people.length) {
    const chartRow = outNowRow + 3 + outRows.length;
    const chart = sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(top, 2, people.length, 1))
      .addRange(sh.getRange(top, 9, people.length, 1))
      .setPosition(chartRow, 2, 0, 0)
      .setOption('title', 'Hours out by employee (selected month)')
      .setOption('titleTextStyle', { color: C.brandDark, fontSize: 14, bold: true })
      .setOption('legend', { position: 'none' })
      .setOption('colors', [C.out])
      .setOption('backgroundColor', C.white)
      .setOption('hAxis', { title: 'Hours', minValue: 0, gridlines: { color: '#efeadf' } })
      .setOption('width', 860)
      .setOption('height', Math.max(260, people.length * 30 + 90))
      .build();
    sh.insertChart(chart);
  }

  sh.setFrozenRows(5);
  ss.setActiveSheet(sh);
}

/* ---------------- Shared table styling ---------------- */

function resetSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) return ss.insertSheet(name);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getBandings().forEach(b => b.remove());
  if (sh.getFilter()) sh.getFilter().remove();
  sh.setConditionalFormatRules([]);
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  sh.clear();
  sh.setFrozenRows(0);
  sh.showColumns(1, sh.getMaxColumns());
  return sh;
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function trimSize_(sh, rows, cols) {
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
}

function table_(ss, name, head, rows, opt) {
  const sh = resetSheet_(ss, name);
  const nCols = head.length, nRows = Math.max(rows.length, 1);
  ensureSize_(sh, nRows + 1, nCols);
  trimSize_(sh, nRows + 1, nCols);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, nRows + 1, nCols).setFontFamily(FONT).setFontSize(10).setFontColor(C.text).setVerticalAlignment('middle');

  const h = sh.getRange(1, 1, 1, nCols);
  h.setValues([head]).setBackground(C.brand).setFontColor(C.white).setFontWeight('bold').setHorizontalAlignment('left');
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);

  if (rows.length) {
    const body = sh.getRange(2, 1, rows.length, nCols);
    (opt.formats || []).forEach((f, i) => sh.getRange(2, i + 1, rows.length, 1).setNumberFormat(f));
    body.setValues(rows);
    sh.setRowHeightsForced(2, rows.length, 26);
    body.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
      .setFirstRowColor(C.white).setSecondRowColor(C.band);
    body.setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(1, 1, rows.length + 1, nCols).createFilter();
  } else {
    sh.getRange(2, 1).setValue('No entries yet').setFontColor(C.muted).setFontStyle('italic');
  }
  (opt.center || []).forEach(c => sh.getRange(1, c, nRows + 1, 1).setHorizontalAlignment('center'));
  (opt.widths || []).forEach((w, i) => sh.setColumnWidth(i + 1, w));
  return sh;
}

function orderTabs_(ss) {
  [TAB.dash, TAB.daily, TAB.away, TAB.log, TAB.emps].forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  [TAB.dash, TAB.daily, TAB.away, TAB.log, TAB.emps].forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.setTabColor([C.brand, '#2b4c8c', C.out, '#6e6a60', C.ink][i]);
  });
  // Remove the empty "Sheet1" a new spreadsheet starts with.
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  ss.setActiveSheet(ss.getSheetByName(TAB.dash));
}
