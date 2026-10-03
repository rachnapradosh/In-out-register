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
 *   Daily      - one row per employee per day, with Present / Half day / Absent
 *   Time Out   - every OUT and the IN that followed it
 *   Punch Log  - every single IN / OUT
 *   Employees  - the employee list
 */

const SECRET = 'change-this-password';

const TAB = { dash: 'Dashboard', daily: 'Daily', away: 'Time Out', log: 'Punch Log', emps: 'Employees', rawP: '_punches', rawE: '_employees', rawA: '_attendance' };
const RAW_P = ['id', 'empId', 'type', 'ts', 'reason', 'note', 'updatedAt'];
const RAW_E = ['id', 'name', 'dept', 'active', 'createdAt', 'updatedAt'];
const RAW_A = ['id', 'empId', 'date', 'status', 'updatedAt'];
const STATUS = { present: 'Present', half: 'Half day', absent: 'Absent' };
const DEFAULT_OFFICE = { open: 10 * 60, close: 19 * 60 }; // 10:00 AM to 7:00 PM
const ALL_MONTHS = 'All months';
const NO_REASON = 'No reason';

const C = {
  navy: '#1b2a4e', text: '#1f2533', muted: '#8a8170', hair: '#ebe4d6', gold: '#b08d57',
  cream: '#f6f1e7', creamMuted: '#cfc6b2', paper: '#fffdf9', band: '#faf6ee', sand: '#efe7d6',
  heat: '#e9d3bf', ink: '#2e6b4a', out: '#9c4a2a', warn: '#8a6416',
};
const TITLE = 'Times New Roman';
const BODY = 'Times New Roman';

/* ---------------- Web app entry points ---------------- */

// Opening the URL in a browser shows nothing useful; all data goes through doPost
// so the password is never part of a URL.
function doGet() {
  return json_({ ok: false, error: 'Not available' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  return handle_(body);
}

function handle_(req) {
  if (SECRET === 'change-this-password' || SECRET.length < 10) return json_({ ok: false, error: 'Set SECRET in the script to a password of at least 10 characters, then deploy again' });
  if (String(req.secret || '') !== SECRET) {
    Utilities.sleep(1500); // slows down anyone guessing passwords
    return json_({ ok: false, error: 'Wrong password' });
  }
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
    const props = PropertiesService.getScriptProperties();
    if (Array.isArray(req.reasons)) props.setProperty('reasons', JSON.stringify(req.reasons));
    if (req.office && isFinite(req.office.open) && isFinite(req.office.close)) props.setProperty('office', JSON.stringify(req.office));
    const emps = upsert_(rawSheet_(ss, TAB.rawE, RAW_E), RAW_E, req.employees || []);
    const punches = upsert_(rawSheet_(ss, TAB.rawP, RAW_P), RAW_P, req.punches || []);
    const att = upsert_(rawSheet_(ss, TAB.rawA, RAW_A), RAW_A, req.attendance || []);
    // The data is safe at this point. A styling problem in the visible tabs is
    // logged (Apps Script > Executions) but never reported as a failed backup.
    try { rebuild_(ss, emps, punches, att); } catch (err) { console.error('rebuild failed: ' + ((err && err.stack) || err)); }
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
    attendance: readRaw_(rawSheet_(ss, TAB.rawA, RAW_A), RAW_A)
      .map(r => Object.assign(r, { date: normDate_(r.date, ss.getSpreadsheetTimeZone()) })),
    office: office_(),
  };
}

function office_() {
  try {
    const o = JSON.parse(PropertiesService.getScriptProperties().getProperty('office') || 'null');
    if (o && isFinite(o.open) && isFinite(o.close)) return { open: Number(o.open), close: Number(o.close) };
  } catch (err) { /* fall back to the default hours */ }
  return DEFAULT_OFFICE;
}

// Attendance dates are kept as plain "yyyy-MM-dd" text.
function normDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const t = String(v || '');
  return /^\d{4}-\d{2}-\d{2}$/.test(t) ? t : '';
}

/** Rebuilds every visible tab. Run it from the editor to restyle the sheet by hand. */
function rebuildAll() {
  const ss = SpreadsheetApp.getActive();
  rebuild_(ss, readRaw_(rawSheet_(ss, TAB.rawE, RAW_E), RAW_E), readRaw_(rawSheet_(ss, TAB.rawP, RAW_P), RAW_P),
    readRaw_(rawSheet_(ss, TAB.rawA, RAW_A), RAW_A));
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
    // Keep dates as text so Sheets does not turn them into date values.
    const dc = cols.indexOf('date');
    if (dc >= 0) {
      const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
      out.forEach(r => { r.date = normDate_(r.date, tz); });
      sh.getRange(2, dc + 1, out.length, 1).setNumberFormat('@');
    }
    sh.getRange(2, 1, out.length, cols.length).setValues(out.map(r => cols.map(c => r[c])));
  }
  return out;
}

/* ---------------- Calculations ---------------- */

/* One entry per employee per day, from punches and attendance.
   OUT followed by IN = time out, cut off at closing time.
   A final OUT with no IN after it counts until closing (until now if the
   office is still open). On a half day that final OUT is not counted. */
function computeDays_(emps, punches, att, tz, office) {
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
  const status = new Map();
  for (const r of att || []) {
    const date = normDate_(r.date, tz);
    if (!date || !r.empId || !STATUS[r.status]) continue;
    const key = r.empId + '|' + date;
    status.set(key, String(r.status));
    if (!groups.has(key)) groups.set(key, []);
  }
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const now = Date.now();
  const days = [];
  groups.forEach((list, key) => {
    list.sort((a, b) => a.ts - b.ts);
    const parts = key.split('|');
    const k = parts[1];
    const attv = status.get(key) || (list.length ? 'present' : null);
    if (!attv) return;
    const emp = byEmp.get(parts[0]) || { id: parts[0], name: 'Unknown', dept: '' };
    const dayStart = Utilities.parseDate(k + ' 00:00', tz, 'yyyy-MM-dd HH:mm');
    const close = dayStart.getTime() + office.close * 60000;
    const st = {
      key: k, emp: emp, list: list, att: attv, date: dayStart, close: close, closeDate: new Date(close),
      firstIn: null, leftAt: null, away: [], awayMin: 0, openIn: null, last: list[list.length - 1] || null,
      isToday: k === today,
    };
    const add = (out, back, until) => {
      const min = Math.max(0, (Math.min(until, close) - out.ts) / 60000);
      if (!back && min <= 0) return;
      st.away.push({ out: out, back: back, min: min, toClose: !back });
      st.awayMin += min;
    };
    for (let i = 0; i < list.length; i++) {
      const a = list[i], b = list[i + 1];
      if (a.type === 'in' && !st.firstIn) st.firstIn = a.date;
      if (b && a.type === b.type) continue;
      if (a.type === 'out') {
        if (b) add(a, b, b.ts);
        else {
          st.leftAt = a.date;
          if (attv !== 'half') add(a, null, st.isToday ? Math.min(now, close) : close);
        }
      } else if (!b) {
        st.openIn = a;
      }
    }
    st.stillOut = st.isToday && now < close && !!st.last && st.last.type === 'out';
    st.month = Utilities.formatDate(dayStart, tz, 'MMMM yyyy');
    days.push(st);
  });
  days.sort((a, b) => (a.key < b.key ? 1 : a.key > b.key ? -1 : String(a.emp.name).localeCompare(String(b.emp.name))));
  return days;
}

const dur_ = min => Math.round(min) / 1440; // minutes as a Sheets duration
const DUR = '[h]"h" mm"m"'; // shows 3h 05m
const fmtMin_ = m => { m = Math.round(m); return m < 60 ? m + 'm' : Math.floor(m / 60) + 'h ' + ('0' + (m % 60)).slice(-2) + 'm'; };

/* ---------------- Building the visible tabs ---------------- */

// Every table tab shares one layout: a margin column A, a title block,
// a navy header on row HR and data from row FR.
const HR = 5;
const FR = 6;
const CO = 1; // column offset for the margin

function rebuild_(ss, emps, punches, att) {
  const tz = ss.getSpreadsheetTimeZone();
  const days = computeDays_(emps, punches, att || [], tz, office_());
  const empMap = new Map(emps.map(e => [String(e.id), e]));
  const stamp = 'Updated ' + Utilities.formatDate(new Date(), tz, "d MMMM yyyy 'at' h:mm a");
  buildDaily_(ss, days, stamp);
  buildAway_(ss, days, stamp);
  buildLog_(ss, punches, empMap, stamp);
  buildEmployees_(ss, emps, stamp);
  buildDashboard_(ss, emps, days, tz, stamp);
  orderTabs_(ss);
}

// Column c (1-based within the table) of the data rows.
const dataCol_ = (sh, c, n) => sh.getRange(FR, c + CO, n, 1);

function gradient_(range, to) {
  return SpreadsheetApp.newConditionalFormatRule().setRanges([range])
    .setGradientMinpointWithValue(C.paper, SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue(to || C.heat, SpreadsheetApp.InterpolationType.PERCENTILE, '95').build();
}
const textRule_ = (range, text, color, italic) => {
  const b = SpreadsheetApp.newConditionalFormatRule().setRanges([range]).whenTextEqualTo(text).setFontColor(color);
  if (italic) b.setItalic(true);
  return b.build();
};

function buildDaily_(ss, days, stamp) {
  const head = ['Date', 'Day', 'Employee', 'Department', 'Attendance', 'First in', 'Last out', 'Times out', 'Time out', 'Reasons', 'Notes', 'Month'];
  const rows = days.map(d => {
    const reasons = {};
    d.away.forEach(a => { const r = a.out.reason || NO_REASON; reasons[r] = (reasons[r] || 0) + a.min; });
    const notes = [];
    if (d.stillOut) notes.push('Out now');
    else if (d.away.some(a => a.toClose)) notes.push('Not back before closing');
    if (d.att === 'half' && d.leftAt && !(d.last && d.last.type === 'in')) notes.push('Left early (half day)');
    return [
      d.date, d.date, d.emp.name, d.emp.dept || '', STATUS[d.att],
      d.firstIn || '', d.leftAt || '', d.away.length, dur_(d.awayMin),
      Object.keys(reasons).map(r => r + ' ' + fmtMin_(reasons[r])).join(', '), notes.join(', '), d.month,
    ];
  });
  const sh = table_(ss, TAB.daily, 'Daily Record', 'One line per employee per day  |  ' + stamp, head, rows, {
    widths: [118, 52, 180, 130, 96, 90, 90, 84, 96, 240, 190, 110],
    formats: ['dd mmm yyyy', 'ddd', '@', '@', '@', 'h:mm am/pm', 'h:mm am/pm', '0', DUR, '@', '@', '@'],
    center: [2, 5, 6, 7, 8, 9],
    bold: [3, 9],
  });
  sh.hideColumns(12 + CO);
  if (rows.length) {
    const n = rows.length;
    const attCol = dataCol_(sh, 5, n);
    sh.setConditionalFormatRules([
      gradient_(dataCol_(sh, 9, n)),
      textRule_(attCol, 'Present', C.ink),
      textRule_(attCol, 'Half day', C.warn),
      textRule_(attCol, 'Absent', C.out, true),
      SpreadsheetApp.newConditionalFormatRule().setRanges([dataCol_(sh, 11, n)])
        .whenTextContains('Out now').setFontColor(C.out).setItalic(true).build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([dataCol_(sh, 11, n)])
        .whenTextContains('closing').setFontColor(C.warn).setItalic(true).build(),
    ]);
    dataCol_(sh, 2, n).setFontColor(C.muted);
    dataCol_(sh, 11, n).setFontColor(C.muted).setFontStyle('italic');
  }
}

function buildAway_(ss, days, stamp) {
  const head = ['Date', 'Employee', 'Department', 'Out', 'Back', 'Duration', 'Reason', 'Note'];
  const items = [];
  days.forEach(d => d.away.forEach(a => items.push({ d: d, a: a })));
  items.sort((x, y) => y.a.out.ts - x.a.out.ts);
  const rows = items.map(x => {
    const a = x.a, still = !a.back && x.d.stillOut;
    const note = [a.out.note, a.back ? a.back.note : '', a.back ? '' : still ? 'Still out, counted so far' : 'Not back before closing'];
    return [
      a.out.date, x.d.emp.name, x.d.emp.dept || '', a.out.date, a.back ? a.back.date : still ? '' : x.d.closeDate, dur_(a.min),
      a.out.reason || NO_REASON, note.filter(String).join('; '),
    ];
  });
  const sh = table_(ss, TAB.away, 'Time Out', 'Every OUT and when they came back, or closing time if they did not  |  ' + stamp, head, rows, {
    widths: [118, 180, 130, 92, 92, 90, 140, 260],
    formats: ['dd mmm yyyy', '@', '@', 'h:mm am/pm', 'h:mm am/pm', DUR, '@', '@'],
    center: [4, 5, 6],
    bold: [2, 6],
  });
  if (rows.length) {
    const n = rows.length;
    sh.setConditionalFormatRules([gradient_(dataCol_(sh, 6, n)), textRule_(dataCol_(sh, 7, n), NO_REASON, C.muted, true)]);
  }
}

function buildLog_(ss, punches, empMap, stamp) {
  const head = ['Date', 'Time', 'Employee', 'Department', 'In / Out', 'Reason', 'Note'];
  const sorted = punches.filter(p => Number(p.ts)).sort((a, b) => Number(b.ts) - Number(a.ts));
  const rows = sorted.map(p => {
    const e = empMap.get(String(p.empId)) || { name: 'Unknown', dept: '' };
    const d = new Date(Number(p.ts));
    return [d, d, e.name, e.dept || '', p.type === 'in' ? 'IN' : 'OUT', p.reason || '', p.note || ''];
  });
  const sh = table_(ss, TAB.log, 'Punch Log', 'Every IN and OUT as recorded, newest first  |  ' + stamp, head, rows, {
    widths: [118, 92, 180, 130, 84, 140, 260],
    formats: ['dd mmm yyyy', 'h:mm am/pm', '@', '@', '@', '@', '@'],
    center: [2, 5],
    bold: [3, 5],
  });
  if (rows.length) {
    const r = dataCol_(sh, 5, rows.length);
    sh.setConditionalFormatRules([textRule_(r, 'IN', C.ink), textRule_(r, 'OUT', C.out)]);
  }
}

function buildEmployees_(ss, emps, stamp) {
  const head = ['Name', 'Department', 'Status', 'Added on'];
  const isActive = e => e.active === true || e.active === 'TRUE';
  const list = emps.slice().sort((a, b) =>
    isActive(a) !== isActive(b) ? (isActive(a) ? -1 : 1) : String(a.name).localeCompare(String(b.name)));
  const rows = list.map(e => [
    e.name, e.dept || '', isActive(e) ? 'Active' : 'Removed',
    Number(e.createdAt) ? new Date(Number(e.createdAt)) : '',
  ]);
  const active = list.filter(isActive).length;
  const sh = table_(ss, TAB.emps, 'Employees', active + ' active' + (list.length > active ? ', ' + (list.length - active) + ' removed' : '') + '  |  ' + stamp, head, rows, {
    widths: [220, 180, 110, 130],
    formats: ['@', '@', '@', 'dd mmm yyyy'],
    center: [3, 4],
    bold: [1],
  });
  if (rows.length) {
    const r = dataCol_(sh, 3, rows.length);
    sh.setConditionalFormatRules([textRule_(r, 'Active', C.ink), textRule_(r, 'Removed', C.muted, true)]);
  }
}

/* Dashboard layout (content in columns B to I, A and J are margins):
   1-4   navy title band with a gold rule
   6     month picker
   8-10  four summary cards
   12..  "Employee Summary" table with a total row
   then  "Out Right Now" and the chart */
function buildDashboard_(ss, emps, days, tz, stamp) {
  let sh = ss.getSheetByName(TAB.dash);
  let prevMonth = '';
  if (sh) {
    const pv = String(sh.getRange('C6').getValue() || sh.getRange('C5').getValue() || '');
    prevMonth = pv;
  }
  sh = resetSheet_(ss, TAB.dash);

  const months = [];
  days.forEach(d => { if (months.indexOf(d.month) < 0) months.push(d.month); });
  const thisMonth = Utilities.formatDate(new Date(), tz, 'MMMM yyyy');
  if (months.indexOf(thisMonth) < 0) months.unshift(thisMonth);
  const options = months.concat([ALL_MONTHS]);
  const month = options.indexOf(prevMonth) >= 0 ? prevMonth : thisMonth;

  const isActive = e => e.active === true || e.active === 'TRUE';
  const withData = new Set(days.map(d => String(d.emp.id)));
  const people = emps.filter(e => isActive(e) || withData.has(String(e.id)))
    .sort((a, b) => String(a.name).localeCompare(String(b.name)));
  const n = Math.max(people.length, 1);
  const secRow = 12, headRow = 13, top = 14, totalRow = top + n;
  const outRow = totalRow + 3;

  const W = 10; // columns A..J
  ensureSize_(sh, Math.max(outRow + 60, 80), 12);
  trimSize_(sh, Math.max(outRow + 60, 80), 12);
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.setFontFamily(BODY).setFontSize(11).setFontColor(C.text).setVerticalAlignment('middle').setBackground(C.cream);
  sh.setRowHeights(1, sh.getMaxRows(), 22);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 28);
  [180, 130, 90, 90, 90, 95, 115, 90].forEach((w, i) => sh.setColumnWidth(i + 2, w));
  sh.setColumnWidth(10, 28);
  sh.setColumnWidth(11, 28);
  sh.hideColumns(12); // L6 holds the month criteria used by every formula

  // Title band
  sh.getRange(1, 1, 4, W).setBackground(C.navy);
  sh.setRowHeight(1, 16);
  sh.setRowHeight(2, 44);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(4, 14);
  sh.getRange(5, 1, 1, W).setBackground(C.gold);
  sh.setRowHeight(5, 3);
  sh.getRange('B2').setValue('In-Out Register').setFontFamily(TITLE).setFontSize(24).setFontColor(C.cream);
  sh.getRange('B3').setValue('Attendance overview  |  ' + stamp).setFontSize(10).setFontColor(C.creamMuted);
  sh.getRange('G2:I3').merge().setValue(ss.getSpreadsheetTimeZone().replace('_', ' '))
    .setFontSize(10).setFontColor(C.creamMuted).setHorizontalAlignment('right').setVerticalAlignment('bottom');

  // Month picker (row 6, after a spacer that is row 5's gold rule)
  sh.setRowHeight(6, 40);
  sh.getRange('B6').setValue('MONTH').setFontSize(9).setFontWeight('bold').setFontColor(C.muted).setHorizontalAlignment('right');
  const pick = sh.getRange('C6:D6').merge();
  pick.setNumberFormat('@').setValue(month).setFontFamily(TITLE).setFontSize(13).setFontColor(C.navy)
    .setBackground(C.paper).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, C.navy, SpreadsheetApp.BorderStyle.SOLID);
  pick.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).setAllowInvalid(false).build());
  sh.getRange('E6:I6').merge().setValue('Choose a month and every figure below updates.')
    .setFontSize(10).setFontColor(C.muted).setFontStyle('italic');
  sh.getRange('L6').setFormula('=IF($C$6="' + ALL_MONTHS + '","*",$C$6)');
  const crit = '$L$6';
  const D = "'" + TAB.daily + "'!";
  const col = letter => D + '$' + letter + '$' + FR + ':$' + letter;
  // Daily tab columns (shifted by the margin): D employee, F attendance, I times out, J time out, M month
  const EMP = col('D'), ATT = col('F'), TIMES = col('I'), OUT = col('J'), MONTH = col('M');

  // Summary cards
  sh.setRowHeight(7, 14);
  sh.setRowHeight(8, 26);
  sh.setRowHeight(9, 44);
  sh.setRowHeight(10, 3);
  sh.setRowHeight(11, 22);
  const cards = [
    ['B', 'C', 'Employees', '=' + people.filter(isActive).length, '0'],
    ['D', 'E', 'Absent days', '=COUNTIFS(' + MONTH + ',' + crit + ',' + ATT + ',"Absent")', '0'],
    ['F', 'G', 'Times out', '=SUMIFS(' + TIMES + ',' + MONTH + ',' + crit + ')', '0'],
    ['H', 'I', 'Total time out', '=SUMIFS(' + OUT + ',' + MONTH + ',' + crit + ')', DUR],
  ];
  cards.forEach(t => {
    sh.getRange(t[0] + '8:' + t[1] + '9').setBackground(C.paper);
    sh.getRange(t[0] + '10:' + t[1] + '10').setBackground(C.gold);
    sh.getRange(t[0] + '8:' + t[1] + '8').merge().setValue(t[2].toUpperCase())
      .setFontSize(9).setFontWeight('bold').setFontColor(C.muted).setHorizontalAlignment('center').setVerticalAlignment('bottom');
    sh.getRange(t[0] + '9:' + t[1] + '9').merge().setFormula(t[3]).setNumberFormat(t[4])
      .setFontFamily(TITLE).setFontSize(22).setFontColor(C.navy).setHorizontalAlignment('center');
    // cream gutters between the cards
    sh.getRange(t[0] + '8:' + t[1] + '10').setBorder(null, true, null, true, null, null, C.cream, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  // Employee summary
  sh.setRowHeight(secRow, 34);
  sh.getRange(secRow, 2).setValue('Employee Summary').setFontFamily(TITLE).setFontSize(15).setFontColor(C.navy).setVerticalAlignment('bottom');
  const head = ['Employee', 'Department', 'Present', 'Half days', 'Absent', 'Times out', 'Time out', 'Hours out'];
  headerRow_(sh.getRange(headRow, 2, 1, head.length), head);
  sh.getRange(headRow, 4, 1, head.length - 2).setHorizontalAlignment('center');

  const rows = [];
  if (people.length) {
    people.forEach((e, i) => {
      const r = top + i;
      const who = '$B' + r;
      const count = label => '=COUNTIFS(' + EMP + ',' + who + ',' + MONTH + ',' + crit + ',' + ATT + ',"' + label + '")';
      rows.push([
        e.name, e.dept || '',
        count('Present'), count('Half day'), count('Absent'),
        '=SUMIFS(' + TIMES + ',' + EMP + ',' + who + ',' + MONTH + ',' + crit + ')',
        '=SUMIFS(' + OUT + ',' + EMP + ',' + who + ',' + MONTH + ',' + crit + ')',
        '=ROUND(H' + r + '*24,2)',
      ]);
    });
  } else {
    rows.push(['No employees yet', '', '', '', '', '', '', '']);
  }
  const body = sh.getRange(top, 2, rows.length, head.length);
  body.setValues(rows);
  bodyStyle_(sh, body, rows.length);
  sh.getRange(top, 4, rows.length, 4).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(top, 4, rows.length, 1).setFontColor(C.ink);
  sh.getRange(top, 5, rows.length, 1).setFontColor(C.warn);
  sh.getRange(top, 6, rows.length, 1).setFontColor(C.out);
  sh.getRange(top, 8, rows.length, 1).setNumberFormat(DUR).setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(top, 9, rows.length, 1).setNumberFormat('0.00').setHorizontalAlignment('center').setFontColor(C.muted);
  sh.getRange(top, 2, rows.length, 1).setFontWeight('bold');
  sh.getRange(top, 3, rows.length, 1).setFontColor(C.muted);

  const last = totalRow - 1;
  sh.getRange(totalRow, 2, 1, head.length).setValues([[
    'Total', '',
    '=SUM(D' + top + ':D' + last + ')', '=SUM(E' + top + ':E' + last + ')', '=SUM(F' + top + ':F' + last + ')',
    '=SUM(G' + top + ':G' + last + ')', '=SUM(H' + top + ':H' + last + ')', '=SUM(I' + top + ':I' + last + ')',
  ]]).setFontWeight('bold').setFontColor(C.navy).setBackground(C.sand)
    .setBorder(true, null, true, null, null, null, C.navy, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(totalRow, 2).setFontFamily(TITLE).setFontSize(11);
  sh.getRange(totalRow, 4, 1, 4).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(totalRow, 8).setNumberFormat(DUR).setHorizontalAlignment('center');
  sh.getRange(totalRow, 9).setNumberFormat('0.00').setHorizontalAlignment('center');
  sh.setRowHeight(totalRow, 30);
  if (people.length) {
    const g = SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(top, 8, people.length, 1)])
      .setGradientMinpointWithValue(C.paper, SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpointWithValue(C.heat, SpreadsheetApp.InterpolationType.MAX, '').build();
    sh.setConditionalFormatRules([g]);
  }

  // Out right now (as of this update)
  const outNow = days.filter(d => d.isToday && d.last && d.last.type === 'out').sort((a, b) => a.last.ts - b.last.ts);
  sh.setRowHeight(outRow, 34);
  sh.getRange(outRow, 2).setValue('Out Right Now').setFontFamily(TITLE).setFontSize(15).setFontColor(C.navy).setVerticalAlignment('bottom');
  sh.getRange(outRow, 3, 1, 4).merge().setValue('As of the last update, including anyone who has left for the day')
    .setFontSize(10).setFontColor(C.muted).setFontStyle('italic').setVerticalAlignment('bottom');
  headerRow_(sh.getRange(outRow + 1, 2, 1, 4), ['Employee', 'Out since', 'Reason', 'Department']);
  sh.getRange(outRow + 1, 3).setHorizontalAlignment('center');
  let outRows = outNow.map(d => [d.emp.name, d.last.date, d.last.reason || NO_REASON, d.emp.dept || '']);
  const nobody = !outRows.length;
  if (nobody) outRows = [['Everyone is in', '', '', '']];
  const ob = sh.getRange(outRow + 2, 2, outRows.length, 4);
  ob.setValues(outRows);
  bodyStyle_(sh, ob, outRows.length);
  sh.getRange(outRow + 2, 3, outRows.length, 1).setNumberFormat('h:mm am/pm').setHorizontalAlignment('center').setFontColor(C.out).setFontWeight('bold');
  sh.getRange(outRow + 2, 2, outRows.length, 1).setFontWeight('bold');
  if (nobody) sh.getRange(outRow + 2, 2).setFontWeight('normal').setFontStyle('italic').setFontColor(C.muted);

  // Chart of hours out per employee for the chosen month
  if (people.length) {
    const chartRow = outRow + 4 + outRows.length;
    const font = { fontName: BODY, color: C.text, fontSize: 11 };
    const base = () => sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(top, 2, people.length, 1))
      .addRange(sh.getRange(top, 9, people.length, 1))
      .setPosition(chartRow, 2, 0, 0)
      .setOption('title', 'Hours out by employee, selected month')
      .setOption('legend', { position: 'none' })
      .setOption('colors', [C.navy])
      .setOption('width', 860)
      .setOption('height', Math.max(260, people.length * 30 + 100));
    // Google drops support for chart options from time to time. Try the styled
    // chart first and fall back to a plain one so the backup never fails here.
    try {
      sh.insertChart(base()
        .setOption('titleTextStyle', { fontName: TITLE, color: C.navy, fontSize: 15 })
        .setOption('backgroundColor', C.paper)
        .setOption('hAxis', { title: 'Hours', minValue: 0, textStyle: font, titleTextStyle: font, gridlines: { color: C.hair } })
        .setOption('vAxis', { textStyle: font })
        .build());
    } catch (err) {
      sh.getCharts().forEach(c => sh.removeChart(c));
      try { sh.insertChart(base().build()); } catch (err2) { /* leave the dashboard without a chart */ }
    }
  }

  sh.setFrozenRows(6);
}

/* ---------------- Shared table styling ---------------- */

function resetSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) return ss.insertSheet(name);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getBandings().forEach(b => b.remove());
  if (sh.getFilter()) sh.getFilter().remove();
  sh.setConditionalFormatRules([]);
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  sh.clear();
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

function headerRow_(range, labels) {
  range.setValues([labels.map(l => String(l).toUpperCase())])
    .setBackground(C.navy).setFontColor(C.cream).setFontFamily(BODY).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(null, null, true, null, null, null, C.gold, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  range.getSheet().setRowHeight(range.getRow(), 32);
}

function bodyStyle_(sh, range, n) {
  range.setFontFamily(BODY).setFontSize(11).setFontColor(C.text).setVerticalAlignment('middle')
    .setBorder(null, null, true, null, null, true, C.hair, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeightsForced(range.getRow(), n, 30);
  for (let i = 0; i < n; i++) sh.getRange(range.getRow() + i, range.getColumn(), 1, range.getNumColumns()).setBackground(i % 2 ? C.band : C.paper);
}

function table_(ss, name, title, subtitle, head, rows, opt) {
  const sh = resetSheet_(ss, name);
  const nCols = head.length, nRows = Math.max(rows.length, 1);
  const lastRow = FR + nRows - 1, totalCols = nCols + CO + 1;
  ensureSize_(sh, lastRow + 2, totalCols);
  trimSize_(sh, lastRow + 2, totalCols);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())
    .setFontFamily(BODY).setFontSize(11).setFontColor(C.text).setVerticalAlignment('middle').setBackground(C.paper);
  sh.setRowHeights(1, sh.getMaxRows(), 22);
  sh.setColumnWidth(1, 28);
  sh.setColumnWidth(totalCols, 28);

  // Title block
  sh.setRowHeight(1, 16);
  sh.setRowHeight(2, 40);
  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 14);
  sh.getRange(2, 1 + CO).setValue(title).setFontFamily(TITLE).setFontSize(20).setFontColor(C.navy).setVerticalAlignment('bottom');
  sh.getRange(3, 1 + CO).setValue(subtitle).setFontSize(10).setFontColor(C.muted).setVerticalAlignment('top');
  sh.getRange(3, 1 + CO, 1, nCols).setBorder(null, null, true, null, null, null, C.gold, SpreadsheetApp.BorderStyle.SOLID);

  headerRow_(sh.getRange(HR, 1 + CO, 1, nCols), head);
  sh.setFrozenRows(HR);

  if (rows.length) {
    (opt.formats || []).forEach((f, i) => sh.getRange(FR, i + 1 + CO, rows.length, 1).setNumberFormat(f));
    const body = sh.getRange(FR, 1 + CO, rows.length, nCols);
    body.setValues(rows);
    body.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false).setFirstRowColor(C.paper).setSecondRowColor(C.band);
    body.setBorder(null, null, true, null, null, true, C.hair, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeightsForced(FR, rows.length, 30);
    (opt.bold || []).forEach(c => sh.getRange(FR, c + CO, rows.length, 1).setFontWeight('bold'));
  } else {
    sh.getRange(FR, 1 + CO).setValue('No entries yet').setFontColor(C.muted).setFontStyle('italic');
    sh.setRowHeight(FR, 30);
  }
  (opt.center || []).forEach(c => sh.getRange(HR, c + CO, nRows + 1, 1).setHorizontalAlignment('center'));
  (opt.widths || []).forEach((w, i) => sh.setColumnWidth(i + 1 + CO, w));
  return sh;
}

function orderTabs_(ss) {
  const order = [TAB.dash, TAB.daily, TAB.away, TAB.log, TAB.emps];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
    sh.setTabColor(i === 0 ? C.navy : C.gold);
  });
  // Remove the empty "Sheet1" a new spreadsheet starts with.
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  ss.setActiveSheet(ss.getSheetByName(TAB.dash));
}
