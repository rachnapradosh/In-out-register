'use strict';

/* ============================================================
   In-Out Register
   Each morning an employee is marked Present (IN at opening time),
   Half day or Absent. After that one tap records OUT and IN using
   this phone's clock; a reason for going out is optional. An OUT
   with no IN after it counts as time out until closing time.
   Data lives on this phone (IndexedDB) and is backed up to a
   Google Sheet when one is connected.
   ============================================================ */

const APP_VERSION = '2.1.0';
const COLOR_COUNT = 6;
const NO_REASON = 'No reason';
const DEFAULT_SETTINGS = {
  reasons: ['Lunch', 'Tea break', 'Personal work', 'Bank', 'Office work'],
  scriptUrl: '',
  secret: '',
  keepAwake: false,
  lastSync: 0,
  openMin: 10 * 60,   // office opens 10:00 AM (minutes after midnight)
  closeMin: 19 * 60,  // office closes 7:00 PM
};
const STATUS_LABEL = { present: 'Present', half: 'Half day', absent: 'Absent' };

const S = {
  employees: [],
  punches: [],
  days: [],      // attendance: one record per employee per day
  settings: structuredClone(DEFAULT_SETTINGS),
  view: 'home',
  search: '',
  filter: 'all',
  report: { range: 'today', from: '', to: '', open: null },
  sheetEmp: null,
  pickReason: '',
  modal: null,
  undo: null,
  sync: { busy: false, error: '' },
  overlays: [],
};

/* ---------- Storage (IndexedDB) ---------- */
const DB = {
  db: null,
  open() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('inout-register', 2);
      req.onupgradeneeded = () => {
        const d = req.result;
        for (const name of ['employees', 'punches', 'days']) {
          if (!d.objectStoreNames.contains(name)) d.createObjectStore(name, { keyPath: 'id' });
        }
        if (!d.objectStoreNames.contains('meta')) d.createObjectStore('meta', { keyPath: 'key' });
      };
      req.onsuccess = () => { this.db = req.result; resolve(); };
      req.onerror = () => reject(req.error);
    });
  },
  tx(store, mode, fn) {
    return new Promise((resolve, reject) => {
      const t = this.db.transaction(store, mode);
      const out = fn(t.objectStore(store));
      t.oncomplete = () => resolve(out ? out.result : undefined);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  },
  all(store) { return this.tx(store, 'readonly', st => st.getAll()); },
  put(store, items) { return this.tx(store, 'readwrite', st => { items.forEach(i => st.put(i)); }); },
  del(store, ids) { return this.tx(store, 'readwrite', st => { ids.forEach(id => st.delete(id)); }); },
};

/* ---------- Helpers ---------- */
const $ = (sel, root = document) => root.querySelector(sel);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const pad = n => String(n).padStart(2, '0');
const dayKey = ts => { const d = new Date(ts); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const todayKey = () => dayKey(Date.now());
const keyToTs = k => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d).getTime(); };
const fmtTime = ts => new Date(ts).toLocaleTimeString('en-IN', { hour: 'numeric', minute: '2-digit', hour12: true }).toUpperCase();
const fmtDate = ts => new Date(ts).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
const fmtLongDate = ts => new Date(ts).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
const fmtDay = ts => {
  const k = dayKey(ts);
  if (k === todayKey()) return 'Today';
  const y = new Date(); y.setDate(y.getDate() - 1);
  if (k === dayKey(y)) return 'Yesterday';
  return new Date(ts).toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' });
};
const fmtMin = m => { m = Math.round(m); return m < 60 ? `${m}m` : `${Math.floor(m / 60)}h ${pad(m % 60)}m`; };
const fmtElapsed = ms => {
  const s = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), x = s % 60;
  return h ? `${h}:${pad(m)}:${pad(x)}` : `${m}:${pad(x)}`;
};
const fmtDateFull = ts => new Date(ts).toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });

/* Date and time pickers. The phone's own datetime box shows 24-hour time
   in whatever layout the phone uses, so the app draws its own:
   a date button (opens the phone's calendar) and hour / minute / AM-PM. */
function dateField(name, key, opts = {}) {
  const change = opts.change ? `data-change="${opts.change}"` : 'data-change="date-label"';
  return `<div class="date-btn" data-action="open-date">
    <svg viewBox="0 0 24 24"><rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4"/></svg>
    <span class="date-text"${opts.short ? ' data-short' : ''}>${(opts.short ? fmtDate : fmtDateFull)(keyToTs(key))}</span>
    <input type="date" name="${name}" value="${key}" max="${todayKey()}" ${change} required tabindex="-1">
  </div>`;
}
function timeField(ts, prefix = '') {
  const d = new Date(ts);
  const h = d.getHours() % 12 || 12, m = d.getMinutes(), pm = d.getHours() >= 12;
  const opt = (v, label, sel) => `<option value="${v}" ${sel ? 'selected' : ''}>${label}</option>`;
  return `<div class="time-pick">
    <select class="input" name="${prefix}hh" aria-label="Hour">${Array.from({ length: 12 }, (_, i) => opt(i + 1, i + 1, i + 1 === h)).join('')}</select>
    <span class="colon">:</span>
    <select class="input" name="${prefix}mm" aria-label="Minute">${Array.from({ length: 60 }, (_, i) => opt(i, pad(i), i === m)).join('')}</select>
    <div class="ampm">
      <label><input type="radio" name="${prefix}ap" value="AM" ${pm ? '' : 'checked'}><span>AM</span></label>
      <label><input type="radio" name="${prefix}ap" value="PM" ${pm ? 'checked' : ''}><span>PM</span></label>
    </div>
  </div>`;
}
// Minutes after midnight from a timeField.
const readTime = (fd, prefix = '') =>
  ((Number(fd.get(prefix + 'hh')) % 12) + (fd.get(prefix + 'ap') === 'PM' ? 12 : 0)) * 60 + Number(fd.get(prefix + 'mm'));
function readDateTime(fd) {
  const key = String(fd.get('date') || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) return null;
  return minToTs(key, readTime(fd));
}
const minToTs = (k, min) => { const [y, mo, d] = k.split('-').map(Number); return new Date(y, mo - 1, d, Math.floor(min / 60), min % 60).getTime(); };
const fmtClock = min => fmtTime(minToTs(todayKey(), min));
function shiftDay(key, days) {
  const d = new Date(keyToTs(key)); d.setDate(d.getDate() + days); return dayKey(d.getTime());
}
const initials = name => name.trim().split(/\s+/).slice(0, 2).map(w => w[0]).join('').toUpperCase();
const byName = (a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
const buzz = () => { try { navigator.vibrate && navigator.vibrate(35); } catch { /* not supported */ } };
const plural = (n, one, many = one + 's') => `${n} ${n === 1 ? one : many}`;

const ICON = {
  close: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  search: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>',
};

/* ---------- Data access ---------- */
const live = () => S.punches.filter(p => !p.deleted);
const activeEmps = () => S.employees.filter(e => e.active).sort(byName);
const empById = id => S.employees.find(e => e.id === id);
const empName = id => empById(id)?.name || 'Unknown';
const colorClass = reason => {
  if (!reason || reason === NO_REASON) return 'c-none';
  const i = S.settings.reasons.indexOf(reason);
  return `c-${(i < 0 ? COLOR_COUNT - 1 : i) % COLOR_COUNT}`;
};
const typeLabel = t => (t === 'in' ? 'IN' : 'OUT');

function pushTo(map, key, item) {
  let arr = map.get(key);
  if (!arr) map.set(key, arr = []);
  arr.push(item);
}
function sortLists(map) {
  for (const arr of map.values()) arr.sort((a, b) => a.ts - b.ts);
  return map;
}

// Today's punches per employee, oldest first.
function todayIndex() {
  const today = todayKey();
  const m = new Map();
  for (const p of S.punches) if (!p.deleted && dayKey(p.ts) === today) pushTo(m, p.empId, p);
  return sortLists(m);
}
// Punches grouped by "empId|yyyy-mm-dd".
function groupByEmpDay(list) {
  const m = new Map();
  for (const p of list) pushTo(m, `${p.empId}|${dayKey(p.ts)}`, p);
  return sortLists(m);
}

// Attendance records keyed "empId|yyyy-mm-dd".
function daysIndex() {
  const m = new Map();
  for (const r of S.days) if (!r.deleted) m.set(r.id, r);
  return m;
}
const dayRec = (empId, k) => S.days.find(r => r.id === `${empId}|${k}` && !r.deleted);
// Present / half day / absent for a day; a day with entries but no record counts as present.
const attOf = (rec, list) => rec ? rec.status : list && list.length ? 'present' : null;

/* Where an employee is right now:
   type: 'in' | 'out' | 'none' (not marked yet) | 'absent'
   att:  'present' | 'half' | 'absent' | null */
function statusFrom(list, rec) {
  const att = attOf(rec, list);
  if (att === 'absent') return { type: 'absent', punch: null, att };
  const last = list && list[list.length - 1];
  return last ? { type: last.type, punch: last, att } : { type: 'none', punch: null, att };
}
const statusOf = empId => statusFrom(todayIndex().get(empId), dayRec(empId, todayKey()));

/* Works out one employee's day from their punches (oldest first).
   OUT followed by IN = time out of the office, cut off at closing time.
   A final OUT with no IN after it counts until closing time (until now
   if the office is still open), because nobody marks IN after closing.
   On a half day the final OUT is not counted: they are on half-day leave. */
function dayStats(list, k, att) {
  const close = minToTs(k, S.settings.closeMin);
  const st = { firstIn: null, leftAt: null, away: [], awayMin: 0, openIn: null, openOut: null, close };
  const add = (out, back, until) => {
    const min = Math.max(0, (Math.min(until, close) - out.ts) / 60000);
    if (!back && min <= 0) return;
    st.away.push({ out, back, min, toClose: !back });
    st.awayMin += min;
  };
  for (let i = 0; i < list.length; i++) {
    const a = list[i], b = list[i + 1];
    if (a.type === 'in' && st.firstIn == null) st.firstIn = a.ts;
    if (b && a.type === b.type) continue; // repeated punch: the later one counts
    if (a.type === 'out') {
      if (b) add(a, b, b.ts);
      else {
        st.openOut = a; st.leftAt = a.ts;
        if (att !== 'half') add(a, null, k === todayKey() ? Math.min(Date.now(), close) : close);
      }
    } else if (!b) {
      st.openIn = a;
    }
  }
  return st;
}

async function savePunch(p) {
  p.updatedAt = Date.now();
  p.dirty = true;
  if (!S.punches.includes(p)) S.punches.push(p);
  await DB.put('punches', [p]);
  queueSync();
}
async function saveEmployee(e) {
  e.updatedAt = Date.now();
  e.dirty = true;
  if (!S.employees.includes(e)) S.employees.push(e);
  await DB.put('employees', [e]);
  queueSync();
}
async function saveDay(rec) {
  rec.updatedAt = Date.now();
  rec.dirty = true;
  if (!S.days.includes(rec)) S.days.push(rec);
  await DB.put('days', [rec]);
  queueSync();
}
async function saveSettings() {
  await DB.put('meta', [{ key: 'settings', value: S.settings }]);
}

/* ---------- Attendance ---------- */
/* Marks today's attendance. Present and half day also record IN at
   inMin (default: opening time) when there is no entry yet today,
   unless addIn is false. Absent removes today's entries.
   Returns a function that undoes it. */
async function setAttendance(empId, status, inMin = S.settings.openMin, addIn = true) {
  const k = todayKey(), id = `${empId}|${k}`;
  const list = todayIndex().get(empId) || [];
  let rec = S.days.find(r => r.id === id);
  const prev = rec ? { status: rec.status, deleted: !!rec.deleted } : null;
  const removed = status === 'absent' ? list : [];
  for (const p of removed) { p.deleted = true; await savePunch(p); }
  if (!rec) rec = { id, empId, date: k, status, deleted: false };
  Object.assign(rec, { status, deleted: false });
  await saveDay(rec);
  let added = null;
  if (addIn && status !== 'absent' && !list.length) {
    added = { id: uid(), empId, type: 'in', ts: minToTs(k, inMin), reason: '', note: '', deleted: false };
    await savePunch(added);
  }
  return async () => {
    if (prev) Object.assign(rec, prev); else rec.deleted = true;
    await saveDay(rec);
    if (added) { added.deleted = true; await savePunch(added); }
    for (const p of removed) { p.deleted = false; await savePunch(p); }
  };
}

async function markAttendance(empId, status, inMin) {
  const list = todayIndex().get(empId) || [];
  if (status === 'absent' && list.length &&
      !confirm(`Mark ${empName(empId)} absent? Today's ${plural(list.length, 'entry', 'entries')} will be removed.`)) return;
  const undo = await setAttendance(empId, status, inMin ?? S.settings.openMin);
  buzz();
  if (S.overlays.includes('sheet') && status === 'absent') closeOverlay();
  render();
  const st = statusOf(empId);
  let msg = `${empName(empId)}: ${STATUS_LABEL[status]}`;
  if (status !== 'absent' && st.punch && st.punch.type === 'in') msg += `, IN ${fmtTime(st.punch.ts)}`;
  toast(msg, async () => { await undo(); render(); });
}

async function markAllPresent() {
  const k = todayKey();
  const idx = todayIndex();
  const todo = activeEmps().filter(e => !dayRec(e.id, k) && !(idx.get(e.id) || []).length);
  if (!todo.length) return;
  if (!confirm(`Mark ${plural(todo.length, 'employee')} present with IN at ${fmtClock(S.settings.openMin)}?`)) return;
  const undos = [];
  for (const e of todo) undos.push(await setAttendance(e.id, 'present'));
  buzz();
  render();
  toast(`${plural(todo.length, 'employee')} marked present`, async () => { for (const u of undos) await u(); render(); });
}

/* ---------- Core action ---------- */
async function punch(empId, type, reason = '') {
  const cur = statusOf(empId);
  if (cur.type === type) { toast(`${empName(empId)} is already ${typeLabel(type)}`); return; }
  if (cur.type === 'absent') { toast(`${empName(empId)} is marked absent today`); return; }
  // An IN or OUT on an unmarked day means they came to work.
  const undoAtt = cur.att ? null : await setAttendance(empId, 'present', 0, false);
  const p = { id: uid(), empId, type, ts: Date.now(), reason: type === 'out' ? reason : '', note: '', deleted: false };
  await savePunch(p);
  buzz();
  S.pickReason = '';
  if (S.overlays.includes('sheet')) closeOverlay();
  render();
  let msg = `${empName(empId)} ${typeLabel(type)} at ${fmtTime(p.ts)}`;
  if (type === 'in' && cur.type === 'out') msg += `, was out ${fmtMin((p.ts - cur.punch.ts) / 60000)}`;
  if (type === 'out' && reason) msg += ` for ${reason}`;
  toast(msg, async () => { p.deleted = true; await savePunch(p); if (undoAtt) await undoAtt(); render(); });
}

/* ---------- Overlays (with Android back-button support) ---------- */
function openOverlay(kind) {
  S.overlays.push(kind);
  history.pushState({ overlay: S.overlays.length }, '');
  renderOverlays();
}
function closeOverlay() {
  if (S.overlays.length) history.back();
}
window.addEventListener('popstate', () => {
  S.overlays.pop();
  if (!S.overlays.includes('sheet')) { S.sheetEmp = null; S.pickReason = ''; }
  if (!S.overlays.includes('modal')) S.modal = null;
  renderOverlays();
  render();
});
function renderOverlays() {
  $('#sheet').classList.toggle('hidden', !S.overlays.includes('sheet'));
  $('#modal').classList.toggle('hidden', !S.overlays.includes('modal'));
  document.body.style.overflow = S.overlays.length ? 'hidden' : '';
  if (S.overlays.includes('sheet')) renderSheet();
  if (S.overlays.includes('modal')) renderModal();
}

/* ---------- Toast ---------- */
let toastTimer = null;
function toast(msg, undo) {
  const t = $('#toast');
  t.innerHTML = `<span>${esc(msg)}</span>${undo ? '<button data-action="undo">UNDO</button>' : ''}`;
  S.undo = undo || null;
  t.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.classList.add('hidden'); S.undo = null; }, undo ? 10000 : 3500);
}

/* ============================================================
   Rendering
   ============================================================ */
function render() {
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('active', b.dataset.view === S.view));
  const views = { home: renderHome, reports: renderReports, employees: renderEmployees, settings: renderSettings };
  views[S.view]();
  if (S.overlays.includes('sheet')) renderSheet();
  updateBadge();
  tick();
}

/* ---------- Home ---------- */
function renderHome() {
  const view = $('#view');
  if (!$('#homeBody')) {
    view.innerHTML = `
      <div class="search">${ICON.search}
        <input class="input" id="search" type="search" placeholder="Search employee" autocomplete="off" value="${esc(S.search)}">
      </div>
      <div id="homeBody"></div>`;
  }
  renderHomeBody();
}

function renderHomeBody() {
  const body = $('#homeBody');
  const emps = activeEmps();
  if (!emps.length) {
    body.innerHTML = `<div class="card empty">
      <h3>No employees yet</h3>
      <p>Add your employees once. Each morning mark them present, then one tap on OUT or IN records the time.</p>
      <button class="btn" data-action="nav" data-view="employees">Add employees</button>
    </div>`;
    return;
  }
  const k = todayKey();
  const idx = todayIndex();
  const recs = daysIndex();
  const status = new Map(emps.map(e => [e.id, statusFrom(idx.get(e.id), recs.get(`${e.id}|${k}`))]));
  const count = { in: 0, out: 0, none: 0, absent: 0 };
  const att = { present: 0, half: 0, absent: 0, none: 0 };
  emps.forEach(e => { const st = status.get(e.id); count[st.type]++; att[st.att || 'none']++; });

  const q = S.search.trim().toLowerCase();
  const match = e => !q || e.name.toLowerCase().includes(q) || (e.dept || '').toLowerCase().includes(q);
  const groups = { none: [], out: [], in: [], absent: [] };
  emps.filter(match).forEach(e => groups[status.get(e.id).type].push(e));
  groups.out.sort((a, b) => status.get(a.id).punch.ts - status.get(b.id).punch.ts);

  let html = `<div class="card att-card">
    <div class="att-head"><span>Today's attendance</span><span class="small muted">Office ${fmtClock(S.settings.openMin)} to ${fmtClock(S.settings.closeMin)}</span></div>
    <div class="att-counts">
      <div class="a-present"><b>${att.present}</b><span>Present</span></div>
      <div class="a-half"><b>${att.half}</b><span>Half day</span></div>
      <div class="a-absent"><b>${att.absent}</b><span>Absent</span></div>
      <div class="a-none"><b>${att.none}</b><span>Not marked</span></div>
    </div>
    ${att.none ? `<button class="btn block" data-action="all-present">Mark ${att.none === emps.length ? 'everyone' : `the other ${att.none}`} present, IN ${fmtClock(S.settings.openMin)}</button>` : ''}
  </div>`;

  const chip = (key, label, n) =>
    `<button class="chip chip-${key} ${S.filter === key ? 'active' : ''}" data-action="filter" data-filter="${key}"><span class="dot"></span>${label} <b>${n}</b></button>`;
  html += `<div class="chips">
    ${chip('all', 'All', emps.length)}
    ${chip('in', 'In', count.in)}
    ${chip('out', 'Out', count.out)}
    ${chip('none', 'Not marked', count.none)}
    ${chip('absent', 'Absent', count.absent)}
  </div>`;

  const sections = [['none', 'Not marked today'], ['out', 'Out now'], ['in', 'In office'], ['absent', 'Absent today']];
  let shown = 0;
  for (const [key, title] of sections) {
    if (S.filter !== 'all' && S.filter !== key) continue;
    const list = groups[key];
    if (!list.length) continue;
    shown += list.length;
    html += `<div class="section-title"><span>${title}</span><span>${list.length}</span></div><div class="list">`;
    html += list.map(e => empRow(e, status.get(e.id), idx.get(e.id) || [])).join('');
    html += '</div>';
  }
  if (!shown) html += `<div class="empty">${q ? 'No matching employees' : 'Nobody in this group'}</div>`;
  body.innerHTML = html;
}

function empRow(e, st, list) {
  const p = st.punch;
  const half = st.att === 'half' ? '<span class="att-tag half">Half day</span> ' : '';
  let sub, right;
  if (st.type === 'out') {
    sub = `${half}<span class="timer" data-since="${p.ts}"></span> out since ${fmtTime(p.ts)}${p.reason ? `, <span class="reason ${colorClass(p.reason)}">${esc(p.reason)}</span>` : ''}`;
  } else if (st.type === 'in') {
    const s = dayStats(list, todayKey(), st.att);
    sub = `${half}In since ${fmtTime(p.ts)}${s.away.length ? `, out ${fmtMin(s.awayMin)} today` : ''}`;
  } else if (st.type === 'absent') {
    sub = 'Absent today';
  } else {
    sub = 'Not marked today';
  }
  if (st.type === 'none') {
    right = `<button class="pbtn present" data-action="mark" data-status="present" data-id="${e.id}">Present</button>
      <button class="pbtn absent" data-action="mark" data-status="absent" data-id="${e.id}">Absent</button>`;
  } else if (st.type === 'absent') {
    right = `<button class="pbtn change" data-action="open-emp" data-id="${e.id}">Change</button>`;
  } else {
    right = `<button class="pbtn in" data-action="punch" data-type="in" data-id="${e.id}" ${st.type === 'in' ? 'disabled' : ''}>IN</button>
      <button class="pbtn out" data-action="punch" data-type="out" data-id="${e.id}" ${st.type === 'out' ? 'disabled' : ''}>OUT</button>`;
  }
  return `<div class="row st-${st.type}">
    <button class="row-main" data-action="open-emp" data-id="${e.id}">
      <span class="avatar">${esc(initials(e.name))}</span>
      <span class="row-text"><span class="name">${esc(e.name)}</span><span class="sub">${sub}</span></span>
    </button>
    <div class="punch-pair">${right}</div>
  </div>`;
}

/* ---------- Employee sheet ---------- */
function renderSheet() {
  const e = empById(S.sheetEmp);
  const panel = $('#sheetPanel');
  if (!e) { panel.innerHTML = ''; return; }
  const k = todayKey();
  const list = todayIndex().get(e.id) || [];
  const st = statusFrom(list, dayRec(e.id, k));
  const head = `
    <div class="grabber"></div>
    <div class="sheet-head">
      <span class="avatar">${esc(initials(e.name))}</span>
      <div><h2>${esc(e.name)}</h2><div class="small muted">${esc(e.dept || '')}</div></div>
      <button class="icon-btn" data-action="close" aria-label="Close">${ICON.close}</button>
    </div>`;

  // Not marked yet, or absent: choose attendance first.
  if (st.type === 'none' || st.type === 'absent') {
    const absent = st.type === 'absent';
    panel.innerHTML = `${head}
      <div class="status-box ${absent ? 'absent' : 'none'}"><div class="label">${absent ? 'Absent today' : 'Not marked today'}</div></div>
      <form data-form="attend">
        <div class="field"><span>Came in at</span>${timeField(minToTs(k, S.settings.openMin), 'a')}</div>
        <div class="att-choices">
          <button type="submit" class="big in" value="present">Present<small>IN at the time above</small></button>
          <button type="submit" class="big half" value="half">Half day<small>Leaving early is not counted</small></button>
        </div>
        ${absent ? '' : '<button type="submit" class="btn danger block" value="absent" style="margin-top:10px">Absent today</button>'}
      </form>
      ${absent ? '' : `<p class="small muted" style="margin-top:12px">Present records IN at ${fmtClock(S.settings.openMin)} unless you change the time. Anyone who goes OUT and does not come back is counted as out until ${fmtClock(S.settings.closeMin)}.</p>`}`;
    return;
  }

  const stats = dayStats(list, k, st.att);
  const p = st.punch;
  const selected = st.type === 'out' ? p.reason : S.pickReason;
  let box;
  if (st.type === 'out') {
    box = `<div class="status-box out">
      <div class="label">Out since ${fmtTime(p.ts)}${p.reason ? ` for ${esc(p.reason)}` : ''}</div>
      <span class="timer" data-since="${p.ts}"></span>
      ${st.att === 'half' ? '' : `<div class="small">Counted until ${fmtClock(S.settings.closeMin)} if not back</div>`}</div>`;
  } else {
    box = `<div class="status-box in">
      <div class="label">In since ${fmtTime(p.ts)}</div>
      <span class="timer" data-since="${p.ts}"></span></div>`;
  }
  const reasonChips = S.settings.reasons.map(r =>
    `<button class="rchip ${colorClass(r)} ${selected === r ? 'active' : ''}" data-action="set-reason" data-reason="${esc(r)}"><span class="dot"></span>${esc(r)}</button>`).join('');
  const attBtn = (value, label) =>
    `<button class="${st.att === value ? 'active' : ''}" data-action="set-att" data-status="${value}" data-id="${e.id}">${label}</button>`;

  panel.innerHTML = `${head}
    <div class="att-switch">${attBtn('present', 'Present')}${attBtn('half', 'Half day')}${attBtn('absent', 'Absent')}</div>
    ${box}
    <div class="big-pair">
      <button class="big in" data-action="punch" data-type="in" data-id="${e.id}" ${st.type === 'in' ? 'disabled' : ''}>IN<small>${st.type === 'in' ? 'Already in' : 'Back in office'}</small></button>
      <button class="big out" data-action="punch" data-type="out" data-id="${e.id}" data-reason="${esc(S.pickReason)}" ${st.type === 'out' ? 'disabled' : ''}>OUT<small>${st.type === 'out' ? 'Already out' : S.pickReason ? esc(S.pickReason) : 'Going out now'}</small></button>
    </div>
    ${S.settings.reasons.length ? `<div class="reason-block">
      <div class="small muted">${st.type === 'out' ? 'Reason for this OUT (optional, tap to change)' : 'Reason for going out (optional, choose before OUT)'}</div>
      <div class="rchips">${reasonChips}</div>
    </div>` : ''}
    <div class="stat-row">
      <div class="stat"><div class="v">${stats.firstIn ? fmtTime(stats.firstIn) : '-'}</div><div class="k">First IN</div></div>
      <div class="stat"><div class="v">${fmtMin(stats.awayMin)}</div><div class="k">Time out today</div></div>
      <div class="stat"><div class="v">${stats.away.length}</div><div class="k">Times out</div></div>
    </div>
    <div class="section-title" style="margin-top:14px"><span>Today</span><span>${plural(list.length, 'entry', 'entries')}</span></div>
    ${list.length ? `<div class="entries">${timeline(list, k, st.att)}</div>` : '<div class="muted small" style="padding:6px 2px">No entries today.</div>'}
    <button class="btn ghost block" style="margin-top:14px" data-action="new-punch" data-id="${e.id}">Add a missed entry</button>`;
}

function punchRow(p) {
  return `<div class="entry">
    <span class="tag ${p.type}">${typeLabel(p.type)}</span>
    <span class="times"><b>${fmtTime(p.ts)}</b>${p.reason ? ` <span class="pill ${colorClass(p.reason)}"><span class="dot"></span>${esc(p.reason)}</span>` : ''}${p.note ? `<br><span class="small muted">${esc(p.note)}</span>` : ''}</span>
    <button class="edit" data-action="edit-punch" data-pid="${p.id}">Edit</button>
  </div>`;
}

// A day's punches in order, with "Out for 25m" between an OUT and the IN that follows it,
// and the time counted until closing after a final OUT.
function timeline(list, k, att) {
  const st = dayStats(list, k, att);
  const gaps = new Map(st.away.filter(a => a.back).map(a => [a.back.id, a.min]));
  let html = list.map(p => (gaps.has(p.id) ? `<div class="gap">Out for ${fmtMin(gaps.get(p.id))}</div>` : '') + punchRow(p)).join('');
  const tail = st.away.find(a => !a.back);
  if (tail) {
    html += Date.now() < st.close && k === todayKey()
      ? `<div class="gap">Out for ${fmtMin(tail.min)} so far</div>`
      : `<div class="gap">Not back before closing, counted until ${fmtClock(S.settings.closeMin)}: ${fmtMin(tail.min)}</div>`;
  } else if (st.openOut && att === 'half') {
    html += '<div class="gap muted-gap">Left for the half day, not counted</div>';
  }
  return html;
}

/* ---------- Modal (edit entry / edit employee) ---------- */
function renderModal() {
  const m = S.modal;
  const panel = $('#modalPanel');
  if (!m) { panel.innerHTML = ''; return; }
  if (m.type === 'punch') {
    const d = m.draft;
    const reasons = [...S.settings.reasons];
    if (d.reason && !reasons.includes(d.reason)) reasons.push(d.reason);
    panel.innerHTML = `
      <div class="sheet-head" style="margin-bottom:12px">
        <div><h2>${m.isNew ? 'Add missed entry' : 'Edit entry'}</h2><div class="small muted">${esc(empName(d.empId))}</div></div>
        <button class="icon-btn" data-action="close" aria-label="Close">${ICON.close}</button>
      </div>
      <form data-form="punch">
        <div class="field"><span>Type</span>
          <div class="seg">
            <label><input type="radio" name="type" value="in" ${d.type === 'in' ? 'checked' : ''} data-action="type-change"><span class="t-in">IN</span></label>
            <label><input type="radio" name="type" value="out" ${d.type === 'out' ? 'checked' : ''} data-action="type-change"><span class="t-out">OUT</span></label>
          </div>
        </div>
        <div class="field"><span>Date</span>
          ${dateField('date', dayKey(d.ts))}
          <div class="quick-days">
            <button type="button" class="mini-btn" data-action="set-date" data-key="${todayKey()}">Today</button>
            <button type="button" class="mini-btn" data-action="set-date" data-key="${shiftDay(todayKey(), -1)}">Yesterday</button>
          </div>
        </div>
        <div class="field"><span>Time</span>${timeField(d.ts)}</div>
        <label class="field ${d.type === 'in' ? 'hidden' : ''}" id="reasonField"><span>Reason (optional)</span>
          <select class="input" name="reason"><option value="">No reason</option>${reasons.map(r => `<option ${r === d.reason ? 'selected' : ''}>${esc(r)}</option>`).join('')}</select>
        </label>
        <label class="field"><span>Note (optional)</span><input class="input" name="note" value="${esc(d.note || '')}" maxlength="200" placeholder="e.g. went to the bank"></label>
        <p class="error-text hidden" id="formError"></p>
        <div class="btn-row">
          ${m.isNew ? '' : '<button type="button" class="btn danger" data-action="delete-punch">Delete</button>'}
          <button type="submit" class="btn">Save</button>
        </div>
      </form>`;
  } else if (m.type === 'employee') {
    const e = empById(m.id);
    panel.innerHTML = `
      <div class="sheet-head" style="margin-bottom:12px">
        <div><h2>Edit employee</h2></div>
        <button class="icon-btn" data-action="close" aria-label="Close">${ICON.close}</button>
      </div>
      <form data-form="employee">
        <label class="field"><span>Name</span><input class="input" name="name" value="${esc(e.name)}" required maxlength="60"></label>
        <label class="field"><span>Department / role (optional)</span><input class="input" name="dept" value="${esc(e.dept || '')}" maxlength="60"></label>
        <div class="btn-row">
          <button type="button" class="btn danger" data-action="remove-emp" data-id="${e.id}">Remove</button>
          <button type="submit" class="btn">Save</button>
        </div>
      </form>`;
  }
}

function openPunchEditor(punchId, empId) {
  const p = punchId ? S.punches.find(x => x.id === punchId) : null;
  if (p) {
    S.modal = { type: 'punch', isNew: false, id: p.id, draft: { ...p } };
  } else {
    const ts = Math.floor((Date.now() - 10 * 60000) / 60000) * 60000;
    S.modal = { type: 'punch', isNew: true, draft: { empId, type: statusOf(empId).type === 'in' ? 'out' : 'in', ts, reason: '', note: '' } };
  }
  openOverlay('modal');
}

async function submitPunch(form) {
  const m = S.modal;
  const fd = new FormData(form);
  const err = msg => { const el = $('#formError'); el.textContent = msg; el.classList.remove('hidden'); };
  const ts = readDateTime(fd);
  const type = fd.get('type');
  if (!type) return err('Choose IN or OUT.');
  if (!ts) return err('Please enter the date and time.');
  if (ts > Date.now() + 60000) return err('The time cannot be in the future.');
  const p = m.isNew ? { id: uid(), empId: m.draft.empId, deleted: false } : S.punches.find(x => x.id === m.id);
  Object.assign(p, { type, ts, reason: type === 'out' ? String(fd.get('reason') || '') : '', note: String(fd.get('note') || '').trim() });
  await savePunch(p);
  closeOverlay();
  toast(m.isNew ? 'Entry added' : 'Entry updated');
}

/* ---------- Reports ---------- */
function rangeKeys() {
  const now = new Date();
  const k = d => dayKey(d.getTime());
  const y = now.getFullYear(), mo = now.getMonth(), d = now.getDate();
  switch (S.report.range) {
    case 'yesterday': { const t = new Date(y, mo, d - 1); return [k(t), k(t)]; }
    case 'week': { const dow = (now.getDay() + 6) % 7; return [k(new Date(y, mo, d - dow)), k(now)]; }
    case 'month': return [k(new Date(y, mo, 1)), k(now)];
    case 'lastmonth': return [k(new Date(y, mo - 1, 1)), k(new Date(y, mo, 0))];
    case 'custom': {
      let f = S.report.from || k(now), t = S.report.to || f;
      if (t < f) [f, t] = [t, f];
      return [f, t];
    }
    default: return [k(now), k(now)];
  }
}
const rangeLabel = (f, t) => (f === t ? fmtLongDate(keyToTs(f)) : `${fmtDate(keyToTs(f))} to ${fmtDate(keyToTs(t))}`);

function reportData() {
  const [from, to] = rangeKeys();
  const inRange = k => k >= from && k <= to;
  const list = live().filter(p => inRange(dayKey(p.ts))).sort((a, b) => a.ts - b.ts);
  const groups = groupByEmpDay(list);
  // Days marked absent (or present / half day) with no entries still belong in the report.
  const recs = new Map();
  for (const r of S.days) if (!r.deleted && inRange(r.date)) {
    recs.set(r.id, r);
    if (!groups.has(r.id)) groups.set(r.id, []);
  }
  const reasonNames = [...S.settings.reasons];
  const byEmp = new Map();
  let totalAway = 0, totalOuts = 0, totalAbsent = 0;
  for (const [key, arr] of groups) {
    const [empId, k] = key.split('|');
    const att = attOf(recs.get(key), arr);
    if (!att) continue;
    const st = dayStats(arr, k, att);
    let row = byEmp.get(empId);
    if (!row) {
      const emp = empById(empId) || { id: empId, name: 'Unknown', dept: '' };
      row = { emp, days: [], awayMin: 0, outs: 0, present: 0, half: 0, absent: 0, reasons: {} };
      byEmp.set(empId, row);
    }
    row.days.push({ k, list: arr, st, att });
    row[att]++;
    row.awayMin += st.awayMin; row.outs += st.away.length;
    for (const a of st.away) {
      const r = a.out.reason || NO_REASON;
      if (!reasonNames.includes(r)) reasonNames.push(r);
      const x = row.reasons[r] || (row.reasons[r] = { count: 0, min: 0 });
      x.count++; x.min += a.min;
    }
    totalAway += st.awayMin; totalOuts += st.away.length;
    if (att === 'absent') totalAbsent++;
  }
  const rows = [...byEmp.values()].sort((a, b) => b.awayMin - a.awayMin || byName(a.emp, b.emp));
  rows.forEach(r => r.days.sort((a, b) => a.k.localeCompare(b.k)));
  // Keep "No reason" last.
  const nr = reasonNames.indexOf(NO_REASON);
  if (nr >= 0) { reasonNames.splice(nr, 1); reasonNames.push(NO_REASON); }
  return { from, to, list, rows, totalAway, totalOuts, totalAbsent, reasonNames };
}

function renderReports() {
  const d = reportData();
  const ranges = [['today', 'Today'], ['yesterday', 'Yesterday'], ['week', 'This week'], ['month', 'This month'], ['lastmonth', 'Last month'], ['custom', 'Pick dates']];
  const canShare = !!(navigator.canShare && window.File);
  let html = `<div class="chips" style="margin-bottom:12px">${ranges.map(([k, l]) =>
    `<button class="chip ${S.report.range === k ? 'active' : ''}" data-action="range" data-range="${k}">${l}</button>`).join('')}</div>`;
  if (S.report.range === 'custom') {
    html += `<div class="card"><div class="date-pair">
      <div class="field"><span>From</span>${dateField('from', d.from, { change: 'rep-from', short: true })}</div>
      <div class="field"><span>To</span>${dateField('to', d.to, { change: 'rep-to', short: true })}</div>
    </div></div>`;
  }
  html += `<div class="card">
    <div class="small muted">${esc(rangeLabel(d.from, d.to))}</div>
    <div class="kpis">
      <div><div class="kpi-v">${fmtMin(d.totalAway)}</div><div class="small muted">Total time out</div></div>
      <div><div class="kpi-v">${d.totalOuts}</div><div class="small muted">Times out</div></div>
      <div><div class="kpi-v">${d.totalAbsent}</div><div class="small muted">Absent days</div></div>
    </div>
    <div class="btn-row" style="margin-top:14px">
      <button class="btn" data-action="excel" ${d.rows.length ? '' : 'disabled'}>Download Excel</button>
      ${canShare ? `<button class="btn secondary" data-action="share-excel" ${d.rows.length ? '' : 'disabled'}>Share Excel</button>` : ''}
    </div>
  </div>`;

  if (!d.rows.length) {
    html += '<div class="empty">No entries in this period.</div>';
  } else {
    html += '<div class="list">' + d.rows.map(row => {
      const open = S.report.open === row.emp.id;
      const attPills = [['present', 'Present'], ['half', 'Half day'], ['absent', 'Absent']]
        .filter(([key]) => row[key]).map(([key, label]) => `<span class="att-tag ${key}">${label} ${row[key]}</span>`).join('');
      const pills = attPills + d.reasonNames.filter(r => row.reasons[r]).map(r =>
        `<span class="pill ${colorClass(r)}"><span class="dot"></span>${esc(r)} ${row.reasons[r].count}x, ${fmtMin(row.reasons[r].min)}</span>`).join('');
      let body = '';
      if (open) {
        body = '<div class="rep-body">' + row.days.map(({ k, list, st, att }) => `
          <div class="day-head">
            <span class="day-label">${fmtDay(keyToTs(k))} <span class="att-tag ${att}">${STATUS_LABEL[att]}</span></span>
            <span class="small muted">${att === 'absent' ? '' : `${st.firstIn ? `IN ${fmtTime(st.firstIn)}` : 'No IN'}, out ${fmtMin(st.awayMin)}`}</span>
          </div>${timeline(list, k, att)}`).join('') + '</div>';
      }
      return `<div class="rep-row">
        <button class="rep-head" data-action="toggle-rep" data-id="${row.emp.id}">
          <span class="avatar">${esc(initials(row.emp.name))}</span>
          <span class="row-text"><span class="name">${esc(row.emp.name)}</span>
            <span class="rep-pills">${pills}</span>
          </span>
          <span class="rep-total"><div class="v">${fmtMin(row.awayMin)}</div><div class="k">${plural(row.outs, 'time')} out</div></span>
        </button>${body}</div>`;
    }).join('') + '</div>';
  }
  $('#view').innerHTML = html;
}

/* ---------- Excel export ---------- */
function buildWorkbook() {
  if (!window.XLSX) throw new Error('Excel library not loaded yet. Please try again in a moment.');
  const d = reportData();
  const R = d.reasonNames.filter(r => d.rows.some(row => row.reasons[r]));
  const wb = XLSX.utils.book_new();
  const title = `In-Out Register: ${rangeLabel(d.from, d.to)}`;
  const round = m => Math.round(m);
  const withFilter = (ws, firstRow, nRows, nCols) => {
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: firstRow, c: 0 }, e: { r: firstRow + nRows, c: nCols - 1 } }) };
  };

  // Summary
  const hours = m => Math.round(m / 60 * 100) / 100;
  const head = ['Employee', 'Department', 'Present', 'Half days', 'Absent', 'Times out', 'Time out', 'Hours out', 'Minutes out', ...R.map(r => `${r} (minutes)`)];
  const rows = d.rows.map(row => [
    row.emp.name, row.emp.dept || '', row.present, row.half, row.absent, row.outs, fmtMin(row.awayMin), hours(row.awayMin), round(row.awayMin),
    ...R.map(r => round(row.reasons[r]?.min || 0)),
  ]);
  const sum = i => rows.reduce((a, r) => a + (typeof r[i] === 'number' ? r[i] : 0), 0);
  const totalRow = head.map((h, i) => (i === 0 ? 'TOTAL' : i === 1 ? '' : i === 6 ? fmtMin(d.totalAway) : i === 7 ? hours(d.totalAway) : sum(i)));
  const wsSum = XLSX.utils.aoa_to_sheet([[title], [`Exported ${fmtDate(Date.now())} at ${fmtTime(Date.now())}`], [], head, ...rows, totalRow]);
  wsSum['!cols'] = head.map((h, i) => ({ wch: i === 0 ? 24 : Math.max(12, h.length + 2) }));
  wsSum['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: Math.min(head.length - 1, 6) } }];
  withFilter(wsSum, 3, rows.length, head.length);
  XLSX.utils.book_append_sheet(wb, wsSum, 'Summary');

  // Daily
  const dHead = ['Date', 'Day', 'Employee', 'Department', 'Attendance', 'First IN', 'Last OUT', 'Times out', 'Time out', 'Minutes out', 'Reasons'];
  const daily = [];
  for (const row of d.rows) for (const { k, st, att } of row.days) daily.push({ k, row, st, att });
  daily.sort((a, b) => a.k.localeCompare(b.k) || byName(a.row.emp, b.row.emp));
  const dRows = daily.map(({ k, row, st, att }) => {
    const reasons = {};
    st.away.forEach(a => { const r = a.out.reason || NO_REASON; reasons[r] = (reasons[r] || 0) + a.min; });
    return [
      fmtDate(keyToTs(k)), new Date(keyToTs(k)).toLocaleDateString('en-GB', { weekday: 'short' }), row.emp.name, row.emp.dept || '',
      STATUS_LABEL[att], st.firstIn ? fmtTime(st.firstIn) : '', st.leftAt ? fmtTime(st.leftAt) : '', st.away.length, fmtMin(st.awayMin), round(st.awayMin),
      Object.entries(reasons).map(([r, m]) => `${r} ${fmtMin(m)}`).join(', '),
    ];
  });
  const wsDay = XLSX.utils.aoa_to_sheet([dHead, ...dRows]);
  wsDay['!cols'] = [13, 6, 24, 16, 12, 11, 11, 10, 11, 12, 40].map(w => ({ wch: w }));
  withFilter(wsDay, 0, dRows.length, dHead.length);
  XLSX.utils.book_append_sheet(wb, wsDay, 'Daily');

  // Time out (each OUT and the IN that followed it)
  const aHead = ['Date', 'Employee', 'Department', 'Out', 'Back', 'Minutes', 'Reason', 'Note'];
  const away = [];
  for (const row of d.rows) for (const { st } of row.days) for (const a of st.away) away.push({ row, a });
  away.sort((x, y) => x.a.out.ts - y.a.out.ts);
  const aRows = away.map(({ row, a }) => [
    fmtDate(a.out.ts), row.emp.name, row.emp.dept || '', fmtTime(a.out.ts),
    a.back ? fmtTime(a.back.ts) : `Closing ${fmtClock(S.settings.closeMin)}`, round(a.min),
    a.out.reason || '', [a.out.note, a.back?.note, a.back ? '' : 'Not back before closing'].filter(Boolean).join('; '),
  ]);
  const wsAway = XLSX.utils.aoa_to_sheet([aHead, ...aRows]);
  wsAway['!cols'] = [13, 24, 16, 11, 15, 9, 16, 32].map(w => ({ wch: w }));
  withFilter(wsAway, 0, aRows.length, aHead.length);
  XLSX.utils.book_append_sheet(wb, wsAway, 'Time Out');

  // Every punch
  const lHead = ['Date', 'Time', 'Employee', 'Department', 'IN / OUT', 'Reason', 'Note'];
  const lRows = d.list.map(p => {
    const e = empById(p.empId) || { name: 'Unknown' };
    return [fmtDate(p.ts), fmtTime(p.ts), e.name, e.dept || '', typeLabel(p.type), p.reason || '', p.note || ''];
  });
  const wsLog = XLSX.utils.aoa_to_sheet([lHead, ...lRows]);
  wsLog['!cols'] = [13, 11, 24, 16, 9, 16, 32].map(w => ({ wch: w }));
  withFilter(wsLog, 0, lRows.length, lHead.length);
  XLSX.utils.book_append_sheet(wb, wsLog, 'Punch Log');

  const name = `In-Out_${d.from}${d.from !== d.to ? '_to_' + d.to : ''}.xlsx`;
  return { wb, name };
}

async function exportExcel(share) {
  try {
    const { wb, name } = buildWorkbook();
    if (share) {
      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const file = new File([buf], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      if (navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: name });
        return;
      }
    }
    XLSX.writeFile(wb, name);
    toast('Excel downloaded');
  } catch (err) {
    if (err.name !== 'AbortError') toast(err.message || 'Could not create the Excel file');
  }
}

/* ---------- Employees ---------- */
function renderEmployees() {
  const active = activeEmps();
  const removed = S.employees.filter(e => !e.active).sort(byName);
  $('#view').innerHTML = `
    <div class="card">
      <h2>Add employee</h2>
      <form data-form="add-emp" style="margin-top:10px">
        <label class="field"><span>Name</span><input class="input" name="name" required maxlength="60" autocomplete="off" placeholder="Full name"></label>
        <label class="field"><span>Department / role (optional)</span><input class="input" name="dept" maxlength="60" autocomplete="off" placeholder="e.g. Accounts"></label>
        <button class="btn block" type="submit">Add employee</button>
      </form>
      <details style="margin-top:12px">
        <summary>Add many at once</summary>
        <form data-form="bulk-emp" style="margin-top:8px">
          <label class="field"><span>One name per line</span><textarea class="input" name="names" placeholder="Ravi Kumar&#10;Priya Sharma&#10;Anil"></textarea></label>
          <button class="btn secondary block" type="submit">Add all</button>
        </form>
      </details>
      <details style="margin-top:4px">
        <summary>Import from Excel</summary>
        <p class="small muted">Names in the first column, department (optional) in the second. A header row is skipped automatically.</p>
        <button class="btn secondary block" data-action="import-emps">Choose Excel or CSV file</button>
        <input type="file" id="importInput" accept=".xlsx,.xls,.csv" class="hidden">
      </details>
    </div>
    <div class="section-title"><span>Employees</span><span>${active.length}</span></div>
    <div class="card" style="padding:4px 16px">
      ${active.length ? `<div class="plain-list">${active.map(e => `
        <div class="plain-item">
          <span class="avatar">${esc(initials(e.name))}</span>
          <span class="row-text"><span class="name">${esc(e.name)}</span><span class="sub">${esc(e.dept || '')}</span></span>
          <button class="mini-btn" data-action="edit-emp" data-id="${e.id}">Edit</button>
        </div>`).join('')}</div>` : '<div class="empty">No employees yet.</div>'}
    </div>
    ${removed.length ? `<details class="card"><summary>Removed employees (${removed.length})</summary>
      <div class="plain-list">${removed.map(e => `
        <div class="plain-item"><span class="row-text"><span class="name">${esc(e.name)}</span></span>
        <button class="mini-btn" data-action="restore-emp" data-id="${e.id}">Restore</button></div>`).join('')}</div>
      <p class="small muted">Removed employees are hidden from the main screen but their past entries stay in reports.</p>
    </details>` : ''}`;
}

async function addEmployees(list) {
  const existing = new Set(S.employees.filter(e => e.active).map(e => e.name.toLowerCase()));
  let added = 0, skipped = 0;
  for (const raw of list) {
    const name = String(raw.name || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    if (!name) continue;
    if (existing.has(name.toLowerCase())) { skipped++; continue; }
    existing.add(name.toLowerCase());
    await saveEmployee({ id: uid(), name, dept: String(raw.dept || '').trim().slice(0, 60), active: true, createdAt: Date.now() });
    added++;
  }
  toast(`${plural(added, 'employee')} added${skipped ? `, ${skipped} already existed` : ''}`);
}

async function importEmployees(file) {
  try {
    if (!window.XLSX) throw new Error('Excel library not loaded yet. Please try again in a moment.');
    const wb = XLSX.read(await file.arrayBuffer());
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, blankrows: false });
    const start = rows[0] && /name/i.test(String(rows[0][0] ?? '')) ? 1 : 0;
    const list = rows.slice(start)
      .map(r => ({ name: String(r[0] ?? '').trim(), dept: String(r[1] ?? '').trim() }))
      .filter(x => x.name);
    if (!list.length) { toast('No names found in the first column'); return; }
    if (!confirm(`Add ${plural(list.length, 'employee')} from ${file.name}?`)) return;
    await addEmployees(list);
    render();
  } catch (err) { toast(err.message || 'Could not read the file'); }
}

/* ---------- Settings ---------- */
function renderSettings() {
  const st = S.settings;
  $('#view').innerHTML = `
    <div class="card">
      <h2>Office hours</h2>
      <p class="small muted">Marking Present records IN at the opening time. If someone goes OUT and does not come back, the time out is counted until closing.</p>
      <form data-form="hours">
        <div class="field"><span>Opens</span>${timeField(minToTs(todayKey(), st.openMin), 'o')}</div>
        <div class="field"><span>Closes</span>${timeField(minToTs(todayKey(), st.closeMin), 'c')}</div>
        <p class="error-text hidden" id="hoursError"></p>
        <button type="submit" class="btn block">Save office hours</button>
      </form>
    </div>

    <div class="card">
      <h2>Reasons for going out</h2>
      <p class="small muted">Shown as optional choices when someone goes OUT. You can always go OUT without picking one.</p>
      <form data-form="reasons">
        <div id="reasonRows">${st.reasons.map(reasonEditRow).join('')}</div>
        <div class="btn-row" style="margin-top:6px">
          <button type="button" class="btn ghost" data-action="add-reason">Add reason</button>
          <button type="submit" class="btn">Save reasons</button>
        </div>
      </form>
    </div>

    <div class="card">
      <h2>Google Sheet backup</h2>
      <p class="small muted">Every entry is copied to your Google Sheet automatically, so nothing is lost if this phone is lost or reset. Entries made without internet are sent when it comes back.</p>
      <div id="syncStatus">${syncStatusHtml()}</div>
      <form data-form="sync">
        <label class="field"><span>Web app URL</span><input class="input" name="url" value="${esc(st.scriptUrl)}" placeholder="https://script.google.com/macros/s/.../exec" autocomplete="off" inputmode="url"></label>
        <label class="field"><span>Password (same as SECRET in the script)</span><input class="input" name="secret" value="${esc(st.secret)}" autocomplete="off"></label>
        <div class="btn-row">
          <button type="submit" class="btn">Save and test</button>
          <button type="button" class="btn secondary" data-action="sync-now" ${st.scriptUrl ? '' : 'disabled'}>Back up now</button>
        </div>
      </form>
      <details style="margin-top:12px">
        <summary>Restore from Google Sheet</summary>
        <p class="small muted">Use this on a new phone. Employees and entries from the sheet are merged into this phone. Nothing here is deleted.</p>
        <button class="btn secondary block" data-action="restore-sheet" ${st.scriptUrl ? '' : 'disabled'}>Restore from Google Sheet</button>
      </details>
    </div>

    <div class="card">
      <h2>Backup file</h2>
      <p class="small muted">Save a full copy of all data as a file, or load one back.</p>
      <div class="btn-row">
        <button class="btn secondary" data-action="backup-file">Save backup file</button>
        <button class="btn ghost" data-action="restore-file">Load backup file</button>
      </div>
      <input type="file" id="restoreInput" accept=".json,application/json" class="hidden">
    </div>

    <div class="card">
      <label class="switch"><span><b>Keep screen on</b><br><span class="small muted">Stops the phone from locking while the app is open</span></span>
        <input type="checkbox" data-change="keep-awake" ${st.keepAwake ? 'checked' : ''}></label>
    </div>

    <p class="small muted" style="text-align:center">In-Out Register ${APP_VERSION}. ${plural(activeEmps().length, 'employee')}, ${plural(live().length, 'entry', 'entries')} stored on this phone.</p>`;
}

function reasonEditRow(name = '') {
  return `<div class="reason-edit" data-orig="${esc(name)}">
    <input class="input" name="rname" value="${esc(name)}" maxlength="24" placeholder="Reason">
    <button type="button" class="x" data-action="remove-reason" aria-label="Remove">&times;</button>
  </div>`;
}

async function submitReasons(form) {
  const reasons = [];
  const renames = [];
  for (const row of form.querySelectorAll('.reason-edit')) {
    const name = row.querySelector('[name=rname]').value.trim();
    if (!name) continue;
    if (name === NO_REASON) { toast(`"${NO_REASON}" is used by the app, please pick another name`); return; }
    if (reasons.some(r => r.toLowerCase() === name.toLowerCase())) { toast(`"${name}" is listed twice`); return; }
    reasons.push(name);
    const orig = row.dataset.orig;
    if (orig && orig !== name) renames.push([orig, name]);
  }
  S.settings.reasons = reasons;
  await saveSettings();
  for (const [from, to] of renames) {
    for (const p of S.punches.filter(x => x.reason === from)) { p.reason = to; await savePunch(p); }
  }
  toast('Reasons saved');
  render();
}

/* ============================================================
   Google Sheet sync
   ============================================================ */
let syncTimer = null;
const pendingCount = () => S.employees.filter(e => e.dirty).length + S.punches.filter(p => p.dirty).length + S.days.filter(d => d.dirty).length;

function queueSync(delay = 1500) {
  updateBadge();
  if (!S.settings.scriptUrl) return;
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => syncNow(false), delay);
}

const empPayload = e => ({ id: e.id, name: e.name, dept: e.dept || '', active: !!e.active, createdAt: e.createdAt || 0, updatedAt: e.updatedAt || 0 });
const dayPayload = r => ({ id: r.id, deleted: !!r.deleted, empId: r.empId, date: r.date, status: r.status, updatedAt: r.updatedAt || 0 });
const punchPayload = p => ({
  id: p.id, deleted: !!p.deleted, empId: p.empId, type: p.type, ts: p.ts,
  reason: p.reason || '', note: p.note || '', updatedAt: p.updatedAt || 0,
});

// Always POST so the password travels in the request body, never in a URL.
async function callScript(payload) {
  const { scriptUrl, secret } = S.settings;
  // text/plain keeps this a "simple" request, which Apps Script accepts without CORS preflight.
  const res = await fetch(scriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ ...payload, secret }),
  });
  let data;
  try { data = await res.json(); } catch { throw new Error('Unexpected reply. Check the web app URL and that access is set to "Anyone".'); }
  if (!data.ok) throw new Error(data.error || 'The sheet refused the request');
  return data;
}

async function syncNow(manual) {
  if (!S.settings.scriptUrl) { if (manual) toast('Set up Google Sheet backup in Settings first'); return; }
  if (S.sync.busy) return;
  if (!navigator.onLine) { S.sync.error = 'offline'; updateBadge(); if (manual) toast('No internet. Entries are safe on this phone and will be backed up later.'); return; }
  S.sync.busy = true; S.sync.error = ''; updateBadge();
  try {
    for (let round = 0; round < 100; round++) {
      const emps = S.employees.filter(e => e.dirty);
      const punches = S.punches.filter(p => p.dirty).slice(0, 500);
      const days = S.days.filter(r => r.dirty).slice(0, 500);
      if (!emps.length && !punches.length && !days.length) break;
      const stamp = new Map([...emps, ...punches, ...days].map(r => [r.id, r.updatedAt]));
      await callScript({
        action: 'sync',
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        reasons: S.settings.reasons,
        employees: emps.map(empPayload),
        punches: punches.map(punchPayload),
        attendance: days.map(dayPayload),
        office: { open: S.settings.openMin, close: S.settings.closeMin },
      });
      // Anything edited while the request was in flight stays dirty for the next round.
      const doneE = [], doneP = [], purge = [];
      emps.forEach(e => { if (e.updatedAt === stamp.get(e.id)) { e.dirty = false; doneE.push(e); } });
      punches.forEach(p => {
        if (p.updatedAt !== stamp.get(p.id)) return;
        if (p.deleted) purge.push(p.id); else { p.dirty = false; doneP.push(p); }
      });
      if (doneE.length) await DB.put('employees', doneE);
      if (doneP.length) await DB.put('punches', doneP);
      if (purge.length) { await DB.del('punches', purge); S.punches = S.punches.filter(p => !purge.includes(p.id)); }
      const doneD = [], purgeD = [];
      days.forEach(r => {
        if (r.updatedAt !== stamp.get(r.id)) return;
        if (r.deleted) purgeD.push(r.id); else { r.dirty = false; doneD.push(r); }
      });
      if (doneD.length) await DB.put('days', doneD);
      if (purgeD.length) { await DB.del('days', purgeD); S.days = S.days.filter(r => !purgeD.includes(r.id)); }
      if (!doneE.length && !doneP.length && !purge.length && !doneD.length && !purgeD.length) break;
    }
    S.settings.lastSync = Date.now();
    await saveSettings();
    if (manual) toast('Backed up to Google Sheet');
  } catch (err) {
    S.sync.error = err.message || String(err);
    if (manual) toast('Backup failed: ' + S.sync.error);
  } finally {
    S.sync.busy = false;
    updateBadge();
  }
}

function syncState() {
  const n = pendingCount();
  if (!S.settings.scriptUrl) return { cls: 'off', label: 'Phone only', text: 'Not connected. Data is stored on this phone only.' };
  if (S.sync.busy) return { cls: 'pending', label: 'Saving', text: 'Backing up now...' };
  if (S.sync.error === 'offline' || !navigator.onLine) return { cls: 'pending', label: n ? `Offline, ${n} waiting` : 'Offline', text: `No internet. ${plural(n, 'change')} waiting; they will be sent automatically.` };
  if (S.sync.error) return { cls: 'error', label: 'Backup failed', text: `Last backup failed: ${S.sync.error}` };
  if (n) return { cls: 'pending', label: `${n} waiting`, text: `${plural(n, 'change')} waiting to be backed up.` };
  return { cls: 'ok', label: 'Backed up', text: `Everything is backed up${S.settings.lastSync ? ` (last at ${fmtTime(S.settings.lastSync)}, ${fmtDay(S.settings.lastSync)})` : ''}.` };
}
const syncStatusHtml = () => { const s = syncState(); return `<div class="status-line ${s.cls}"><span class="dot"></span>${esc(s.text)}</div>`; };

function updateBadge() {
  const s = syncState();
  const b = $('#syncBadge');
  b.className = `sync-badge ${s.cls}`;
  b.innerHTML = `<span class="dot"></span>${esc(s.label)}`;
  const box = $('#syncStatus');
  if (box) box.innerHTML = syncStatusHtml();
}

async function mergeIncoming(emps, punches, markDirty, days = []) {
  let nE = 0, nP = 0;
  const putE = [], putP = [];
  for (const inc of emps || []) {
    if (!inc || !inc.id || !inc.name) continue;
    const cur = empById(inc.id);
    if (cur && (cur.updatedAt || 0) >= (inc.updatedAt || 0)) continue;
    const e = { id: inc.id, name: inc.name, dept: inc.dept || '', active: inc.active !== false, createdAt: Number(inc.createdAt) || Date.now(), updatedAt: Number(inc.updatedAt) || Date.now(), dirty: markDirty };
    if (cur) Object.assign(cur, e); else S.employees.push(e);
    putE.push(cur || e); nE++;
  }
  for (const inc of punches || []) {
    if (!inc || !inc.id || !inc.empId || !inc.ts || inc.deleted || (inc.type !== 'in' && inc.type !== 'out')) continue;
    const cur = S.punches.find(p => p.id === inc.id);
    if (cur && (cur.updatedAt || 0) >= (inc.updatedAt || 0)) continue;
    const p = { id: inc.id, empId: inc.empId, type: inc.type, ts: Number(inc.ts), reason: inc.reason || '', note: inc.note || '', deleted: false, updatedAt: Number(inc.updatedAt) || Date.now(), dirty: markDirty };
    if (cur) Object.assign(cur, p); else S.punches.push(p);
    putP.push(cur || p); nP++;
  }
  const putD = [];
  for (const inc of days || []) {
    if (!inc || !inc.empId || !/^\d{4}-\d{2}-\d{2}$/.test(String(inc.date)) || inc.deleted || !STATUS_LABEL[inc.status]) continue;
    const id = `${inc.empId}|${inc.date}`;
    const cur = S.days.find(r => r.id === id);
    if (cur && (cur.updatedAt || 0) >= (inc.updatedAt || 0)) continue;
    const r = { id, empId: inc.empId, date: String(inc.date), status: inc.status, deleted: false, updatedAt: Number(inc.updatedAt) || Date.now(), dirty: markDirty };
    if (cur) Object.assign(cur, r); else S.days.push(r);
    putD.push(cur || r);
  }
  if (putE.length) await DB.put('employees', putE);
  if (putP.length) await DB.put('punches', putP);
  if (putD.length) await DB.put('days', putD);
  if (markDirty) queueSync();
  return { nE, nP };
}

/* ---------- Backup file ---------- */
function backupFile() {
  const data = {
    app: 'inout-register', version: APP_VERSION, exportedAt: new Date().toISOString(),
    settings: { reasons: S.settings.reasons },
    employees: S.employees.map(({ dirty, ...e }) => e),
    punches: live().map(({ dirty, ...p }) => p),
    days: S.days.filter(r => !r.deleted).map(({ dirty, ...r }) => r),
  };
  const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `In-Out-backup_${todayKey()}.json`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

async function restoreFile(file) {
  try {
    const data = JSON.parse(await file.text());
    if (data.app !== 'inout-register') throw new Error('This is not an In-Out Register backup file');
    if (!confirm(`Load ${plural(data.employees?.length || 0, 'employee')} and ${plural(data.punches?.length || 0, 'entry', 'entries')} from this file? They will be merged with what is on this phone.`)) return;
    if (Array.isArray(data.settings?.reasons)) {
      for (const r of data.settings.reasons) if (!S.settings.reasons.includes(r)) S.settings.reasons.push(r);
      await saveSettings();
    }
    const { nE, nP } = await mergeIncoming(data.employees, data.punches, true, data.days);
    toast(`Loaded ${plural(nE, 'employee')} and ${plural(nP, 'entry', 'entries')}`);
    render();
  } catch (err) { toast(err.message || 'Could not read the file'); }
}

/* ---------- Keep screen awake ---------- */
let wakeLock = null;
async function applyWakeLock() {
  try {
    if (S.settings.keepAwake && 'wakeLock' in navigator && document.visibilityState === 'visible') {
      if (!wakeLock) {
        wakeLock = await navigator.wakeLock.request('screen');
        wakeLock.addEventListener('release', () => { wakeLock = null; });
      }
    } else if (wakeLock) { await wakeLock.release(); wakeLock = null; }
  } catch { /* not supported or denied */ }
}

/* ---------- Clock & live timers ---------- */
let lastDay = todayKey();
function tick() {
  const now = Date.now();
  $('#clock').textContent = new Date(now).toLocaleTimeString('en-IN', { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true }).toUpperCase();
  $('#today').textContent = new Date(now).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long' });
  document.querySelectorAll('[data-since]').forEach(el => { el.textContent = fmtElapsed(now - Number(el.dataset.since)); });
  if (todayKey() !== lastDay) { lastDay = todayKey(); render(); }
}

/* ============================================================
   Events
   ============================================================ */
const actions = {
  nav: el => { S.view = el.dataset.view; window.scrollTo(0, 0); render(); },
  filter: el => { S.filter = S.filter === el.dataset.filter ? 'all' : el.dataset.filter; renderHomeBody(); tick(); },
  'open-emp': el => { S.sheetEmp = el.dataset.id; S.pickReason = ''; openOverlay('sheet'); tick(); },
  punch: el => punch(el.dataset.id, el.dataset.type, el.dataset.reason || ''),
  mark: el => markAttendance(el.dataset.id, el.dataset.status),
  'all-present': () => markAllPresent(),
  'set-att': el => {
    const st = statusOf(el.dataset.id);
    if (st.att === el.dataset.status) return;
    markAttendance(el.dataset.id, el.dataset.status);
  },
  'set-reason': async el => {
    const r = el.dataset.reason;
    const st = statusOf(S.sheetEmp);
    if (st.type === 'out') {
      st.punch.reason = st.punch.reason === r ? '' : r;
      await savePunch(st.punch);
      render();
    } else {
      S.pickReason = S.pickReason === r ? '' : r;
      renderSheet(); tick();
    }
  },
  close: () => closeOverlay(),
  undo: async () => {
    const f = S.undo; S.undo = null;
    $('#toast').classList.add('hidden');
    if (f) { await f(); toast('Undone'); }
  },
  'edit-punch': el => openPunchEditor(el.dataset.pid),
  'new-punch': el => openPunchEditor(null, el.dataset.id),
  'open-date': el => {
    const input = el.querySelector('input[type=date]');
    try { input.showPicker(); } catch { input.focus(); input.click(); }
  },
  'set-date': el => {
    const input = el.closest('.field').querySelector('input[type=date]');
    input.value = el.dataset.key;
    input.dispatchEvent(new Event('change', { bubbles: true }));
  },
  'type-change': el => { $('#reasonField').classList.toggle('hidden', el.value === 'in'); },
  'delete-punch': async () => {
    const p = S.punches.find(x => x.id === S.modal.id);
    if (!p || !confirm('Delete this entry permanently?')) return;
    p.deleted = true; await savePunch(p);
    closeOverlay(); toast('Entry deleted');
  },
  range: el => { S.report.range = el.dataset.range; S.report.open = null; renderReports(); },
  'toggle-rep': el => { S.report.open = S.report.open === el.dataset.id ? null : el.dataset.id; renderReports(); },
  excel: () => exportExcel(false),
  'share-excel': () => exportExcel(true),
  'edit-emp': el => { S.modal = { type: 'employee', id: el.dataset.id }; openOverlay('modal'); },
  'remove-emp': async el => {
    const e = empById(el.dataset.id);
    if (!e || !confirm(`Remove ${e.name}? Their past entries stay in reports.`)) return;
    e.active = false; await saveEmployee(e);
    closeOverlay(); toast(`${e.name} removed`);
  },
  'restore-emp': async el => { const e = empById(el.dataset.id); e.active = true; await saveEmployee(e); render(); toast(`${e.name} restored`); },
  'import-emps': () => $('#importInput').click(),
  'add-reason': () => { $('#reasonRows').insertAdjacentHTML('beforeend', reasonEditRow()); $('#reasonRows .reason-edit:last-child input').focus(); },
  'remove-reason': el => el.closest('.reason-edit').remove(),
  'sync-now': () => syncNow(true),
  badge: () => { if (S.settings.scriptUrl) syncNow(true); else { S.view = 'settings'; render(); } },
  'restore-sheet': async () => {
    if (!confirm('Download all employees and entries from the Google Sheet and merge them into this phone?')) return;
    try {
      toast('Restoring...');
      const data = await callScript({ action: 'restore' });
      if (Array.isArray(data.reasons)) {
        for (const r of data.reasons) if (!S.settings.reasons.includes(r)) S.settings.reasons.push(r);
        await saveSettings();
      }
      if (data.office && Number.isFinite(data.office.open) && Number.isFinite(data.office.close)) {
        S.settings.openMin = data.office.open; S.settings.closeMin = data.office.close;
        await saveSettings();
      }
      const { nE, nP } = await mergeIncoming(data.employees, data.punches, false, data.attendance);
      toast(`Restored ${plural(nE, 'employee')} and ${plural(nP, 'entry', 'entries')}`);
      render();
    } catch (err) { toast('Restore failed: ' + err.message); }
  },
  'backup-file': () => backupFile(),
  'restore-file': () => $('#restoreInput').click(),
};

document.addEventListener('click', e => {
  const overlay = e.target.classList && e.target.classList.contains('overlay') ? e.target : null;
  if (overlay) { closeOverlay(); return; }
  const el = e.target.closest('[data-action]');
  if (!el) return;
  const fn = actions[el.dataset.action];
  if (fn) fn(el, e);
});

document.addEventListener('input', e => {
  if (e.target.id === 'search') { S.search = e.target.value; renderHomeBody(); tick(); }
});

document.addEventListener('change', async e => {
  const t = e.target;
  if (t.id === 'restoreInput' && t.files[0]) { await restoreFile(t.files[0]); t.value = ''; return; }
  if (t.id === 'importInput' && t.files[0]) { await importEmployees(t.files[0]); t.value = ''; return; }
  switch (t.dataset.change) {
    case 'date-label':
      if (t.value) t.closest('.date-btn').querySelector('.date-text').textContent = fmtDateFull(keyToTs(t.value));
      break;
    case 'rep-from': S.report.from = t.value || S.report.from; renderReports(); break;
    case 'rep-to': S.report.to = t.value || S.report.to; renderReports(); break;
    case 'keep-awake': S.settings.keepAwake = t.checked; await saveSettings(); applyWakeLock(); break;
  }
});

document.addEventListener('submit', async e => {
  const form = e.target;
  if (!form.dataset.form) return;
  e.preventDefault();
  const fd = new FormData(form);
  switch (form.dataset.form) {
    case 'punch': await submitPunch(form); break;
    case 'employee': {
      const emp = empById(S.modal.id);
      const name = String(fd.get('name')).trim().replace(/\s+/g, ' ');
      if (!name) return;
      emp.name = name; emp.dept = String(fd.get('dept') || '').trim();
      await saveEmployee(emp);
      closeOverlay(); toast('Employee updated');
      break;
    }
    case 'add-emp':
      await addEmployees([{ name: String(fd.get('name')), dept: String(fd.get('dept') || '') }]);
      render();
      $('#view input[name=name]')?.focus();
      break;
    case 'bulk-emp':
      await addEmployees(String(fd.get('names')).split(/\r?\n/).map(name => ({ name })));
      render();
      break;
    case 'reasons': await submitReasons(form); break;
    case 'attend': {
      const status = e.submitter?.value || 'present';
      await markAttendance(S.sheetEmp, status, readTime(fd, 'a'));
      break;
    }
    case 'hours': {
      const open = readTime(fd, 'o'), close = readTime(fd, 'c');
      if (close <= open) {
        const el = $('#hoursError'); el.textContent = 'Closing time must be after opening time.'; el.classList.remove('hidden');
        break;
      }
      S.settings.openMin = open; S.settings.closeMin = close;
      await saveSettings();
      queueSync();
      toast(`Office hours saved: ${fmtClock(open)} to ${fmtClock(close)}`);
      render();
      break;
    }
    case 'sync': {
      S.settings.scriptUrl = String(fd.get('url')).trim();
      S.settings.secret = String(fd.get('secret')).trim();
      S.sync.error = '';
      await saveSettings();
      if (!S.settings.scriptUrl) { toast('Google Sheet backup turned off'); render(); break; }
      try {
        toast('Testing connection...');
        await callScript({ action: 'ping' });
        // First connection: send everything so the sheet has a full copy.
        S.employees.forEach(x => { x.dirty = true; });
        S.punches.forEach(x => { x.dirty = true; });
        S.days.forEach(x => { x.dirty = true; });
        await DB.put('days', S.days);
        await DB.put('employees', S.employees);
        await DB.put('punches', S.punches);
        toast('Connected. Backing up all data...');
        render();
        await syncNow(true);
      } catch (err) {
        S.sync.error = err.message;
        toast('Connection failed: ' + err.message);
        render();
      }
      break;
    }
  }
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') { applyWakeLock(); render(); queueSync(300); }
});
window.addEventListener('online', () => { S.sync.error = ''; queueSync(300); });
window.addEventListener('offline', () => updateBadge());

/* ---------- Start ---------- */
async function init() {
  try {
    await DB.open();
    const [emps, punches, days, meta] = await Promise.all([DB.all('employees'), DB.all('punches'), DB.all('days'), DB.all('meta')]);
    S.employees = emps;
    S.punches = punches;
    S.days = days;
    const saved = meta.find(m => m.key === 'settings');
    if (saved) S.settings = { ...structuredClone(DEFAULT_SETTINGS), ...saved.value };
    // Deleted entries are only kept until the sheet hears about them.
    if (!S.settings.scriptUrl) {
      const gone = S.punches.filter(p => p.deleted).map(p => p.id);
      if (gone.length) { await DB.del('punches', gone); S.punches = S.punches.filter(p => !p.deleted); }
      const goneD = S.days.filter(r => r.deleted).map(r => r.id);
      if (goneD.length) { await DB.del('days', goneD); S.days = S.days.filter(r => !r.deleted); }
    }
  } catch (err) {
    $('#view').innerHTML = `<div class="card empty"><h3>Storage is not available</h3><p>${esc(err.message || err)}</p><p>Open this page in Chrome (not in private / incognito mode).</p></div>`;
    return;
  }
  render();
  setInterval(tick, 1000);
  setInterval(() => { if (pendingCount()) queueSync(0); }, 60000);
  queueSync(800);
  applyWakeLock();
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
}
init();
