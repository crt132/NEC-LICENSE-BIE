/**
 * NEC Industrial Engineering Practice - shared backend
 * (batch leaderboard + reported questions) for the GitHub Pages copy.
 *
 * Setup: open a new Google Sheet > Extensions > Apps Script,
 * replace everything with this file, then Deploy > New deployment >
 * Web app (Execute as: Me, Who has access: Anyone). Copy the /exec URL
 * into BACKEND_URL in index.html.
 */
const BOARD = 'board', REPORTS = 'reports';
const BOARD_HEAD = ['id', 'name', 'sets', 'pts', 'best', 'avg', 'at'];
const REP_HEAD = ['id', 'k', 'r', 'n', 'at'];
const REASONS = ['key', 'q', 'opt', 'dup', 'other'];
const VISITS = 'visits';
const VISIT_HEAD = ['id', 'first', 'last', 'count'];
const NPT = 345 * 60000; // Nepal time offset

function visitStats_() {
  const v = sheet_(VISITS, VISIT_HEAD).getDataRange().getValues().slice(1);
  const now = Date.now(), n = new Date(now + NPT);
  const todayStart = Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), n.getUTCDate()) - NPT;
  let today = 0, week = 0, opens = 0;
  v.forEach(x => { const last = +x[2]; if (last >= todayStart) today++; if (last >= now - 7 * 864e5) week++; opens += +x[3] || 0; });
  return { devices: v.length, today: today, week: week, opens: opens };
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clean_(s, max) { // plain text, no spreadsheet formulas
  return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').replace(/^[=+\-@]+/, '').trim().slice(0, max);
}
function num_(v, lo, hi) { v = Number(v); return isFinite(v) ? Math.min(hi, Math.max(lo, v)) : lo; }
function findRow_(sh, test) {
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (test(vals[i])) return i + 1;
  return 0;
}

function doGet() {
  const b = sheet_(BOARD, BOARD_HEAD).getDataRange().getValues().slice(1);
  const r = sheet_(REPORTS, REP_HEAD).getDataRange().getValues().slice(1);
  return out_({
    ok: true,
    board: b.map(x => ({ id: x[0], name: x[1], sets: x[2], pts: x[3], best: x[4], avg: x[5], at: x[6] })),
    reports: r.map(x => ({ id: x[0], k: x[1], r: x[2], n: x[3], at: x[4] })),
    visits: visitStats_()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || '');
    if (!/^[a-z0-9]{8,40}$/i.test(id)) return out_({ ok: false });

    if (d.action === 'visit') {
      const sh = sheet_(VISITS, VISIT_HEAD);
      const at = findRow_(sh, v => String(v[0]) === id);
      if (at) { const c = sh.getRange(at, 3, 1, 2).getValues()[0]; sh.getRange(at, 3, 1, 2).setValues([[Date.now(), (+c[1] || 0) + 1]]); }
      else sh.appendRow(["'" + id, Date.now(), Date.now(), 1]);
      return out_({ ok: true });
    }
    if (d.action === 'board') {
      const sh = sheet_(BOARD, BOARD_HEAD);
      const row = ["'" + id, "'" + clean_(d.name, 40), num_(d.sets, 0, 40), num_(d.pts, 0, 4000),
                   num_(d.best, 0, 100), num_(d.avg, 0, 100), Date.now()];
      if (row[1] === "'") return out_({ ok: false });
      const at = findRow_(sh, v => String(v[0]) === id);
      if (at) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
      return out_({ ok: true });
    }
    if (d.action === 'unboard') {
      const sh = sheet_(BOARD, BOARD_HEAD);
      const at = findRow_(sh, v => String(v[0]) === id);
      if (at) sh.deleteRow(at);
      return out_({ ok: true });
    }
    if (d.action === 'report') {
      const k = String(d.k || '');
      const m = k.match(/^(\d{1,2})-(\d{1,3})$/);
      if (!m || +m[1] < 1 || +m[1] > 40 || +m[2] < 1 || +m[2] > 100) return out_({ ok: false });
      const reason = REASONS.indexOf(d.r) >= 0 ? d.r : 'other';
      const sh = sheet_(REPORTS, REP_HEAD);
      const row = ["'" + id, "'" + k, reason, "'" + clean_(d.n, 500), Date.now()];  // ' keeps values as text
      const at = findRow_(sh, v => String(v[0]) === id && String(v[1]) === k);
      if (at) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
      return out_({ ok: true });
    }
    return out_({ ok: false });
  } catch (err) {
    return out_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}
