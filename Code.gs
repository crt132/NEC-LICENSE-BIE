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
const NOTES = 'notes';
const NOTE_HEAD = ['noteId', 'id', 'name', 'title', 'chapter', 'text', 'fileName', 'fileType', 'fileSize', 'fileId', 'url', 'at', 'hidden'];
const NOTES_FOLDER = 'NEC Batch 2078 - Shared Notes';
const MAX_FILE = 10 * 1024 * 1024;          // 10 MB per file
const MAX_UPLOADS_PER_DAY = 20;             // per device
const ALLOWED = ['pdf','png','jpg','jpeg','webp','gif','doc','docx','ppt','pptx','xls','xlsx','txt','csv'];
const VISITS = 'visits';
const VISIT_HEAD = ['id', 'first', 'last', 'count'];
const NPT = 345 * 60000; // Nepal time offset

function notesFolder_() {
  const it = DriveApp.getFoldersByName(NOTES_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NOTES_FOLDER);
}
function notesList_() {
  const v = sheet_(NOTES, NOTE_HEAD).getDataRange().getValues().slice(1);
  return v.filter(x => !x[12]).map(x => ({ noteId: String(x[0]), by: String(x[1]), name: x[2], title: x[3], ch: +x[4] || 0, text: x[5],
    fileName: x[6], fileType: x[7], fileSize: +x[8] || 0, url: x[10], at: +x[11] })).sort((a, b) => b.at - a.at).slice(0, 500);
}
function addNote_(d, id) {
  const title = clean_(d.title, 120), name = clean_(d.name, 40), text = clean_(d.text, 3000), ch = num_(d.ch, 0, 10);
  if (!title || !name) return { ok: false, error: 'Title and name are required' };
  const sh = sheet_(NOTES, NOTE_HEAD), now = Date.now();
  const recent = sh.getDataRange().getValues().slice(1).filter(x => String(x[1]) === id && +x[11] > now - 864e5).length;
  if (recent >= MAX_UPLOADS_PER_DAY) return { ok: false, error: 'Daily upload limit reached' };
  let fileName = '', fileType = '', fileSize = 0, fileId = '', url = '';
  if (d.file && d.file.data) {
    fileName = clean_(d.file.name, 120).replace(/[\\/:*?"<>|]/g, '_');
    const ext = (fileName.split('.').pop() || '').toLowerCase();
    if (ALLOWED.indexOf(ext) < 0) return { ok: false, error: 'File type not allowed' };
    const bytes = Utilities.base64Decode(String(d.file.data));
    if (bytes.length > MAX_FILE) return { ok: false, error: 'File is larger than 10 MB' };
    fileType = clean_(d.file.type, 100) || 'application/octet-stream';
    const blob = Utilities.newBlob(bytes, fileType, fileName);
    const f = notesFolder_().createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileSize = bytes.length; fileId = f.getId(); url = 'https://drive.google.com/file/d/' + fileId + '/view';
  } else if (!text) return { ok: false, error: 'Add a file or some text' };
  const noteId = Utilities.getUuid();
  sh.appendRow(["'" + noteId, "'" + id, "'" + name, "'" + title, ch, "'" + text, "'" + fileName, fileType, fileSize, fileId, url, now, '']);
  return { ok: true, noteId: noteId };
}
function deleteNote_(d, id) { // uploader can remove their own note (hidden, file trashed)
  const sh = sheet_(NOTES, NOTE_HEAD);
  const at = findRow_(sh, v => String(v[0]) === String(d.noteId) && String(v[1]) === id);
  if (!at) return { ok: false, error: 'Not found' };
  const fid = sh.getRange(at, 10).getValue();
  if (fid) try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {}
  sh.getRange(at, 13).setValue('deleted');
  return { ok: true };
}

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
    visits: visitStats_(),
    notes: notesList_()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || '');
    if (!/^[a-z0-9]{8,40}$/i.test(id)) return out_({ ok: false });

    if (d.action === 'note') return out_(addNote_(d, id));
    if (d.action === 'deleteNote') return out_(deleteNote_(d, id));
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
