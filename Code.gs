/**
 * NEC Industrial Engineering Practice - shared backend
 * (leaderboard, reports, notes, visits, question discussion and
 * friend challenges) for the GitHub Pages copy.
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
const COMMENTS = 'comments', COMMENT_HEAD = ['cid', 'id', 'name', 'k', 't', 'at', 'hidden'];
const CHALS = 'challenges', CHAL_HEAD = ['code', 'id', 'name', 'title', 'keys', 'min', 'at'];
const CHRES = 'chres', CHRES_HEAD = ['code', 'id', 'name', 'c', 't', 's', 'at'];
const KEY_RE = /^(\d{1,2})-(\d{1,3})$/;
function keyOk_(k) { const m = String(k).match(KEY_RE); return !!m && +m[1] >= 1 && +m[1] <= 40 && +m[2] >= 1 && +m[2] <= 100; }
function socialData_() {
  const c = sheet_(COMMENTS, COMMENT_HEAD).getDataRange().getValues().slice(1).filter(x => !x[6]).slice(-1500)
    .map(x => ({ cid: String(x[0]), id: String(x[1]), name: x[2], k: String(x[3]), t: x[4], at: +x[5] }));
  const ch = sheet_(CHALS, CHAL_HEAD).getDataRange().getValues().slice(1).slice(-300)
    .map(x => ({ code: String(x[0]), id: String(x[1]), name: x[2], title: x[3], keys: String(x[4]), min: +x[5], at: +x[6] }));
  const r = sheet_(CHRES, CHRES_HEAD).getDataRange().getValues().slice(1).slice(-3000)
    .map(x => ({ code: String(x[0]), id: String(x[1]), name: x[2], c: +x[3], t: +x[4], s: +x[5], at: +x[6] }));
  return { comments: c, challenges: ch, chres: r };
}
function addComment_(d, id) {
  const k = String(d.k || ''), t = clean_(d.t, 500), name = clean_(d.name, 40);
  if (!keyOk_(k) || !t || !name) return { ok: false, error: 'Invalid comment' };
  const sh = sheet_(COMMENTS, COMMENT_HEAD), now = Date.now();
  const recent = sh.getDataRange().getValues().slice(1).filter(x => String(x[1]) === id && +x[5] > now - 864e5).length;
  if (recent >= 60) return { ok: false, error: 'Daily comment limit reached' };
  sh.appendRow(["'" + Utilities.getUuid(), "'" + id, "'" + name, "'" + k, "'" + t, now, '']);
  return { ok: true };
}
function delComment_(d, id) {
  const sh = sheet_(COMMENTS, COMMENT_HEAD);
  const at = findRow_(sh, v => String(v[0]) === String(d.cid) && String(v[1]) === id);
  if (!at) return { ok: false, error: 'Not found' };
  sh.getRange(at, 7).setValue('deleted'); return { ok: true };
}
function addChallenge_(d, id) {
  const code = String(d.code || '').toUpperCase(), name = clean_(d.name, 40), title = clean_(d.title, 80);
  const keys = Array.isArray(d.keys) ? d.keys.map(String) : [];
  if (!/^[A-Z2-9]{6}$/.test(code) || !name || !keys.length || keys.length > 100 || !keys.every(keyOk_)) return { ok: false, error: 'Invalid challenge' };
  const sh = sheet_(CHALS, CHAL_HEAD);
  if (findRow_(sh, v => String(v[0]) === code)) return { ok: false, error: 'Code already used, try again' };
  sh.appendRow(["'" + code, "'" + id, "'" + name, "'" + title, "'" + JSON.stringify(keys), num_(d.min, 0, 150), Date.now()]);
  return { ok: true };
}
function addChres_(d, id) {
  const code = String(d.code || '').toUpperCase(), name = clean_(d.name, 40);
  if (!findRow_(sheet_(CHALS, CHAL_HEAD), v => String(v[0]) === code) || !name) return { ok: false, error: 'Unknown challenge' };
  sheet_(CHRES, CHRES_HEAD).appendRow(["'" + code, "'" + id, "'" + name, num_(d.c, 0, 100), num_(d.t, 1, 100), num_(d.s, 0, 36000), Date.now()]);
  return { ok: true };
}

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

/* ---- live quiz battle (kept in the script cache, 6 h) ---- */
function btGet_(code) {
  const c = CacheService.getScriptCache(), b = c.get('bt_' + code), p = c.get('btp_' + code);
  return { ok: true, battle: b ? JSON.parse(b) : null, players: p ? JSON.parse(p) : {} };
}
function btPost_(d, id) {
  const code = String(d.code || '').toUpperCase();
  if (!/^[A-Z2-9]{6}$/.test(code)) return { ok: false, error: 'Bad code' };
  const c = CacheService.getScriptCache(), bk = 'bt_' + code, pk = 'btp_' + code;
  const cur = c.get(bk) ? JSON.parse(c.get(bk)) : null;
  if (d.op === 'host') {
    if (cur && cur.host !== id) return { ok: false, error: 'Code in use' };
    const b = d.data || {}; if (!Array.isArray(b.keys) || !b.keys.length || b.keys.length > 30 || !b.keys.every(keyOk_)) return { ok: false, error: 'Bad battle' };
    b.code = code; b.host = id; b.hostName = clean_(d.name, 40); b.at = Date.now();
    c.put(bk, JSON.stringify(b), 21600); if (!cur) c.put(pk, '{}', 21600); return { ok: true };
  }
  if (d.op === 'upd') {
    if (!cur || cur.host !== id) return { ok: false, error: 'Not host' };
    const patch = d.data || {}; ['st', 'qi', 'round', 'keys', 'at', 'v'].forEach(f => { if (f in patch) cur[f] = patch[f]; });
    cur.at = Date.now(); c.put(bk, JSON.stringify(cur), 21600); return { ok: true };
  }
  if (d.op === 'player') {
    if (!cur) return { ok: false, error: 'No battle' };
    const ps = c.get(pk) ? JSON.parse(c.get(pk)) : {}, p = d.data || {};
    if (!p.code) delete ps[id]; else ps[id] = { code: code, r: +p.r || 1, a: p.a || {}, name: clean_(d.name, 40), at: Date.now() };
    c.put(pk, JSON.stringify(ps), 21600); return { ok: true };
  }
  return { ok: false };
}
/* ---- notice board edited from the site (admin key in Script Properties: ADMIN_KEY) ---- */
function noticesGet_() { const v = PropertiesService.getScriptProperties().getProperty('NOTICES'); try { return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function noticesSave_(d) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key || String(d.key || '') !== key) return { ok: false, error: 'Wrong admin key' };
  const j = JSON.stringify(d.data || {}); if (j.length > 8500) return { ok: false, error: 'Too long' };
  PropertiesService.getScriptProperties().setProperty('NOTICES', j); return { ok: true };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.bt) return out_(btGet_(String(e.parameter.bt).toUpperCase()));
  if (e && e.parameter && e.parameter.adminCheck !== undefined) { const k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'); return out_({ ok: !!k && e.parameter.adminCheck === k }); }
  const b = sheet_(BOARD, BOARD_HEAD).getDataRange().getValues().slice(1);
  const r = sheet_(REPORTS, REP_HEAD).getDataRange().getValues().slice(1);
  return out_(Object.assign({
    ok: true,
    notices: noticesGet_(),
    board: b.map(x => ({ id: x[0], name: x[1], sets: x[2], pts: x[3], best: x[4], avg: x[5], at: x[6] })),
    reports: r.map(x => ({ id: x[0], k: x[1], r: x[2], n: x[3], at: x[4] })),
    visits: visitStats_(),
    notes: notesList_()
  }, socialData_()));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || '');
    if (!/^[a-z0-9]{8,40}$/i.test(id)) return out_({ ok: false });

    if (d.action === 'bt') return out_(btPost_(d, id));
    if (d.action === 'saveNotices') return out_(noticesSave_(d));
    if (d.action === 'note') return out_(addNote_(d, id));
    if (d.action === 'deleteNote') return out_(deleteNote_(d, id));
    if (d.action === 'comment') return out_(addComment_(d, id));
    if (d.action === 'delComment') return out_(delComment_(d, id));
    if (d.action === 'challenge') return out_(addChallenge_(d, id));
    if (d.action === 'chres') return out_(addChres_(d, id));
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
