/**
 * 루브르 신진 화가展 · 학급 갤러리 서버 (Google Apps Script)
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여 넣으세요.
 * 자세한 순서는 '서버_설치_안내.txt'를 보세요.
 */

// ▼▼▼ 교사 비밀번호: 꼭 바꾸세요 (숫자 4~8자리 권장) ▼▼▼
const TEACHER_PIN = '1234';
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

const SHEET_NAME = 'artworks';
const HEAD = ['id','createdAt','updatedAt','cls','nickname','title','note','pin','shapes','tags','likes','hidden','pinned','comments'];

function doGet() {
  return json_({ ok: true, msg: '루브르 갤러리 서버가 작동 중입니다.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'invalid' }); }
  try { return json_(handle_(req)); }
  catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

function handle_(r) {
  const a = r.action;

  // ---- 읽기 ----
  if (a === 'list') {
    return { ok: true, items: readAll_().filter(x => !x.hidden && (!r.cls || x.cls === r.cls)).map(pub_) };
  }
  if (a === 'adminLogin') { checkTeacher_(r); return { ok: true }; }
  if (a === 'adminList')  { checkTeacher_(r); return { ok: true, items: readAll_().map(pub_) }; }

  // ---- 쓰기 (동시에 여러 학생이 보내도 안전하게 잠금) ----
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('busy');
  try {
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    const idx = id => { for (let i = 1; i < rows.length; i++) if (String(rows[i][0]) === String(id)) return i; return -1; };
    const get = i => toObj_(rows[i]);
    const put = (i, o) => sh.getRange(i + 1, 1, 1, HEAD.length).setValues([toRow_(o)]);

    if (a === 'submit') {
      const shapes = checkShapes_(r.shapes);
      const o = {
        id: 'w' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        createdAt: Date.now(), updatedAt: Date.now(),
        cls: txt_(r.cls, 10, true), nickname: txt_(r.nickname, 12, true),
        title: txt_(r.title, 24, true), note: txt_(r.note, 80),
        pin: /^\d{4}$/.test(String(r.pin || '')) ? String(r.pin) : '',
        shapes: shapes, tags: checkTags_(r.tags), likes: 0, hidden: false, pinned: false, comments: []
      };
      sh.appendRow(toRow_(o));
      return { ok: true, item: pub_(o) };
    }

    const i = idx(r.id);
    if (a === 'adminDelete') { checkTeacher_(r); if (i > 0) sh.deleteRow(i + 1); return { ok: true }; }
    if (i < 0) throw new Error('not_found');
    const o = get(i);

    switch (a) {
      case 'verify':
        if (!o.pin) throw new Error('no_pin');
        if (String(r.pin) !== o.pin) throw new Error('wrong_pin');
        return { ok: true };
      case 'update':
        if (!o.pin || String(r.pin) !== o.pin) throw new Error('wrong_pin');
        o.title = txt_(r.title, 24, true); o.note = txt_(r.note, 80);
        o.shapes = checkShapes_(r.shapes); o.tags = checkTags_(r.tags); o.updatedAt = Date.now();
        put(i, o); return { ok: true, item: pub_(o) };
      case 'deleteOwn':
        if (!o.pin || String(r.pin) !== o.pin) throw new Error('wrong_pin');
        sh.deleteRow(i + 1); return { ok: true };
      case 'like':
        o.likes = Math.max(0, (Number(o.likes) || 0) + (Number(r.delta) < 0 ? -1 : 1));
        put(i, o); return { ok: true, likes: o.likes };
      case 'comment':
        o.comments.push({ id: Math.random().toString(36).slice(2, 9), nickname: txt_(r.nickname, 12, true), text: txt_(r.text, 40, true), at: Date.now() });
        o.comments = o.comments.slice(-50);
        put(i, o); return { ok: true, comments: o.comments };
      case 'adminSet':
        checkTeacher_(r);
        if ('hidden' in r) o.hidden = !!r.hidden;
        if ('pinned' in r) o.pinned = !!r.pinned;
        put(i, o); return { ok: true, item: pub_(o) };
      case 'adminDeleteComment':
        checkTeacher_(r);
        o.comments = o.comments.filter(c => c.id !== r.cid);
        put(i, o); return { ok: true, comments: o.comments };
    }
    throw new Error('invalid');
  } finally {
    lock.releaseLock();
  }
}

// ---------- 도우미 ----------
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function readAll_() {
  const v = sheet_().getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(toObj_);
}
function toObj_(r) {
  const o = {}; HEAD.forEach((k, i) => o[k] = r[i]);
  o.id = String(o.id); o.pin = String(o.pin || '').replace(/^p/, '');
  o.shapes = parse_(o.shapes, []); o.tags = parse_(o.tags, []); o.comments = parse_(o.comments, []);
  o.hidden = o.hidden === true || o.hidden === 'TRUE'; o.pinned = o.pinned === true || o.pinned === 'TRUE';
  o.likes = Number(o.likes) || 0; o.createdAt = Number(o.createdAt) || 0; o.updatedAt = Number(o.updatedAt) || 0;
  return o;
}
function toRow_(o) {
  return HEAD.map(k => {
    if (k === 'pin') return o.pin ? 'p' + o.pin : '';      // 앞자리 0이 사라지지 않게 글자로 저장
    if (k === 'shapes' || k === 'tags' || k === 'comments') return JSON.stringify(o[k] || []);
    return o[k];
  });
}
function pub_(o) {
  return { id: o.id, createdAt: o.createdAt, updatedAt: o.updatedAt, cls: o.cls, nickname: o.nickname, title: o.title, note: o.note,
    shapes: o.shapes, tags: o.tags, likes: o.likes, hidden: o.hidden, pinned: o.pinned, comments: o.comments, hasPin: !!o.pin };
}
function parse_(s, d) { try { return typeof s === 'string' && s ? JSON.parse(s) : (Array.isArray(s) ? s : d); } catch (e) { return d; } }
function txt_(s, max, required) {
  let t = String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().replace(/^[=+\-@]+/, '').slice(0, max);
  if (required && !t) throw new Error('invalid');
  return t;
}
function checkShapes_(a) {
  if (!Array.isArray(a) || a.length < 1 || a.length > 150) throw new Error('invalid');
  const s = JSON.stringify(a);
  if (s.length > 48000) throw new Error('invalid');
  return a;
}
function checkTags_(a) { return (Array.isArray(a) ? a : []).map(String).slice(0, 6); }
function checkTeacher_(r) { if (String(r.tpin || '') !== TEACHER_PIN) throw new Error('wrong_pin'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
