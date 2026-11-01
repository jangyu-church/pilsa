/**
 * 쓰다보면 — 성경필사 모임 서버 (구글 시트 + 앱스 스크립트)
 *
 * 이 코드는 구글 시트 "쓰다보면"에 연결된 앱스 스크립트(확장 프로그램 > Apps Script)에 붙여 넣습니다.
 * 처음 한 번 setup() 을 실행하고, 웹 앱으로 배포하면 나오는 주소(…/exec)를 config.js 에 넣습니다.
 * 자세한 순서는 저장소의 README.md 를 보세요.
 */

var TZ = 'Asia/Seoul';
var PHOTO_KEEP_DAYS = 14;      // 인증 사진은 최근 2주만 보관
var MAX_MEMBERS = 200;
var MAX_PLAN_DAYS = 800;
var MAX_PHOTO_CHARS = 450000;  // 압축된 사진(base64) 최대 길이

/* ───────── 처음 한 번 실행하는 준비 함수 ───────── */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var layout = {
    members: ['id', 'name', 'pinHash', 'joined'],
    admins: ['name'],
    config: ['key', 'value'],
    plan: ['date', 'label', 'units'],
    checks: ['memberId', 'date', 'at', 'photoId'],
    notice: ['text', 'at']
  };
  Object.keys(layout).forEach(function (name) {
    var s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getLastRow() === 0) s.getRange(1, 1, 1, layout[name].length).setValues([layout[name]]);
    s.setFrozenRows(1);
  });
  // 날짜 칸이 자동으로 날짜 형식으로 바뀌지 않도록 글자 형식으로 고정
  ss.getSheetByName('plan').getRange('A:B').setNumberFormat('@');
  ss.getSheetByName('checks').getRange('B:B').setNumberFormat('@');
  var cfg = ss.getSheetByName('config');
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, 1, 2).setValues([['title', '']]);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('PHOTO_FOLDER')) {
    var folder = DriveApp.createFolder('쓰다보면 인증사진');
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty('PHOTO_FOLDER', folder.getId());
  }
  Logger.log('준비 완료. 이제 [배포 > 새 배포 > 웹 앱]으로 배포하세요.');
}

/* ───────── 웹 앱 진입점 ───────── */
function doGet() { return out_({ ok: true, app: '쓰다보면' }); }

function doPost(e) {
  var res;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    res = route_(body);
  } catch (err) {
    res = (err && err.error) ? { ok: false, error: err.error, message: err.message || '' }
                             : { ok: false, error: 'server', message: String(err && err.message || err) };
  }
  return out_(res);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

var READ_ONLY = { state: 1, login: 1 };
var HANDLERS = {
  register: register_, login: login_, state: state_, check: check_, photo: photo_, rename: rename_,
  setPlan: setPlan_, setNotice: setNotice_, removeMember: removeMember_, resetPin: resetPin_
};

function route_(b) {
  var h = HANDLERS[b.action];
  if (!h) throw { error: 'bad_action' };
  if (READ_ONLY[b.action]) return h(b);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return h(b); } finally { lock.releaseLock(); }
}

/* ───────── 시트 읽기 도구 ───────── */
var WIDTH = { members: 4, admins: 1, config: 2, plan: 3, checks: 4, notice: 2 };
function sheet_(name) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw { error: 'server', message: '시트 "' + name + '"이(가) 없어요. setup() 을 실행했는지 확인해 주세요.' };
  return s;
}
function rows_(name) {
  var s = sheet_(name), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, WIDTH[name]).getValues();
}
function ds_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}
function todayKst_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/* ───────── 입력 검사 ───────── */
function cleanName_(x) {
  var s = String(x == null ? '' : x).replace(/[\u0000-\u001f]/g, '').replace(/\s+/g, ' ').trim();
  if (s.length < 1 || s.length > 12) return '';
  if (/^[=+\-@]/.test(s)) return '';     // 시트 수식으로 읽히는 이름 방지
  return s;
}
function cleanText_(x, max) {
  return String(x == null ? '' : x).replace(/[\u0000-\u001f]/g, ' ').replace(/^[=+\-@\s]+/, '').trim().slice(0, max);
}
function cleanHash_(x) { var s = String(x || '').toLowerCase(); return /^[0-9a-f]{64}$/.test(s) ? s : ''; }
function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var p = s.split('-'), d = new Date(+p[0], +p[1] - 1, +p[2]);
  return d.getFullYear() === +p[0] && d.getMonth() === +p[1] - 1 && d.getDate() === +p[2];
}

/* ───────── 인증 ───────── */
function secret_() {
  var s = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!s) throw { error: 'server', message: '준비가 끝나지 않았어요. setup() 을 먼저 실행해 주세요.' };
  return s;
}
function tokenFor_(id) {
  var sig = Utilities.computeHmacSha256Signature(id, secret_());
  return sig.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function isAdmin_(name) {
  return rows_('admins').some(function (r) { return String(r[0]).trim() === name; });
}
function need_(b) {
  var id = String(b.id || '');
  if (!id || !b.token || String(b.token) !== tokenFor_(id)) throw { error: 'auth' };
  var ms = rows_('members');
  for (var i = 0; i < ms.length; i++) {
    if (String(ms[i][0]) === id) return { id: id, name: String(ms[i][1]), row: i + 2 };
  }
  throw { error: 'auth' };
}
function needAdmin_(b) {
  var me = need_(b);
  if (!isAdmin_(me.name)) throw { error: 'forbidden', message: '관리자만 할 수 있어요.' };
  return me;
}

/* ───────── 회원 ───────── */
function register_(b) {
  var name = cleanName_(b.name), pin = cleanHash_(b.pin);
  if (!name) throw { error: 'bad_name' };
  if (!pin) throw { error: 'bad_pin' };
  var ms = rows_('members');
  if (ms.length >= MAX_MEMBERS) throw { error: 'full' };
  for (var i = 0; i < ms.length; i++) if (String(ms[i][1]) === name) throw { error: 'name_taken' };
  var id = 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  sheet_('members').appendRow([id, name, pin, Date.now()]);
  return { ok: true, id: id, token: tokenFor_(id), name: name, admin: isAdmin_(name) };
}

function login_(b) {
  var name = cleanName_(b.name), pin = cleanHash_(b.pin);
  if (!name || !pin) throw { error: 'bad_login' };
  var cache = CacheService.getScriptCache(), key = 'fail:' + name;
  if (Number(cache.get(key) || 0) >= 5) throw { error: 'locked' };
  var ms = rows_('members');
  for (var i = 0; i < ms.length; i++) {
    if (String(ms[i][1]) === name) {
      if (String(ms[i][2]) === pin) {
        cache.remove(key);
        return { ok: true, id: String(ms[i][0]), token: tokenFor_(String(ms[i][0])), name: name, admin: isAdmin_(name) };
      }
      break;
    }
  }
  cache.put(key, String(Number(cache.get(key) || 0) + 1), 600);
  throw { error: 'bad_login' };
}

function rename_(b) {
  var me = need_(b), name = cleanName_(b.name);
  if (!name) throw { error: 'bad_name' };
  var ms = rows_('members');
  for (var i = 0; i < ms.length; i++) if (String(ms[i][1]) === name && String(ms[i][0]) !== me.id) throw { error: 'name_taken' };
  sheet_('members').getRange(me.row, 2).setValue(name);
  return { ok: true, name: name, admin: isAdmin_(name) };
}

/* ───────── 화면에 필요한 모든 데이터 ───────── */
function state_(b) {
  var me = need_(b);
  var plan = rows_('plan').map(function (r) { return { d: ds_(r[0]), t: String(r[1]), n: Number(r[2]) || 1 }; })
    .filter(function (x) { return validDate_(x.d); })
    .sort(function (a, c) { return a.d < c.d ? -1 : 1; });
  var cfg = rows_('config'), title = '';
  cfg.forEach(function (r) { if (String(r[0]) === 'title') title = String(r[1]); });
  var members = rows_('members').map(function (r) { return { id: String(r[0]), name: String(r[1]) }; });
  var checks = {};
  rows_('checks').forEach(function (r) {
    var id = String(r[0]); (checks[id] = checks[id] || []).push([ds_(r[1]), Number(r[2]) || 0, String(r[3] || '')]);
  });
  var nt = rows_('notice'), notice = nt.length ? String(nt[nt.length - 1][0]) : '';
  return { ok: true, title: title, plan: plan, members: members, checks: checks, notice: notice,
           admin: isAdmin_(me.name), today: todayKst_(), now: Date.now() };
}

/* ───────── 체크 / 사진 ───────── */
function planHas_(date) { return rows_('plan').some(function (r) { return ds_(r[0]) === date; }); }
function checkDate_(b) {
  var d = String(b.date || '');
  if (!validDate_(d) || d > todayKst_()) throw { error: 'bad_run', message: '아직 체크할 수 없는 날짜예요.' };
  if (!planHas_(d)) throw { error: 'bad_run', message: '일정에 없는 날짜예요.' };
  return d;
}
function findCheck_(rows, id, date) {
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === id && ds_(rows[i][1]) === date) return i;
  return -1;
}
function trashPhoto_(pid) {
  if (!pid) return;
  try { DriveApp.getFileById(pid).setTrashed(true); } catch (e) { /* 이미 지워졌으면 넘어감 */ }
}

function check_(b) {
  var me = need_(b), d = checkDate_(b), s = sheet_('checks'), rs = rows_('checks'), i = findCheck_(rs, me.id, d);
  if (b.on === false) {
    if (i >= 0) { trashPhoto_(String(rs[i][3] || '')); s.deleteRow(i + 2); }
    return { ok: true };
  }
  var at = Date.now();
  if (i >= 0) return { ok: true, at: Number(rs[i][2]) || at };
  s.appendRow([me.id, d, at, '']);
  s.getRange(s.getLastRow(), 2).setNumberFormat('@').setValue(d);
  return { ok: true, at: at };
}

function photo_(b) {
  var me = need_(b), d = checkDate_(b);
  var m = /^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/.exec(String(b.img || ''));
  if (!m || m[1].length > MAX_PHOTO_CHARS) throw { error: 'bad_run', message: '사진 크기나 형식이 올바르지 않아요.' };
  var props = PropertiesService.getScriptProperties();
  var folder = DriveApp.getFolderById(props.getProperty('PHOTO_FOLDER'));
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', me.id + '_' + d + '.jpg'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var s = sheet_('checks'), rs = rows_('checks'), i = findCheck_(rs, me.id, d), at = Date.now();
  if (i >= 0) {
    trashPhoto_(String(rs[i][3] || ''));
    s.getRange(i + 2, 4).setValue(file.getId());
    at = Number(rs[i][2]) || at;
  } else {
    s.appendRow([me.id, d, at, file.getId()]);
    s.getRange(s.getLastRow(), 2).setNumberFormat('@').setValue(d);
  }
  // 오래된 내 사진 정리
  var cutoff = Utilities.formatDate(new Date(Date.now() - PHOTO_KEEP_DAYS * 864e5), TZ, 'yyyy-MM-dd');
  rs = rows_('checks');
  rs.forEach(function (r, k) {
    if (String(r[0]) === me.id && String(r[3]) && ds_(r[1]) < cutoff) {
      trashPhoto_(String(r[3])); s.getRange(k + 2, 4).setValue('');
    }
  });
  return { ok: true, photoId: file.getId(), at: at };
}

/* ───────── 관리자 ───────── */
function setPlan_(b) {
  needAdmin_(b);
  var items = Array.isArray(b.items) ? b.items : [];
  if (items.length > MAX_PLAN_DAYS) throw { error: 'bad_run', message: '일정이 너무 길어요.' };
  var seen = {}, clean = [];
  items.forEach(function (x) {
    var d = String(x.d || ''), t = cleanText_(x.t, 80), n = Number(x.n);
    if (!validDate_(d) || !t || !(n > 0 && n <= 200) || seen[d]) throw { error: 'bad_run', message: '일정 형식을 확인해 주세요.' };
    seen[d] = 1; clean.push([d, t, n]);
  });
  clean.sort(function (a, c) { return a[0] < c[0] ? -1 : 1; });
  var s = sheet_('plan');
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
  if (clean.length) {
    s.getRange(2, 1, clean.length, 2).setNumberFormat('@');
    s.getRange(2, 1, clean.length, 3).setValues(clean);
  }
  var cfg = sheet_('config'), title = cleanText_(b.title, 40), rs = rows_('config'), done = false;
  rs.forEach(function (r, i) { if (String(r[0]) === 'title') { cfg.getRange(i + 2, 2).setValue(title); done = true; } });
  if (!done) cfg.appendRow(['title', title]);
  return { ok: true, count: clean.length };
}

function setNotice_(b) {
  needAdmin_(b);
  sheet_('notice').appendRow([cleanText_(b.text, 200), Date.now()]);
  return { ok: true };
}

function removeMember_(b) {
  needAdmin_(b);
  var target = String(b.target || ''), ms = rows_('members');
  for (var i = 0; i < ms.length; i++) {
    if (String(ms[i][0]) === target) {
      sheet_('members').deleteRow(i + 2);
      var cs = sheet_('checks'), rs = rows_('checks');
      for (var k = rs.length - 1; k >= 0; k--) {
        if (String(rs[k][0]) === target) { trashPhoto_(String(rs[k][3] || '')); cs.deleteRow(k + 2); }
      }
      return { ok: true };
    }
  }
  throw { error: 'bad_run', message: '없는 회원이에요.' };
}

function resetPin_(b) {
  needAdmin_(b);
  var target = String(b.target || ''), pin = cleanHash_(b.pin), ms = rows_('members');
  if (!pin) throw { error: 'bad_pin' };
  for (var i = 0; i < ms.length; i++) {
    if (String(ms[i][0]) === target) {
      sheet_('members').getRange(i + 2, 3).setValue(pin);
      CacheService.getScriptCache().remove('fail:' + String(ms[i][1]));
      return { ok: true };
    }
  }
  throw { error: 'bad_run', message: '없는 회원이에요.' };
}
