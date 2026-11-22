/** @OnlyCurrentDoc */
// ===== الإعدادات =====
// مهم: عند أي تحديث لا تنشئي «عملية نشر جديدة». استخدمي: نشر ← إدارة عمليات النشر ← ✏️ تعديل ← الإصدار: «إصدار جديد» ← نشر.
// بهذه الطريقة يبقى رابط /exec نفسه، فلا يحتاج أي مستخدم لتغيير رابط Meta أو أي ربط آخر.
const WA_VERIFY_TOKEN = 'furqan-wa-2026';   // أي كلمة تختارينها، وتكتبينها نفسها في Meta
const VERSION = '2026-11-22-c';               // رقم الإصدار: يظهر عند فتح رابط النشر في المتصفح، ليُعرف هل النشر يعمل بآخر كود
const MSG = 'الرسائل', USR = 'المستخدمون';
const COLS = 11;                              // عدد أعمدة جدول الرسائل
const DEL = 'محذوفة';                         // قديم: صفوف قديمة تُحذف تلقائيًا من الشيت ولا يُكتب هذا النص بعد الآن
const STATUSES_ = ['جديدة', 'تم الرد', 'لم يُرد عليها'];
const AUTO_REPLIED = false;                   // false = «تم الرد» تُضبط يدويًا فقط من الموقع (لا تتغير تلقائيًا عند الرد من الجوال أو البريد)
// إشعارات البريد: تُرسل لرسائل «النموذج» فقط. رسائل الواتساب والبريد لا يُرسل عنها أي إشعار إطلاقًا.

// ===== التشفير =====
// كل بيانات الرسالة (الاسم، الجوال، البريد، النص، الملاحظة) تُحفظ في الشيت مشفّرة،
// فلو فتحتِ الشيت ما ترين إلا رموزًا غير مقروءة. تظهر الرسائل فقط داخل الموقع بعد تسجيل الدخول.
// المفتاح يُنشأ تلقائيًا ويُحفظ في خصائص السكربت (Script Properties) وليس في الشيت.
let KEY_ = null;
function key_() {
  if (KEY_) return KEY_;
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('MSG_KEY');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid(); p.setProperty('MSG_KEY', k); }
  return (KEY_ = k);
}
// تيار مفاتيح من HMAC-SHA256 (نمط العدّاد) ثم XOR مع بايتات النص
function xor_(bytes, nonce) {
  const key = key_(), out = new Array(bytes.length);
  for (let i = 0; i < bytes.length; i += 32) {
    const ks = Utilities.computeHmacSha256Signature(nonce + ':' + (i / 32), key);
    for (let j = 0; j < 32 && i + j < bytes.length; j++) out[i + j] = bytes[i + j] ^ ks[j];
  }
  return out;
}
function enc_(s) {
  s = String(s == null ? '' : s);
  if (!s) return '';
  const nonce = Utilities.getUuid();
  return 'enc1:' + nonce + ':' + Utilities.base64Encode(xor_(Utilities.newBlob(s).getBytes(), nonce));
}
function dec_(s) {
  s = String(s == null ? '' : s);
  if (s.indexOf('enc1:') !== 0) return s;               // نص قديم غير مشفّر (قبل التحديث)
  const p = s.split(':');
  try { return Utilities.newBlob(xor_(Utilities.base64Decode(p[2]), p[1])).getDataAsString('UTF-8'); }
  catch (err) { return ''; }
}

// ===== الجداول =====
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, header) {
  let sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function msgs_() {
  const sh = sheet_(MSG, ['المعرف', 'التاريخ', 'الاسم', 'الجوال', 'البريد', 'الرسالة', 'الحالة', 'ملاحظة', 'المصدر', 'المالكة', 'مؤرشفة']);
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('HDR')) {   // فحص العناوين مرة واحدة فقط
    const h = sh.getRange(1, 1, 1, COLS).getValues()[0];
    [[9, 'المصدر'], [10, 'المالكة'], [11, 'مؤرشفة']].forEach(x => { if (!h[x[0] - 1]) sh.getRange(1, x[0]).setValue(x[1]); });
    P.setProperty('HDR', '1');
  }
  return sh;
}
// ===== التذكيرات والمتابعات =====
// تُحفظ في ورقة مستقلة «التذكيرات»، والعنوان والملاحظة وبيانات الشخص مشفّرة بنفس طريقة الرسائل. الموعد (رقم) والحالة فقط ظاهران.
const REM = 'التذكيرات', RCOLS = 15;
function rems_() {
  const sh = sheet_(REM, ['المعرف', 'الإنشاء', 'المالكة', 'الموعد', 'الحالة', 'العنوان', 'ملاحظة', 'معرّف الرسالة', 'الاسم', 'الجوال', 'البريد', 'المصدر', 'تاريخ الإكمال', 'موعد آخر تنبيه جوال', 'وقت إرسال التنبيه']);
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('REMH2')) { sh.getRange(1, 14, 1, 2).setValues([['موعد آخر تنبيه جوال', 'وقت إرسال التنبيه']]); P.setProperty('REMH2', '1'); }
  return sh;
}
function remOut_(r) {
  return {
    id: r[0], created: new Date(r[1]).getTime() || 0, due: Number(r[3]) || 0, done: r[4] === 'مكتمل', doneAt: Number(r[12]) || 0,
    title: dec_(r[5]), note: dec_(r[6]), msgId: String(r[7] || ''), name: dec_(r[8]), phone: dec_(r[9]), email: dec_(r[10]), source: r[11] || ''
  };
}
function rtouch_(id) { const v = String(Date.now()); PropertiesService.getScriptProperties().setProperty('RC_' + id, v); return Number(v); }
function remChOf_(id) { return Number(PropertiesService.getScriptProperties().getProperty('RC_' + id) || 0); }
function findRem_(sh, id, me) {
  const n = sh.getLastRow() - 1; if (n < 1) return null;
  const vals = sh.getRange(2, 1, n, RCOLS).getValues(), i = vals.findIndex(r => r[0] === id && r[2] === me[0]);
  return i < 0 ? null : { row: i + 2, vals: vals[i] };
}
// ===== تنبيهات الجوال (Web Push) =====
// تُرسل من الخادم عند حلول موعد التذكير حتى والموقع مغلق. بلا مكتبات: توقيع VAPID (ES256) مكتوب هنا بـ BigInt.
// المشغّل الزمني pushTick يعمل كل دقيقة (يُنشأ مرة واحدة بتشغيل installPushTrigger من المحرر).
const B0 = BigInt(0), B1 = BigInt(1), B2 = BigInt(2), B3 = BigInt(3);   // بلا حروف BigInt الحرفية (0n) لأن محرر Apps Script لا يدعمها
const P256 = {
  p: BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'),
  n: BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
  g: [BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'), BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')]
};
function pmod_(a, m) { const r = a % m; return r >= B0 ? r : r + m; }
function pinv_(a, m) {
  let r0 = pmod_(a, m), r1 = m, s0 = B1, s1 = B0;
  while (r1) { const q = r0 / r1; [r0, r1] = [r1, r0 - q * r1]; [s0, s1] = [s1, s0 - q * s1]; }
  return pmod_(s0, m);
}
function padd_(A, B) {
  if (!A) return B; if (!B) return A;
  const p = P256.p; let l;
  if (A[0] === B[0]) { if (pmod_(A[1] + B[1], p) === B0) return null; l = pmod_((B3 * A[0] * A[0] - B3) * pinv_(B2 * A[1], p), p); }
  else l = pmod_((B[1] - A[1]) * pinv_(B[0] - A[0], p), p);
  const x = pmod_(l * l - A[0] - B[0], p);
  return [x, pmod_(l * (A[0] - x) - A[1], p)];
}
function pmul_(k, A) { let R = null; while (k > B0) { if (k & B1) R = padd_(R, A); A = padd_(A, A); k >>= B1; } return R; }
function big_(bytes) { let h = '0x'; bytes.forEach(b => { h += ('0' + (b & 255).toString(16)).slice(-2); }); return BigInt(h); }
function bytes_(x, len) { let h = x.toString(16); while (h.length < len * 2) h = '0' + h; return h.match(/../g).map(s => parseInt(s, 16)); }
function sgn_(a) { return a.map(v => v > 127 ? v - 256 : v); }
function b64u_(bytes) { return Utilities.base64EncodeWebSafe(sgn_(bytes)).replace(/=+$/, ''); }
function b64us_(str) { return b64u_(Utilities.newBlob(str).getBytes()); }
function vapid_() {   // المفتاح يُولَّد مرة واحدة داخل الخادم ويبقى في خصائص السكربت (لا يظهر في الموقع)
  const P = PropertiesService.getScriptProperties();
  let hex = P.getProperty('VAPID_D');
  if (!hex) {
    const seed = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now());
    hex = bytes_(pmod_(big_(seed), P256.n - B1) + B1, 32).map(b => ('0' + b.toString(16)).slice(-2)).join('');
    P.setProperty('VAPID_D', hex);
  }
  const d = BigInt('0x' + hex), Q = pmul_(d, P256.g);
  return { d: d, pub: b64u_([4].concat(bytes_(Q[0], 32), bytes_(Q[1], 32))) };
}
function es256_(d, msg) {
  const z = big_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, msg)), zb = bytes_(z, 32), db = sgn_(bytes_(d, 32));
  for (let c = 0; c < 20; c++) {   // k حتمي (RFC 6979 مبسّط)
    const k = pmod_(big_(Utilities.computeHmacSha256Signature(sgn_(zb.concat([c])), db)), P256.n);
    if (k === B0) continue;
    const r = pmod_(pmul_(k, P256.g)[0], P256.n); if (r === B0) continue;
    const s = pmod_(pinv_(k, P256.n) * (z + r * d), P256.n); if (s === B0) continue;
    return bytes_(r, 32).concat(bytes_(s, 32));
  }
  throw new Error('sign');
}
function vapidAuth_(endpoint, cache) {
  const aud = (/^(https:\/\/[^\/]+)/.exec(endpoint) || [])[1]; if (!aud) return null;
  const v = vapid_();
  if (!cache[aud]) {
    let sub = 'mailto:admin@example.com'; try { const m = Session.getEffectiveUser().getEmail(); if (m) sub = 'mailto:' + m; } catch (e) {}
    const input = b64us_(JSON.stringify({ typ: 'JWT', alg: 'ES256' })) + '.' + b64us_(JSON.stringify({ aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: sub }));
    cache[aud] = input + '.' + b64u_(es256_(v.d, input));
  }
  return 'vapid t=' + cache[aud] + ', k=' + v.pub;
}
function sendPush_(endpoint, cache) {   // إشعار بلا حمولة: الـ Service Worker يجلب التفاصيل من الخادم
  const auth = vapidAuth_(endpoint, cache); if (!auth) return 0;
  try { return UrlFetchApp.fetch(endpoint, { method: 'post', headers: { Authorization: auth, TTL: '86400', Urgency: 'high' }, muteHttpExceptions: true }).getResponseCode(); }
  catch (e) { return 0; }
}
function subs_(uid) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('PS_' + uid) || '[]'); } catch (e) { return []; } }
function setSubs_(uid, a) { const P = PropertiesService.getScriptProperties(); if (a.length) P.setProperty('PS_' + uid, JSON.stringify(a)); else P.deleteProperty('PS_' + uid); }
function pushAll_(uid, cache) {   // يرسل لكل أجهزتها، ويحذف الأجهزة المنتهية. يعيد عدد ما وصل
  const keep = []; let okN = 0;
  subs_(uid).forEach(s => {
    const c = sendPush_(s.e, cache);
    if (c === 404 || c === 410) return;
    keep.push(s); if (c >= 200 && c < 300) okN++;
  });
  setSubs_(uid, keep); return okN;
}
function pushTick() {
  const now = Date.now(), cand = {}, lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = rems_(), n = sh.getLastRow() - 1; if (n < 1) return;
    sh.getRange(2, 1, n, RCOLS).getValues().forEach(r => {
      const due = Number(r[3]) || 0;
      if (r[4] === 'مكتمل' || !due || due > now || now - due > 86400000 || Number(r[13]) === due) return;
      (cand[r[2]] = cand[r[2]] || []).push({ id: r[0], due: due });
    });
  } finally { lock.releaseLock(); }
  const cache = {}, done = [];
  Object.keys(cand).forEach(uid => { if (subs_(uid).length && pushAll_(uid, cache) > 0) cand[uid].forEach(c => done.push(c)); });
  if (!done.length) return;
  lock.waitLock(20000);
  try {
    const sh = rems_(), n = sh.getLastRow() - 1, ids = sh.getRange(2, 1, n, 1).getValues();
    done.forEach(c => { const i = ids.findIndex(x => x[0] === c.id); if (i >= 0) sh.getRange(i + 2, 14, 1, 2).setValues([[c.due, now]]); });
  } finally { lock.releaseLock(); }
}
function installPushTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'pushTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pushTick').timeBased().everyMinutes(1).create();
  vapid_();
  Logger.log('تم تفعيل مشغّل تنبيهات الجوال (كل دقيقة).');
}

function dueOk_(v) { v = Number(v); return isFinite(v) && v > 946684800000 && v < 4102444800000; }

function users_() {
  const sh = sheet_(USR, ['المعرف', 'الاسم', 'الجوال', 'البريد', 'التجزئة', 'الملح', 'الرمز', 'التاريخ']);
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('PH_FMT')) { sh.getRange('C:C').setNumberFormat('@'); P.setProperty('PH_FMT', '1'); }   // مرة واحدة فقط (كانت تُنفَّذ مع كل طلب فتبطّئ الموقع)
  return sh;
}

// علامة «تغيّر شيء» لكل مستخدمة: يفحصها الموقع كل بضع ثوانٍ بطلب خفيف، فتظهر الرسائل الجديدة فورًا دون تحديث الصفحة
function touch_(id) { PropertiesService.getScriptProperties().setProperty('CH_' + id, String(Date.now())); }
function chOf_(id) { return Number(PropertiesService.getScriptProperties().getProperty('CH_' + id) || 0); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function norm_(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (d.indexOf('00') === 0) d = d.slice(2);
  if (d.length === 10 && d[0] === '0') d = '966' + d.slice(1);
  else if (d.length === 9 && d[0] === '5') d = '966' + d;
  return d;
}
function hash_(pw, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw));
}
function pub_(r) { return { id: r[0], name: r[1], phone: r[2], email: r[3] }; }

let UC_ = null;   // المستخدمون: تُقرأ من ذاكرة مؤقتة (10 دقائق) بدل قراءة الشيت مع كل طلب
function allUsers_() {
  if (UC_) return UC_;
  const cache = CacheService.getScriptCache(), c = cache.get('USERS');
  if (c) { try { return (UC_ = JSON.parse(c)); } catch (e) {} }
  const sh = users_();
  UC_ = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  try { cache.put('USERS', JSON.stringify(UC_), 600); } catch (e) {}
  return UC_;
}
function dropUsers_() { UC_ = null; CacheService.getScriptCache().remove('USERS'); }
// ===== الجلسات: منفصلة تمامًا عن «مفتاح الربط» =====
// الجلسة (تسجيل الدخول من الموقع): تُحفظ بصمتها فقط، وتنتهي بعد 30 يومًا، وتُلغى عند تغيير/استعادة كلمة المرور.
// مفتاح الربط (العمود «الرمز» في المستخدمون): يُستخدم في روابط الواتساب وسكربت البريد، ويسمح بإضافة الرسائل فقط ولا يقرأ شيئًا.
function sh_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'ss:' + s)).slice(0, 16); }
function sessions_(id) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('S_' + id) || '[]'); } catch (e) { return []; } }
function newSession_(id) {
  const secret = Utilities.getUuid() + Utilities.getUuid(), now = Math.floor(Date.now() / 60000);
  const keep = sessions_(id).filter(x => parseInt(x.slice(16), 36) > now).slice(-4);
  keep.push(sh_(secret) + (now + 30 * 1440).toString(36));
  PropertiesService.getScriptProperties().setProperty('S_' + id, JSON.stringify(keep));
  return id + '.' + secret;
}
function endSessions_(id) { PropertiesService.getScriptProperties().deleteProperty('S_' + id); }
function dropSession_(t) {
  t = String(t || ''); const i = t.indexOf('.'); if (i < 1) return;
  const id = t.slice(0, i), h = sh_(t.slice(i + 1));
  PropertiesService.getScriptProperties().setProperty('S_' + id, JSON.stringify(sessions_(id).filter(x => x.slice(0, 16) !== h)));
}
function userBySession_(t) {
  t = String(t || ''); const i = t.indexOf('.'); if (i < 1) return null;
  const id = t.slice(0, i), h = sh_(t.slice(i + 1)), now = Math.floor(Date.now() / 60000);
  return sessions_(id).some(x => x.slice(0, 16) === h && parseInt(x.slice(16), 36) > now) ? (userById_(id) || null) : null;
}
function userByKey_(k) { return k ? allUsers_().find(r => r[6] === k) : null; }
function userById_(id) { return allUsers_().find(r => r[0] === id); }
function userByPhone_(p) { const n = norm_(p); return n ? allUsers_().find(r => norm_(r[2]) === n) : null; }

function rowOut_(r) {
  return {
    id: r[0], date: r[1], name: dec_(r[2]), phone: dec_(r[3]), email: dec_(r[4]), message: dec_(r[5]),
    status: r[6], note: dec_(r[7]), source: r[8] || 'النموذج', starred: r[10] === 'نعم'
  };
}
// يجد رسالة تخص هذه المستخدمة (الرسائل القديمة بلا مالكة تعود لأول حساب)
function findRow_(sh, id, me) {
  if (sh.getLastRow() < 2) return null;
  const first = allUsers_()[0][0];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues();
  const i = vals.findIndex(r => r[0] === id && (r[9] || first) === me[0] && r[6] !== DEL);
  return i < 0 ? null : { row: i + 2, vals: vals[i] };
}

function doGet(e) {
  const p = e.parameter || {};
  if (p['hub.mode'] === 'subscribe') {
    return ContentService.createTextOutput(p['hub.verify_token'] === WA_VERIFY_TOKEN ? p['hub.challenge'] : 'forbidden');
  }
  if (p.action === 'vapid') return out_({ ok: true, pub: vapid_().pub });
  if (p.action === 'profile') {
    const u = userById_(p.u);
    return u ? out_({ ok: true, user: { name: u[1], phone: u[2], email: u[3] } }) : out_({ ok: false, error: 'الصفحة غير موجودة' });
  }
  if (p.action === 'mailIn' && p.k) {
    const me = userByKey_(p.k);
    if (!me) return out_({ ok: false, auth: false });
    const regEmail = String(me[3] || '').trim().toLowerCase();
    const senderEmail = String(p.em || '').trim().toLowerCase();
    if (senderEmail && regEmail && senderEmail !== regEmail) {
      return out_({ ok: false, emailMismatch: true, expected: regEmail, actual: senderEmail });
    }
    let items = [];
    if (p.items) {
      try { items = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p.items)).getDataAsString('UTF-8')); } catch (err) {}
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      mailIn_(me, items);
      PropertiesService.getScriptProperties().setProperty('ML_' + me[0], String(Date.now()));
    } finally { lock.releaseLock(); }
    return out_({ ok: true, count: items.length, expected: regEmail });
  }
  return out_({ ok: true, v: VERSION, url: ScriptApp.getService().getUrl() });
}

function doPost(e) {
  const raw = e.postData.contents;
  let d;
  if (raw.indexOf('wa§§') === 0) {                         // إشعار واتساب من تطبيق الجوال
    const p = raw.split('§§');
    d = { action: 'waNotify', token: p[1], title: p[2], text: p.slice(3).join('§§') };
  } else d = JSON.parse(raw);
  const lock = LockService.getScriptLock();
  const needLock = d.action !== 'list' && d.action !== 'keys' && d.action !== 'poll' && d.action !== 'remList' && d.action !== 'msgOne' && d.action !== 'remDue' && d.action !== 'pushTest' && d.action !== 'archPage';   // القراءة لا تنتظر بعضها، فيسرع الموقع
  if (needLock) lock.waitLock(20000);
  try {
    if (e.parameter && e.parameter.wh) { handleWhapi_(e.parameter.wh, d); return out_({ ok: true }); }   // Whapi.Cloud
    if (d.object === 'whatsapp_business_account') { handleWhatsApp_(d, e.parameter && e.parameter.k); return out_({ ok: true }); }

    // ---- إنشاء حساب ----
    if (d.action === 'register') {
      const name = String(d.name || '').trim(), phone = norm_(d.phone), email = String(d.email || '').trim().toLowerCase(), pw = String(d.password || '');
      if (!name || phone.length < 11 || email.indexOf('@') < 1 || pw.length < 8)
        return out_({ ok: false, error: 'يرجى إدخال الاسم ورقم جوال صحيح وبريد صحيح وكلمة مرور لا تقل عن 8 أحرف.' });
      const dup = allUsers_().some(r => norm_(r[2]) === phone || String(r[3]).toLowerCase() === email);
      if (dup) return out_({ ok: false, error: 'رقم الجوال أو البريد مسجّل مسبقًا، يرجى تسجيل الدخول.' });
      const id = Utilities.getUuid().replace(/-/g, '').slice(0, 10), salt = Utilities.getUuid(), token = Utilities.getUuid();
      users_().appendRow([id, name, phone, email, hash_(pw, salt), salt, token, new Date()]); dropUsers_();
      return out_({ ok: true, token: newSession_(id), user: { id, name, phone, email } });
    }

    // ---- تسجيل الدخول ----
    if (d.action === 'login') {
      const idf = String(d.identifier || '').trim().toLowerCase(), n = norm_(idf);
      const cache = CacheService.getScriptCache();
      const lk = 'lf_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idf)).slice(0, 12);
      const tries = Number(cache.get(lk) || 0);
      if (tries >= 8) return out_({ ok: false, error: 'محاولات كثيرة، يرجى المحاولة بعد 15 دقيقة.' });
      const u = allUsers_().find(r => String(r[3]).toLowerCase() === idf || (n && norm_(r[2]) === n));
      if (!u || hash_(String(d.password || ''), u[5]) !== u[4]) { cache.put(lk, String(tries + 1), 900); return out_({ ok: false, error: 'بيانات الدخول غير صحيحة.' }); }
      cache.remove(lk);
      return out_({ ok: true, token: newSession_(u[0]), user: pub_(u) });
    }

    // ---- نسيت كلمة المرور: يُرسل رمز من 6 أرقام إلى البريد المسجّل (صالح 15 دقيقة) ----
    if (d.action === 'forgot') {
      const idf = String(d.identifier || '').trim().toLowerCase(), n = norm_(idf), cache = CacheService.getScriptCache();
      const u = idf && allUsers_().find(r => String(r[3]).toLowerCase() === idf || (n && norm_(r[2]) === n));
      if (u && !cache.get('rq_' + u[0])) {
        const code = String(100000 + Math.floor(Math.random() * 900000));
        cache.put('rq_' + u[0], '1', 60);
        cache.put('rs_' + u[0], JSON.stringify({ h: hash_(code, u[5]), t: 0 }), 900);
        sendMail_(u[3], 'رمز استعادة كلمة المرور', 'رمز التحقق: ' + code + '\n\nالرمز صالح لمدة 15 دقيقة. إن لم تطلبي استعادة كلمة المرور فتجاهلي هذه الرسالة، ولن يتغير شيء.');
      }
      return out_({ ok: true });   // نفس الرد دائمًا حتى لا يُعرف هل الحساب موجود
    }
    if (d.action === 'reset') {
      const idf = String(d.identifier || '').trim().toLowerCase(), n = norm_(idf), cache = CacheService.getScriptCache();
      const u = idf && allUsers_().find(r => String(r[3]).toLowerCase() === idf || (n && norm_(r[2]) === n));
      const bad = out_({ ok: false, error: 'الرمز غير صحيح أو منتهي.' });
      const raw = u && cache.get('rs_' + u[0]);
      if (!raw) return bad;
      const pw = String(d.password || '');
      if (pw.length < 8) return out_({ ok: false, error: 'كلمة المرور الجديدة لا تقل عن 8 أحرف.' });
      const st = JSON.parse(raw);
      if (hash_(String(d.code || '').trim(), u[5]) !== st.h) {
        st.t++; if (st.t >= 5) cache.remove('rs_' + u[0]); else cache.put('rs_' + u[0], JSON.stringify(st), 900);
        return bad;
      }
      setPw_(u, pw); endSessions_(u[0]); cache.remove('rs_' + u[0]);
      sendMail_(u[3], 'تم تغيير كلمة المرور', 'تم تغيير كلمة مرور حسابك في صندوق الرسائل. إن لم تكوني أنتِ، فتواصلي مع المسؤول فورًا.');
      return out_({ ok: true });
    }

    // ---- رسالة من زائر إلى صاحبة الصفحة ----
    if (d.action === 'add') {
      const owner = userById_(d.owner);
      if (!owner) return out_({ ok: false, error: 'الصفحة غير موجودة.' });
      if (!d.name || !d.message) return out_({ ok: false, error: 'الاسم والرسالة مطلوبان.' });
      const name = String(d.name).slice(0, 200), message = String(d.message).slice(0, 5000);
      const phone = String(d.phone || '').slice(0, 50), email = String(d.email || '').slice(0, 200);
      msgs_().appendRow([Utilities.getUuid(), new Date(), enc_(name), enc_(phone), enc_(email), enc_(message), 'جديدة', '', 'النموذج', owner[0], '']);
      touch_(owner[0]);
      notify_(owner[3], name, phone, email, message, 'النموذج');
      return out_({ ok: true });
    }

    // ---- ما يلزم له تسجيل دخول ----
    const me = (d.action === 'waNotify' || d.action === 'mailIn') ? userByKey_(d.token) : userBySession_(d.token);
    if (!me) return out_({ ok: false, error: 'انتهت الجلسة، يرجى تسجيل الدخول.', auth: false });

    if (d.action === 'poll') return out_({ ok: true, ch: chOf_(me[0]), rc: remChOf_(me[0]), mail: Number(PropertiesService.getScriptProperties().getProperty('ML_' + me[0]) || 0) });
    if (d.action === 'keys') return out_({ ok: true, key: me[6] });
    if (d.action === 'logout') { if (d.endpoint) setSubs_(me[0], subs_(me[0]).filter(s => s.e !== String(d.endpoint))); dropSession_(d.token); return out_({ ok: true }); }

    if (d.action === 'changePassword') {
      const np = String(d.password || '');
      if (hash_(String(d.current || ''), me[5]) !== me[4]) return out_({ ok: false, error: 'كلمة المرور الحالية غير صحيحة.' });
      if (np.length < 8) return out_({ ok: false, error: 'كلمة المرور الجديدة لا تقل عن 8 أحرف.' });
      setPw_(me, np); endSessions_(me[0]);
      sendMail_(me[3], 'تم تغيير كلمة المرور', 'تم تغيير كلمة مرور حسابك في صندوق الرسائل. إن لم تكوني أنتِ، فاستعيديها فورًا من خيار «نسيت كلمة المرور».');
      return out_({ ok: true, token: newSession_(me[0]) });
    }

    // ---- تغيير رقم الجوال ----
    if (d.action === 'changePhone') {
      const bad = pwCheck_(me, d.current); if (bad) return out_({ ok: false, error: bad });
      const phone = norm_(d.phone);
      if (phone.length < 11) return out_({ ok: false, error: 'يرجى إدخال رقم جوال صحيح.' });
      if (phone === norm_(me[2])) return out_({ ok: false, error: 'هذا هو رقم جوالك الحالي.' });
      if (allUsers_().some(r => r[0] !== me[0] && norm_(r[2]) === phone)) return out_({ ok: false, error: 'رقم الجوال مسجّل لحساب آخر.' });
      if (!setContact_(me, 3, phone)) return out_({ ok: false, error: 'تعذّر الحفظ.' });
      sendMail_(me[3], 'تم تغيير رقم الجوال', 'تم تغيير رقم الجوال المسجّل في حسابك بصندوق الرسائل إلى: ' + phone + '\nإن لم تكوني أنتِ، فاستعيدي كلمة المرور فورًا من خيار «نسيت كلمة المرور».');
      return out_({ ok: true, user: { name: me[1], phone: phone, email: me[3] } });
    }

    // ---- تغيير البريد: خطوة 1 (إرسال رمز تحقق إلى البريد الجديد) ----
    if (d.action === 'emailCode') {
      const bad = pwCheck_(me, d.current); if (bad) return out_({ ok: false, error: bad });
      const em = String(d.email || '').trim().toLowerCase(), cache = CacheService.getScriptCache();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return out_({ ok: false, error: 'يرجى إدخال بريد إلكتروني صحيح.' });
      if (em === String(me[3]).toLowerCase()) return out_({ ok: false, error: 'هذا هو بريدك الحالي.' });
      if (allUsers_().some(r => r[0] !== me[0] && String(r[3]).toLowerCase() === em)) return out_({ ok: false, error: 'البريد مسجّل لحساب آخر.' });
      if (cache.get('eq_' + me[0])) return out_({ ok: false, error: 'تم إرسال رمز قبل قليل، انتظري دقيقة ثم أعيدي المحاولة.' });
      const code = String(100000 + Math.floor(Math.random() * 900000));
      cache.put('eq_' + me[0], '1', 60);
      cache.put('ec_' + me[0], JSON.stringify({ h: hash_(code, me[5]), e: em, t: 0 }), 900);
      sendMail_(em, 'رمز تأكيد البريد الجديد', 'رمز التحقق: ' + code + '\n\nالرمز صالح لمدة 15 دقيقة. إن لم تطلبي تغيير البريد فتجاهلي هذه الرسالة.');
      return out_({ ok: true });
    }
    // ---- تغيير البريد: خطوة 2 (التحقق من الرمز ثم الحفظ) ----
    if (d.action === 'changeEmail') {
      const cache = CacheService.getScriptCache(), raw = cache.get('ec_' + me[0]);
      if (!raw) return out_({ ok: false, error: 'الرمز غير صحيح أو منتهي.' });
      const st = JSON.parse(raw);
      if (hash_(String(d.code || '').trim(), me[5]) !== st.h) {
        st.t++; if (st.t >= 5) cache.remove('ec_' + me[0]); else cache.put('ec_' + me[0], JSON.stringify(st), 900);
        return out_({ ok: false, error: 'الرمز غير صحيح أو منتهي.' });
      }
      if (allUsers_().some(r => r[0] !== me[0] && String(r[3]).toLowerCase() === st.e)) { cache.remove('ec_' + me[0]); return out_({ ok: false, error: 'البريد مسجّل لحساب آخر.' }); }
      const old = me[3];
      if (!setContact_(me, 4, st.e)) return out_({ ok: false, error: 'تعذّر الحفظ.' });
      cache.remove('ec_' + me[0]);
      sendMail_(old, 'تم تغيير البريد الإلكتروني', 'تم تغيير البريد المسجّل في حسابك بصندوق الرسائل إلى: ' + st.e + '\nإن لم تكوني أنتِ، فتواصلي مع المسؤول فورًا.');
      return out_({ ok: true, user: { name: me[1], phone: me[2], email: st.e } });
    }

    if (d.action === 'waNotify') {
      const title = String(d.title || '').trim(), text = String(d.text || '').trim();
      if (!title || !text) return out_({ ok: true });
      if (/^(whatsapp|واتساب)/i.test(title) || /\d+\s*(new messages|رسائل جديدة|رسالة جديدة)/i.test(text)) return out_({ ok: true });   // إشعارات التجميع
      if (isGoneC_(cid_(me[0], title, text))) return out_({ ok: true });   // رسالة حُذفت قبل قليل: لا تعود
      const sh = msgs_(), n = sh.getLastRow() - 1;
      const recent = n > 0 ? sh.getRange(Math.max(2, n - 98), 1, Math.min(n, 100), COLS).getValues() : [];
      if (recent.some(r => r[9] === me[0] && r[6] !== DEL && dec_(r[2]) === title && dec_(r[5]) === text)) return out_({ ok: true });   // منع التكرار
      sh.appendRow([Utilities.getUuid(), new Date(), enc_(title), '', '', enc_(text), 'جديدة', '', 'واتساب', me[0], '']);
      touch_(me[0]);
      return out_({ ok: true });
    }

    if (d.action === 'mailIn') {
      const regEmail = String(me[3] || '').trim().toLowerCase();
      const senderEmail = String(d.em || '').trim().toLowerCase();
      if (senderEmail && regEmail && senderEmail !== regEmail) {
        return out_({ ok: false, emailMismatch: true, expected: regEmail, actual: senderEmail });
      }
      mailIn_(me, d.items);
      if (Array.isArray(d.items) && d.items.length) touch_(me[0]);
      PropertiesService.getScriptProperties().setProperty('ML_' + me[0], String(Date.now()));
      return out_({ ok: true, expected: regEmail });
    }

    if (d.action === 'addManual') {
      if (!d.name) return out_({ ok: false, error: 'الاسم مطلوب.' });
      msgs_().appendRow([Utilities.getUuid(), new Date(), enc_(String(d.name).slice(0, 200)), enc_(norm_(d.phone)), '', enc_(String(d.message || '').slice(0, 5000)), 'جديدة', '', 'واتساب', me[0], '']);
      touch_(me[0]);
      return out_({ ok: true });
    }

    // ---- جلب محادثة واحدة (لفتحها من التذكير) ----
    if (d.action === 'msgOne') {
      const f = findRow_(msgs_(), String(d.id || ''), me);
      return f ? out_({ ok: true, row: rowOut_(f.vals) }) : out_({ ok: false, error: 'الرسالة غير موجودة.' });
    }

    // ---- تنبيهات الجوال ----
    if (d.action === 'pushSub') {
      const e = String((d.sub && d.sub.endpoint) || d.endpoint || '');
      if (!/^https:\/\//.test(e) || e.length > 700) return out_({ ok: false, error: 'اشتراك غير صالح.' });
      const mine = subs_(me[0]);
      if (!mine.some(s => s.e === e)) {   // جهاز جديد لهذا الحساب: يُفصل عن أي حساب آخر سُجّل به سابقًا (دخلت منه مستخدمة أخرى)
        allUsers_().forEach(u => { if (u[0] !== me[0]) { const o = subs_(u[0]); if (o.some(s => s.e === e)) setSubs_(u[0], o.filter(s => s.e !== e)); } });
      }
      const a = mine.filter(s => s.e !== e); a.push({ e: e, t: Date.now() });
      setSubs_(me[0], a.slice(-5));
      return out_({ ok: true });
    }
    if (d.action === 'pushUnsub') { setSubs_(me[0], subs_(me[0]).filter(s => s.e !== String(d.endpoint || ''))); return out_({ ok: true }); }
    if (d.action === 'pushTest') {
      PropertiesService.getScriptProperties().setProperty('PT_' + me[0], String(Date.now()));
      const n = subs_(me[0]).length; if (!n) return out_({ ok: false, error: 'لا يوجد جهاز مفعّل.' });
      const okN = pushAll_(me[0], {});
      return out_(okN ? { ok: true, sent: okN } : { ok: false, error: 'تعذّر الإرسال. تأكدي من تشغيل installPushTrigger ومنح الصلاحيات ثم نشر إصدار جديد.' });
    }
    if (d.action === 'remDue') {   // يستدعيها الـ Service Worker عند وصول إشعار
      const sh = rems_(), n = sh.getLastRow() - 1, now = Date.now(), rows = [];
      if (n > 0) sh.getRange(2, 1, n, RCOLS).getValues().forEach(r => { if (r[2] === me[0] && r[4] !== 'مكتمل' && Number(r[14]) > now - 180000) rows.push(r); });
      const pt = Number(PropertiesService.getScriptProperties().getProperty('PT_' + me[0]) || 0);
      return out_({ ok: true, rows: rows.slice(0, 5).map(remOut_), test: now - pt < 180000 });
    }

    // ---- التذكيرات والمتابعات ----
    if (d.action === 'remList') {
      // كل التذكيرات المعلّقة + آخر 100 مكتملة (ويُفكّ تشفير المعروض فقط)
      const sh = rems_(), n = sh.getLastRow() - 1, pend = [], done = [];
      if (n > 0) sh.getRange(2, 1, n, RCOLS).getValues().forEach(r => { if (r[2] === me[0]) (r[4] === 'مكتمل' ? done : pend).push(r); });
      pend.sort((a, b) => (Number(a[3]) || 0) - (Number(b[3]) || 0));
      done.sort((a, b) => (Number(b[12]) || 0) - (Number(a[12]) || 0));
      return out_({ ok: true, rows: pend.slice(0, 500).concat(done.slice(0, 100)).map(remOut_), doneTotal: done.length, rc: remChOf_(me[0]) });
    }
    if (d.action === 'remAdd') {
      const title = String(d.title || '').trim().slice(0, 200), due = Number(d.due);
      if (!title) return out_({ ok: false, error: 'عنوان التذكير مطلوب.' });
      if (!dueOk_(due)) return out_({ ok: false, error: 'موعد التذكير غير صحيح.' });
      let name = String(d.name || '').trim().slice(0, 200), phone = norm_(d.phone), email = '', source = '', msgId = '';
      if (d.msgId) {   // بيانات الشخص تُؤخذ من الرسالة نفسها في الخادم (ولا تُقبل إلا رسالة تخص هذه المستخدمة)
        const f = findRow_(msgs_(), String(d.msgId), me);
        if (!f) return out_({ ok: false, error: 'المحادثة غير موجودة.' });
        msgId = String(f.vals[0]); name = dec_(f.vals[2]); phone = dec_(f.vals[3]); email = dec_(f.vals[4]); source = f.vals[8] || 'النموذج';
      }
      const row = [Utilities.getUuid(), new Date(), me[0], due, 'معلّق', enc_(title), enc_(String(d.note || '').slice(0, 2000)), msgId, enc_(name), enc_(phone), enc_(email), source, ''];
      rems_().appendRow(row);
      return out_({ ok: true, rem: remOut_(row), rc: rtouch_(me[0]) });
    }
    if (d.action === 'remUpdate') {
      const sh = rems_(), f = findRem_(sh, String(d.id || ''), me);
      if (!f) return out_({ ok: false, error: 'التذكير غير موجود.' });
      const row = f.vals;
      if (d.title !== undefined) { const t = String(d.title).trim().slice(0, 200); if (!t) return out_({ ok: false, error: 'عنوان التذكير مطلوب.' }); row[5] = enc_(t); }
      if (d.note !== undefined) row[6] = enc_(String(d.note).slice(0, 2000));
      if (d.due !== undefined) { if (!dueOk_(d.due)) return out_({ ok: false, error: 'موعد التذكير غير صحيح.' }); row[3] = Number(d.due); }
      if (d.done !== undefined) { row[4] = d.done ? 'مكتمل' : 'معلّق'; row[12] = d.done ? Date.now() : ''; }
      if (!row[7]) {   // التذكير اليدوي فقط يمكن تعديل بيانات الشخص فيه
        if (d.name !== undefined) row[8] = enc_(String(d.name).trim().slice(0, 200));
        if (d.phone !== undefined) row[9] = enc_(norm_(d.phone));
      }
      sh.getRange(f.row, 1, 1, RCOLS).setValues([row]);
      return out_({ ok: true, rc: rtouch_(me[0]) });
    }
    if (d.action === 'remDelete') {
      const sh = rems_(), f = findRem_(sh, String(d.id || ''), me);
      if (f) sh.deleteRow(f.row);
      return out_({ ok: true, rc: rtouch_(me[0]) });
    }

    if (d.action === 'list') {
      // يُرجع أحدث رسائل العرض المطلوب فقط (الافتراضي 100) مع أعداد الحالات كاملة، ويفك تشفير المعروض فقط
      if (d.archive) return listArchive_(me, d);   // عرض الرسائل المؤرشفة (للقراءة والحذف فقط)
      const first = allUsers_()[0][0];
      const sh = msgs_(), wantStarred = !!(d.starred || d.archived), limit = Math.min(Math.max(Number(d.limit) || 100, 1), 1000);
      cleanDeleted_(sh);
      let starredCount = 0;
      const mine = [], counts = {};
      if (sh.getLastRow() >= 2) {
        sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach(r => {
          if ((r[9] || first) !== me[0] || r[6] === DEL) return;       // الرسائل القديمة بلا مالكة تعود لأول حساب
          const isStar = r[10] === 'نعم';                              // عمود «نعم» القديم (الأرشيف) صار يعني: مميزة بنجمة
          if (isStar) starredCount++;
          counts[r[6]] = (counts[r[6]] || 0) + 1;
          if (!wantStarred || isStar) mine.push(r);
        });
      }
      const ord = new Map(mine.map((r, i) => [r, i]));
      mine.sort((a, b) => (new Date(b[1]).getTime() || 0) - (new Date(a[1]).getTime() || 0) || ord.get(b) - ord.get(a));   // الأحدث أولًا
      return out_({ ok: true, v: VERSION, rows: mine.slice(0, limit).map(rowOut_), starredCount, counts, total: mine.length, archDays: ARCHIVE_DAYS, archKeep: ARCHIVE_KEEP_DAYS, archWarn: archWarn_(me[0]), ch: chOf_(me[0]), mail: Number(PropertiesService.getScriptProperties().getProperty('ML_' + me[0]) || 0) });
    }

    // ---- تحميل الرسائل المؤرشفة (صفحات) ----
    if (d.action === 'archPage') {
      const all = archMine_(me), off = Math.max(0, Number(d.offset) || 0), lim = Math.min(Math.max(Number(d.limit) || 200, 1), 300);
      return out_({ ok: true, total: all.length, rows: all.slice(off, off + lim).map(archOut_) });
    }

    // ---- حذف رسالة مؤرشفة نهائيًا ----
    if (d.action === 'archDelete') {
      const first = allUsers_()[0][0];
      for (const id of archIds_()) {
        try {
          const ash = SpreadsheetApp.openById(id).getSheetByName('الأرشيف');
          if (!ash || ash.getLastRow() < 2) continue;
          const vals = ash.getRange(2, 1, ash.getLastRow() - 1, COLS).getValues();
          const i = vals.findIndex(r => r[0] === d.id && (r[9] || first) === me[0]);
          if (i >= 0) { ash.deleteRow(i + 2); setArchWarn_(me[0], archWarnCalc_(archMine_(me))); return out_({ ok: true }); }
        } catch (e) {}
      }
      return out_({ ok: false, error: 'الرسالة غير موجودة.' });
    }

    if (d.action === 'update') {
      const sh = msgs_(), f = findRow_(sh, d.id, me);
      if (!f) return out_({ ok: false, error: 'الرسالة غير موجودة.' });
      if (d.status !== undefined && STATUSES_.indexOf(d.status) < 0) return out_({ ok: false, error: 'حالة غير صحيحة.' });
      if (d.status !== undefined) sh.getRange(f.row, 7).setValue(d.status);
      if (d.note !== undefined) sh.getRange(f.row, 8).setValue(enc_(String(d.note).slice(0, 2000)));
      const st = d.starred !== undefined ? d.starred : d.archived;
      if (st !== undefined) sh.getRange(f.row, 11).setValue(st ? 'نعم' : '');
      touch_(me[0]);
      return out_({ ok: true });
    }

    if (d.action === 'delete') {
      const sh = msgs_(), f = findRow_(sh, d.id, me);
      if (!f) return out_({ ok: false, error: 'الرسالة غير موجودة.' });
      buryRows_([f.vals], me[0]); sh.deleteRow(f.row); touch_(me[0]);          // حذف الصف كاملًا من الشيت، بلا أي أثر نصي
      return out_({ ok: true });
    }
    if (d.action === 'bulk') {
      // إجراء جماعي على رسائل محددة: status (تم الرد...) | star (تمييز/إزالة) | delete
      const ids = (Array.isArray(d.ids) ? d.ids : []).map(String).slice(0, 1000), set = new Set(ids);
      if (!set.size) return out_({ ok: true, count: 0 });
      if (['status', 'star', 'delete'].indexOf(d.op) < 0) return out_({ ok: false, error: 'إجراء غير صحيح.' });
      if (d.op === 'status' && STATUSES_.indexOf(d.value) < 0) return out_({ ok: false, error: 'حالة غير صحيحة.' });
      const sh = msgs_();
      if (sh.getLastRow() < 2) return out_({ ok: true, count: 0 });
      const first = allUsers_()[0][0], hit = [];
      sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach((r, i) => {
        if ((r[9] || first) === me[0] && set.has(String(r[0]))) hit.push({ row: i + 2, vals: r });
      });
      if (d.op === 'delete') {
        if (hit.length) { buryRows_(hit.map(h => h.vals), me[0]); hit.slice().reverse().forEach(h => sh.deleteRow(h.row)); }
      } else {
        hit.forEach(h => d.op === 'status' ? sh.getRange(h.row, 7).setValue(d.value) : sh.getRange(h.row, 11).setValue(d.value ? 'نعم' : ''));
      }
      if (hit.length) touch_(me[0]);
      return out_({ ok: true, count: hit.length });
    }
    if (d.action === 'deleteReplied') {
      const sh = msgs_();
      if (sh.getLastRow() < 2) return out_({ ok: true, deleted: 0 });
      const first = allUsers_()[0][0], ids = [], rows = [];
      sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach((r, i) => {
        if ((r[9] || first) !== me[0] || r[6] !== 'تم الرد' || r[10] === 'نعم') return;
        ids.push(r); rows.push(i + 2);
      });
      if (ids.length) { buryRows_(ids, me[0]); rows.reverse().forEach(x => sh.deleteRow(x)); touch_(me[0]); }
      return out_({ ok: true, deleted: ids.length });
    }
    return out_({ ok: false, error: 'طلب غير معروف.' });
  } finally { if (needLock) lock.releaseLock(); }
}

// بريد المستخدمة: يصل من سكربت تشغّله هي في حسابها (Gmail) فلا يُقرأ بريد أي حساب آخر
function mailIn_(me, items) {
  items = Array.isArray(items) ? items.slice(0, 50) : [];
  if (!items.length) return;                                        // لا شيء لمعالجته: لا نقرأ الشيت أصلًا
  const sh = msgs_(), n = sh.getLastRow() - 1;
  const idCol = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [];   // عمود المعرّفات فقط (أسرع بكثير من قراءة كل الأعمدة)
  const index = {};
  idCol.forEach((r, i) => { index[r[0]] = i; });
  items.forEach(it => {
    if (!it || !it.id) return;
    const id = 'mail_' + me[0] + '_' + String(it.id).slice(0, 40);
    if (isGone_(id)) return;
    const when = it.date ? new Date(Number(it.date)) : new Date();
    const text = (String(it.subject || '') + '\n' + String(it.body || '')).slice(0, 1000);
    if (id in index) {
      const i = index[id], row = sh.getRange(i + 2, 1, 1, COLS).getValues()[0];   // نقرأ الصف المطلوب فقط
      if (row[9] !== me[0] || row[6] === DEL) return;
      if (it.fromMe) { if (AUTO_REPLIED && row[6] !== 'تم الرد') sh.getRange(i + 2, 7).setValue('تم الرد'); }
      else if ((row[6] === 'تم الرد' || row[10] === 'نعم') && when > new Date(row[1])) {   // رسالة جديدة بعد ردّك أو بعد الأرشفة
        sh.getRange(i + 2, 7).setValue('جديدة');
        sh.getRange(i + 2, 2).setValue(when);
        sh.getRange(i + 2, 6).setValue(enc_(text));
        sh.getRange(i + 2, 11).setValue('');
      }
      return;
    }
    if (it.fromMe) return;
    const from = String(it.from || '').slice(0, 300);
    const name = from.replace(/<.*>/, '').replace(/"/g, '').trim() || from;
    const mail = (from.match(/<(.+?)>/) || [])[1] || from;
    sh.appendRow([id, when, enc_(name), '', enc_(mail), enc_(text), 'جديدة', '', 'البريد', me[0], '']);
  });
}

function handleWhatsApp_(d, key) {
  const owner = key ? userByKey_(key) : null;          // الرابط الخاص بكل حساب يحمل مفتاحه (?k=...)
  if (!owner) return;
  const sh = msgs_();
  let ids = null;
  (d.entry || []).forEach(entry => (entry.changes || []).forEach(ch => {
    const v = ch.value || {};

    (v.messages || []).forEach(m => {
      if (!ids) ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];   // مرة واحدة لكل طلب
      if (ids.indexOf(m.id) >= 0 || isGone_(m.id)) return;
      ids.push(m.id);
      const c = (v.contacts || []).find(x => x.wa_id === m.from);
      const name = (c && c.profile && c.profile.name) || m.from;
      let text = m.type === 'text' ? m.text.body : m.type === 'button' ? m.button.text : '[رسالة من نوع: ' + m.type + ']';
      const ts = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
      sh.appendRow([m.id, ts, enc_(name), enc_(m.from), '', enc_(text), 'جديدة', '', 'واتساب', owner[0], '']);
    });

    (v.message_echoes || []).forEach(m => {
      const to = norm_(m.to);
      if (!AUTO_REPLIED || !to || sh.getLastRow() < 2) return;
      sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach((r, i) => {
        if (r[8] === 'واتساب' && r[9] === owner[0] && r[6] !== 'تم الرد' && r[6] !== DEL && norm_(dec_(r[3])) === to) sh.getRange(i + 2, 7).setValue('تم الرد');
      });
    });
  }));
  touch_(owner[0]);
}

// الإشعار يُرسل دائمًا إلى البريد المسجّل لصاحبة الصندوق فقط، ولا تُرسل نسخة إلى أي بريد آخر.
// الردود لا تعود إلى بريد حساب النشر: تذهب إلى بريد الزائر إن وُجد، وإلا يُرسل الإشعار بلا خاصية الرد.
function notify_(to, name, phone, email, message, source) {
  if (source !== 'النموذج') return;   // لا إشعارات بريد لرسائل الواتساب
  to = String(to || '').trim();
  if (to.indexOf('@') < 1) return;
  const opt = { name: 'صندوق الرسائل' };
  const reply = String(email || '').trim();
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(reply)) opt.replyTo = reply; else opt.noReply = true;
  try {
    MailApp.sendEmail(to, 'رسالة جديدة (' + source + ') من ' + name,
      'الاسم: ' + name + '\nالجوال: ' + (phone || '-') + '\nالبريد: ' + (email || '-') + '\n\n' + message, opt);
  } catch (err) { Logger.log('تعذّر إرسال إشعار البريد: ' + err); }
}


// ===== سجل منع العودة =====
// عند الحذف يُمسح الصف كاملًا من الشيت، وتُحفظ بصمة مقتطعة (Hash) لمعرّف الرسالة فقط في خصائص السكربت (ليس في الشيت).
// السجل مقسّم بالأيام، فلا يمتلئ ولا يُمسح بعضه ببعض مهما كثر الحذف (كان سابقًا سقفه 450 بصمة لكل المستخدمات معًا).
// معرّفات الرسائل تُحفظ 9 أيام (Meta قد تعيد الإرسال حتى 7 أيام)، وبصمات نص إشعارات الجوال 4 أيام.
const TOMB_DAYS = 9, TOMBC_DAYS = 4;
function tb_(id) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'tb:' + id)).slice(0, 8); }
function day_() { return Math.floor(Date.now() / 86400000); }
let TS_ = {};   // ذاكرة داخل التنفيذ الواحد
function tombSet_(pre, days) {
  const k = pre + days;
  if (TS_[k]) return TS_[k];
  const P = PropertiesService.getScriptProperties(), set = {}, d0 = day_();
  for (let d = d0; d > d0 - days; d--) for (let n = 0; ; n++) {
    const v = P.getProperty(pre + d + '.' + n); if (v == null) break;
    for (let i = 0; i + 8 <= v.length; i += 8) set[v.substr(i, 8)] = 1;
  }
  try {   // السجل القديم (قبل هذا الإصدار) يبقى معتمدًا حتى لا يعود ما حُذف قريبًا
    JSON.parse(P.getProperty(pre === 'T' ? 'TOMB' : 'TOMBC') || '[]').forEach(x => { set[String(x).slice(0, 8)] = 1; });
  } catch (e) {}
  return (TS_[k] = set);
}
function addTombs_(hs, pre, days) {
  const P = PropertiesService.getScriptProperties(), d = day_(), set = tombSet_(pre, days);
  let n = 0, cur = '';
  while (true) { cur = P.getProperty(pre + d + '.' + n) || ''; if (cur.length + 8 <= 8000) break; n++; }   // الخاصية الواحدة تتسع لنحو 1000 بصمة
  hs.forEach(h => {
    set[h] = 1;
    if (cur.length + 8 > 8000) { P.setProperty(pre + d + '.' + n, cur); n++; cur = ''; }
    cur += h;
  });
  P.setProperty(pre + d + '.' + n, cur);
  for (let od = d - days - 2; od <= d - days; od++) for (let m = 0; P.getProperty(pre + od + '.' + m) != null; m++) P.deleteProperty(pre + od + '.' + m);   // تنظيف الأيام المنتهية
}
function isGone_(id) { return !!tombSet_('T', TOMB_DAYS)[tb_(id)]; }
function bury_(ids) { if (ids.length) addTombs_(ids.map(tb_), 'T', TOMB_DAYS); }
// إشعارات الواتساب القادمة من تطبيق الجوال ليس لها معرّف ثابت، فنمنع عودتها بحسب (المالكة + المرسل + النص).
function cid_(owner, name, text) { return owner + '|' + name + '|' + text; }
function isGoneC_(c) { return !!tombSet_('C', TOMBC_DAYS)[tb_(c)]; }
function buryRows_(rows, owner) {   // rows: صفوف الشيت المحذوفة (مصفوفات)
  bury_(rows.map(r => r[0]));
  const c = rows.filter(r => r[8] === 'واتساب').map(r => cid_(owner, dec_(r[2]), dec_(r[5])));
  if (c.length) addTombs_(c.map(tb_), 'C', TOMBC_DAYS);
}
// تنظيف الصفوف القديمة المكتوب فيها «محذوفة»: تُحذف نهائيًا من الشيت
function cleanDeleted_(sh) {
  const P = PropertiesService.getScriptProperties(); if (P.getProperty('CLEAN1')) return;   // تنظيف لمرة واحدة فقط
  const n = sh.getLastRow() - 1; if (n < 1) return;
  const gone = [], rows = [];
  sh.getRange(2, 1, n, 7).getValues().forEach((r, i) => { if (r[6] === DEL) { gone.push(r[0]); rows.push(i + 2); } });
  if (gone.length) { bury_(gone); rows.reverse().forEach(x => sh.deleteRow(x)); }
  P.setProperty('CLEAN1', '1');
}

// ===== كلمة المرور =====
function sendMail_(to, subject, body) {
  try { MailApp.sendEmail(to, subject, body, { name: 'صندوق الرسائل', noReply: true }); } catch (err) { Logger.log('تعذّر إرسال البريد: ' + err); }
}
// تغيير بيانات التواصل (الجوال/البريد): الرسائل والتذكيرات مرتبطة بمعرّف الحساب (العمود الأول) لا بالجوال ولا بالبريد، فلا يضيع منها شيء.
function setContact_(u, col, val) {
  const i = allUsers_().findIndex(r => r[0] === u[0]); if (i < 0) return false;
  users_().getRange(i + 2, col).setValue(val); dropUsers_(); return true;
}
// فحص كلمة المرور الحالية مع حدّ للمحاولات (5 كل 15 دقيقة). يعيد نص الخطأ أو '' عند النجاح
function pwCheck_(me, pw) {
  const c = CacheService.getScriptCache(), k = 'pf_' + me[0], t = Number(c.get(k) || 0);
  if (t >= 5) return 'محاولات كثيرة، يرجى المحاولة بعد 15 دقيقة.';
  if (hash_(String(pw || ''), me[5]) !== me[4]) { c.put(k, String(t + 1), 900); return 'كلمة المرور الحالية غير صحيحة.'; }
  c.remove(k); return '';
}
function setPw_(u, pw) {
  const i = allUsers_().findIndex(r => r[0] === u[0]); if (i < 0) return;
  const salt = Utilities.getUuid();
  users_().getRange(i + 2, 5, 1, 2).setValues([[hash_(pw, salt), salt]]); dropUsers_();
}

// ===== تشفير الرسائل القديمة (شغّليها مرة واحدة فقط بعد تحديث الكود) =====
function migrateEncrypt() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = msgs_(), n = sh.getLastRow() - 1;
    if (n < 1) { Logger.log('لا توجد رسائل لتشفيرها.'); return; }
    const rng = sh.getRange(2, 1, n, COLS), vals = rng.getValues();
    let done = 0;
    vals.forEach(r => {
      [2, 3, 4, 5, 7].forEach(c => {
        const v = String(r[c] == null ? '' : r[c]);
        if (v && v.indexOf('enc1:') !== 0) { r[c] = enc_(v); done++; }
      });
    });
    rng.setValues(vals);
    Logger.log('تم تشفير ' + done + ' خانة.');
  } finally { lock.releaseLock(); }
}


// ===== استقبال الواتساب من Whapi.Cloud (ربط بمسح QR) =====
function handleWhapi_(token, d) {
  const me = userByKey_(token);
  if (!me) return;
  const sh = msgs_();
  (d.messages || []).forEach(m => {
    const chat = String(m.chat_id || '');
    if (chat.indexOf('@s.whatsapp.net') < 0) return;                 // المحادثات الفردية فقط
    const phone = chat.split('@')[0];
    const n = sh.getLastRow() - 1;
    const vals = n > 0 ? sh.getRange(2, 1, n, COLS).getValues() : [];

    if (m.from_me) {                                                 // ردّك من الجوال: تُعلَّم المحادثة «تم الرد»
      if (AUTO_REPLIED) vals.forEach((r, i) => {
        if (r[9] === me[0] && r[8] === 'واتساب' && r[6] !== 'تم الرد' && r[6] !== DEL && norm_(dec_(r[3])) === phone) sh.getRange(i + 2, 7).setValue('تم الرد');
      });
      return;
    }
    if (vals.some(r => r[0] === m.id) || isGone_(m.id)) return;                       // منع التكرار
    const text = m.type === 'text' ? ((m.text && m.text.body) || '') : '[رسالة من نوع: ' + m.type + ']';
    const when = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
    sh.appendRow([m.id, when, enc_(m.from_name || phone), enc_(phone), '', enc_(text), 'جديدة', '', 'واتساب', me[0], '']);
  });
  touch_(me[0]);
}
// ===== للمسؤولة: إعادة تعيين كلمة مرور أي مستخدمة بدون بريد =====
// غيّري القيمتين ثم اختاري adminResetPassword وشغّليها من المحرر، ثم امسحي كلمة المرور من الكود.
function adminResetPassword() {
  const IDENT = 'ضعي-هنا-البريد-أو-الجوال', NEW_PASSWORD = 'ضعي-هنا-كلمة-مرور-جديدة';
  const idf = IDENT.trim().toLowerCase(), n = norm_(idf);
  const u = allUsers_().find(r => String(r[3]).toLowerCase() === idf || (n && norm_(r[2]) === n));
  if (!u) { Logger.log('لا يوجد حساب بهذه البيانات.'); return; }
  if (NEW_PASSWORD.length < 8) { Logger.log('كلمة المرور لا تقل عن 8 أحرف.'); return; }
  setPw_(u, NEW_PASSWORD); endSessions_(u[0]);
  Logger.log('تم تعيين كلمة مرور جديدة للحساب: ' + u[1] + '. يمكنها الدخول بها الآن.');
}

function diag() {
  Logger.log('الرصيد اليومي المتبقي: ' + MailApp.getRemainingDailyQuota());
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'تجربة', 'وصلت الرسالة');
}


// ===== الأرشفة التلقائية: تحافظ على سرعة الموقع مهما كثرت الرسائل =====
// الرسائل الأقدم من ARCHIVE_DAYS يومًا (وغير المميّزة بنجمة) تُنقل تلقائيًا كل يوم إلى ملف أرشيف منفصل
// (بنفس التشفير، وبلا حد المليون خلية للملف الرئيسي)، فيبقى شيت الرسائل صغيرًا وسريعًا.
// لا تظهر الرسائل المؤرشفة في الموقع. المميّزة بنجمة تبقى دائمًا.
// تشغيل مرة واحدة فقط من المحرر: installArchiveTrigger  (ثم امنحي الصلاحيات)
const ARCHIVE_DAYS = 90, ARCHIVE_MAX_ROWS = 800000;
const ARCHIVE_KEEP_DAYS = 30, ARCHIVE_WARN_DAYS = 7, REM_KEEP_DAYS = 30, ARCH_COLS = COLS + 1;   // بقاء الأرشيف / مدة التنبيه قبل الحذف / بقاء التذكيرات المكتملة

function archSheet_(need) {
  const P = PropertiesService.getScriptProperties();
  let id = P.getProperty('ARCH_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  let sh = ss && ss.getSheetByName('الأرشيف');
  if (sh && sh.getLastRow() + need > ARCHIVE_MAX_ROWS) { ss = null; sh = null; }   // امتلأ ملف الأرشيف → ملف جديد
  if (!sh) {
    ss = SpreadsheetApp.create('أرشيف صندوق الرسائل - ' + Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd'));
    sh = ss.getSheets()[0]; sh.setName('الأرشيف');
    sh.appendRow(msgs_().getRange(1, 1, 1, COLS).getValues()[0].concat(['تاريخ الأرشفة'])); sh.setFrozenRows(1);
    P.setProperty('ARCH_ID', ss.getId());
    P.setProperty('ARCH_ALL', (P.getProperty('ARCH_ALL') ? P.getProperty('ARCH_ALL') + ',' : '') + ss.getId());   // سجل بكل ملفات الأرشيف
  }
  return sh;
}

function archiveOld() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    [moveOld_, purgeArchive_, purgeReminders_].forEach(f => { try { f(); } catch (e) { Logger.log(f.name + ': ' + e); } });
  } finally { lock.releaseLock(); }
}
function moveOld_() {
  {
    const sh = msgs_(), n = sh.getLastRow() - 1;
    if (n < 1) return;
    const vals = sh.getRange(2, 1, n, COLS).getValues(), cutoff = Date.now() - ARCHIVE_DAYS * 86400000;
    const move = [], keep = [];
    vals.forEach(r => {
      const t = new Date(r[1]).getTime();
      (t && t < cutoff && r[10] !== 'نعم') ? move.push(r) : keep.push(r);
    });
    if (!move.length) { Logger.log('لا توجد رسائل للأرشفة.'); return; }
    const ar = archSheet_(move.length);
    const stamp = Date.now();
    ar.getRange(ar.getLastRow() + 1, 1, move.length, ARCH_COLS).setValues(move.map(r => r.concat([stamp])));   // أولًا نكتب في الأرشيف، وبعدها فقط نحذف من الأصل
    sh.getRange(2, 1, n, COLS).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, COLS).setValues(keep);
    if (n > keep.length && sh.getMaxRows() - (keep.length + 1) > 1) sh.deleteRows(keep.length + 2, Math.min(n - keep.length, sh.getMaxRows() - (keep.length + 1)));
    Logger.log('تمت أرشفة ' + move.length + ' رسالة، وبقي ' + keep.length + ' في الصندوق.');
  }
}

// ===== مدة بقاء الأرشيف والتذكيرات المكتملة =====
// الرسائل تبقى في الأرشيف ARCHIVE_KEEP_DAYS يومًا من تاريخ نقلها إليه ثم تُحذف نهائيًا، وقبل ذلك بـ ARCHIVE_WARN_DAYS أيام
// يظهر للمستخدمة تنبيه مع خيار تحميل رسائلها. والتذكيرات المكتملة تُحذف بعد REM_KEEP_DAYS يومًا من إكمالها.
// يعمل كل ذلك ضمن الأرشفة اليومية archiveOld (الساعة 3 فجرًا) دون مشغّل إضافي.
function purgeArchive_() {
  const now = Date.now(), life = ARCHIVE_KEEP_DAYS * 86400000, warnMs = ARCHIVE_WARN_DAYS * 86400000, users = allUsers_(), first = users[0][0], warn = {};
  archIds_().forEach(id => {
    try {
      const sh = SpreadsheetApp.openById(id).getSheetByName('الأرشيف'); if (!sh) return;
      if (!sh.getRange(1, ARCH_COLS).getValue()) sh.getRange(1, ARCH_COLS).setValue('تاريخ الأرشفة');
      const n = sh.getLastRow() - 1; if (n < 1) return;
      const vals = sh.getRange(2, 1, n, ARCH_COLS).getValues(), keep = []; let stamped = false;
      vals.forEach(r => {
        let t = Number(r[COLS]);
        if (!t) { t = now; r[COLS] = t; stamped = true; }   // رسائل أُرشفت قبل هذه الميزة: تبدأ مدتها من الآن
        if (now - t >= life) return;                       // انتهت مدتها → تُحذف
        keep.push(r);
        const exp = t + life;
        if (exp - now <= warnMs) { const o = r[9] || first, w = warn[o] || (warn[o] = { n: 0, at: exp }); w.n++; if (exp < w.at) w.at = exp; }
      });
      if (keep.length === n && !stamped) return;
      sh.getRange(2, 1, n, ARCH_COLS).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, ARCH_COLS).setValues(keep);
      if (n > keep.length && sh.getMaxRows() - (keep.length + 1) > 1) sh.deleteRows(keep.length + 2, Math.min(n - keep.length, sh.getMaxRows() - (keep.length + 1)));
      Logger.log('حُذفت ' + (n - keep.length) + ' رسالة منتهية من الأرشيف.');
    } catch (e) { Logger.log('purgeArchive_ ' + id + ': ' + e); }
  });
  const P = PropertiesService.getScriptProperties();
  users.forEach(u => { const w = warn[u[0]]; if (w) P.setProperty('AW_' + u[0], JSON.stringify(w)); else P.deleteProperty('AW_' + u[0]); });
}
function purgeReminders_() {
  const sh = rems_(), n = sh.getLastRow() - 1; if (n < 1) return;
  const vals = sh.getRange(2, 1, n, 5).getValues(), cut = Date.now() - REM_KEEP_DAYS * 86400000, extra = sh.getRange(2, 13, n, 1).getValues(), del = [];
  vals.forEach((r, i) => { if (r[4] === 'مكتمل') { const t = Number(extra[i][0]) || Number(r[3]) || 0; if (t && t < cut) del.push(i + 2); } });
  del.reverse().forEach(row => sh.deleteRow(row));   // من الأسفل حتى لا تتغير أرقام الصفوف
  if (del.length) Logger.log('حُذف ' + del.length + ' تذكير مكتمل قديم.');
}
function archWarn_(uid) {
  try { const w = JSON.parse(PropertiesService.getScriptProperties().getProperty('AW_' + uid) || 'null'); return w && w.at > Date.now() ? w : null; } catch (e) { return null; }
}
// رسائل هذه المستخدمة فقط من كل ملفات الأرشيف (خام)، الأحدث أولًا
function archMine_(me) {
  const first = allUsers_()[0][0], mine = [];
  archIds_().forEach(id => {
    try {
      const ash = SpreadsheetApp.openById(id).getSheetByName('الأرشيف');
      if (ash && ash.getLastRow() >= 2) ash.getRange(2, 1, ash.getLastRow() - 1, ARCH_COLS).getValues().forEach(r => { if ((r[9] || first) === me[0]) mine.push(r); });
    } catch (e) {}
  });
  mine.sort((a, b) => (new Date(b[1]).getTime() || 0) - (new Date(a[1]).getTime() || 0));
  return mine;
}
// تنبيه «ستُحذف قريبًا» يُحسب من الرسائل الموجودة فعلًا في الأرشيف (فيزول فور حذفها أو تحميلها وحذفها)
function archWarnCalc_(rows) {
  const now = Date.now(), life = ARCHIVE_KEEP_DAYS * 86400000, warnMs = ARCHIVE_WARN_DAYS * 86400000; let n = 0, at = 0;
  rows.forEach(r => { const t = Number(r[COLS]); if (!t) return; const exp = t + life; if (exp - now <= warnMs) { n++; if (!at || exp < at) at = exp; } });
  return n ? { n: n, at: at } : null;
}
function setArchWarn_(uid, w) { const P = PropertiesService.getScriptProperties(); if (w) P.setProperty('AW_' + uid, JSON.stringify(w)); else P.deleteProperty('AW_' + uid); }
function archOut_(r) { const o = rowOut_(r), t = Number(r[COLS]); o.exp = t ? t + ARCHIVE_KEEP_DAYS * 86400000 : 0; return o; }

function installArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'archiveOld') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('archiveOld').timeBased().everyDays(1).atHour(3).create();
  Logger.log('تم ضبط الأرشفة اليومية (الساعة 3 فجرًا). يمكنك تشغيل archiveOld الآن للتجربة.');
}


function archIds_() { return (PropertiesService.getScriptProperties().getProperty('ARCH_ALL') || '').split(',').filter(Boolean); }
function listArchive_(me, d) {
  const limit = Math.min(Math.max(Number(d.limit) || 100, 1), 1000), mine = archMine_(me);
  return out_({ ok: true, v: VERSION, archive: true, archDays: ARCHIVE_DAYS, archKeep: ARCHIVE_KEEP_DAYS, archWarn: (function () { const w = archWarnCalc_(mine); setArchWarn_(me[0], w); return w; })(), rows: mine.slice(0, limit).map(archOut_), total: mine.length });
}
