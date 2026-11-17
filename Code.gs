/** @OnlyCurrentDoc */
// ===== الإعدادات =====
const WA_VERIFY_TOKEN = 'furqan-wa-2026';   // أي كلمة تختارينها، وتكتبينها نفسها في Meta
const VERSION = '2026-11-17-b';               // رقم الإصدار: يظهر عند فتح رابط النشر في المتصفح، ليُعرف هل النشر يعمل بآخر كود
const MSG = 'الرسائل', USR = 'المستخدمون';
const COLS = 11;                              // عدد أعمدة جدول الرسائل
const DEL = 'محذوفة';                         // قديم: صفوف قديمة تُحذف تلقائيًا من الشيت ولا يُكتب هذا النص بعد الآن
const STATUSES_ = ['جديدة', 'تم الرد', 'لم يُرد عليها'];
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
function users_() {
  const sh = sheet_(USR, ['المعرف', 'الاسم', 'الجوال', 'البريد', 'التجزئة', 'الملح', 'الرمز', 'التاريخ']);
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('PH_FMT')) { sh.getRange('C:C').setNumberFormat('@'); P.setProperty('PH_FMT', '1'); }   // مرة واحدة فقط (كانت تُنفَّذ مع كل طلب فتبطّئ الموقع)
  return sh;
}

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
    status: r[6], note: dec_(r[7]), source: r[8] || 'النموذج', archived: r[10] === 'نعم'
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
  if (p.action === 'profile') {
    const u = userById_(p.u);
    return u ? out_({ ok: true, user: { name: u[1], phone: u[2], email: u[3] } }) : out_({ ok: false, error: 'الصفحة غير موجودة' });
  }
  return out_({ ok: true, v: VERSION });
}

function doPost(e) {
  const raw = e.postData.contents;
  let d;
  if (raw.indexOf('wa§§') === 0) {                         // إشعار واتساب من تطبيق الجوال
    const p = raw.split('§§');
    d = { action: 'waNotify', token: p[1], title: p[2], text: p.slice(3).join('§§') };
  } else d = JSON.parse(raw);
  const lock = LockService.getScriptLock();
  const needLock = d.action !== 'list' && d.action !== 'keys';   // القراءة لا تنتظر بعضها، فيسرع الموقع
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
      notify_(owner[3], name, phone, email, message, 'النموذج');
      return out_({ ok: true });
    }

    // ---- ما يلزم له تسجيل دخول ----
    const me = (d.action === 'waNotify' || d.action === 'mailIn') ? userByKey_(d.token) : userBySession_(d.token);
    if (!me) return out_({ ok: false, error: 'انتهت الجلسة، يرجى تسجيل الدخول.', auth: false });

    if (d.action === 'keys') return out_({ ok: true, key: me[6] });
    if (d.action === 'logout') { dropSession_(d.token); return out_({ ok: true }); }

    if (d.action === 'changePassword') {
      const np = String(d.password || '');
      if (hash_(String(d.current || ''), me[5]) !== me[4]) return out_({ ok: false, error: 'كلمة المرور الحالية غير صحيحة.' });
      if (np.length < 8) return out_({ ok: false, error: 'كلمة المرور الجديدة لا تقل عن 8 أحرف.' });
      setPw_(me, np); endSessions_(me[0]);
      sendMail_(me[3], 'تم تغيير كلمة المرور', 'تم تغيير كلمة مرور حسابك في صندوق الرسائل. إن لم تكوني أنتِ، فاستعيديها فورًا من خيار «نسيت كلمة المرور».');
      return out_({ ok: true, token: newSession_(me[0]) });
    }

    if (d.action === 'waNotify') {
      const title = String(d.title || '').trim(), text = String(d.text || '').trim();
      if (!title || !text) return out_({ ok: true });
      if (/^(whatsapp|واتساب)/i.test(title) || /\d+\s*(new messages|رسائل جديدة|رسالة جديدة)/i.test(text)) return out_({ ok: true });   // إشعارات التجميع
      if (isGoneC_(cid_(me[0], title, text))) return out_({ ok: true });   // رسالة حُذفت قبل قليل: لا تعود
      const sh = msgs_(), n = sh.getLastRow() - 1;
      const recent = n > 0 ? sh.getRange(Math.max(2, n - 18), 1, Math.min(n, 20), COLS).getValues() : [];
      if (recent.some(r => r[9] === me[0] && r[6] !== DEL && dec_(r[2]) === title && dec_(r[5]) === text)) return out_({ ok: true });   // منع التكرار
      sh.appendRow([Utilities.getUuid(), new Date(), enc_(title), '', '', enc_(text), 'جديدة', '', 'واتساب', me[0], '']);
      return out_({ ok: true });
    }

    if (d.action === 'mailIn') { mailIn_(me, d.items); return out_({ ok: true }); }

    if (d.action === 'addManual') {
      if (!d.name) return out_({ ok: false, error: 'الاسم مطلوب.' });
      msgs_().appendRow([Utilities.getUuid(), new Date(), enc_(String(d.name).slice(0, 200)), enc_(norm_(d.phone)), '', enc_(String(d.message || '').slice(0, 5000)), 'جديدة', '', 'واتساب', me[0], '']);
      return out_({ ok: true });
    }

    if (d.action === 'list') {
      // يُرجع أحدث رسائل العرض المطلوب فقط (الافتراضي 100) مع أعداد الحالات كاملة، ويفك تشفير المعروض فقط
      const first = allUsers_()[0][0];
      const sh = msgs_(), wantArchived = !!d.archived, limit = Math.min(Math.max(Number(d.limit) || 100, 1), 1000);
      cleanDeleted_(sh);
      let archivedCount = 0;
      const mine = [], counts = {};
      if (sh.getLastRow() >= 2) {
        sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach(r => {
          if ((r[9] || first) !== me[0] || r[6] === DEL) return;       // الرسائل القديمة بلا مالكة تعود لأول حساب
          const isArch = r[10] === 'نعم';
          if (isArch) archivedCount++; else counts[r[6]] = (counts[r[6]] || 0) + 1;
          if (isArch === wantArchived) mine.push(r);
        });
      }
      const ord = new Map(mine.map((r, i) => [r, i]));
      mine.sort((a, b) => (new Date(b[1]).getTime() || 0) - (new Date(a[1]).getTime() || 0) || ord.get(b) - ord.get(a));   // الأحدث أولًا
      return out_({ ok: true, rows: mine.slice(0, limit).map(rowOut_), archivedCount, counts, total: mine.length });
    }

    if (d.action === 'update') {
      const sh = msgs_(), f = findRow_(sh, d.id, me);
      if (!f) return out_({ ok: false, error: 'الرسالة غير موجودة.' });
      if (d.status !== undefined && STATUSES_.indexOf(d.status) < 0) return out_({ ok: false, error: 'حالة غير صحيحة.' });
      if (d.status !== undefined) sh.getRange(f.row, 7).setValue(d.status);
      if (d.note !== undefined) sh.getRange(f.row, 8).setValue(enc_(String(d.note).slice(0, 2000)));
      if (d.archived !== undefined) sh.getRange(f.row, 11).setValue(d.archived ? 'نعم' : '');
      return out_({ ok: true });
    }

    if (d.action === 'delete') {
      const sh = msgs_(), f = findRow_(sh, d.id, me);
      if (!f) return out_({ ok: false, error: 'الرسالة غير موجودة.' });
      buryRows_([f.vals], me[0]); sh.deleteRow(f.row);          // حذف الصف كاملًا من الشيت، بلا أي أثر نصي
      return out_({ ok: true });
    }
    if (d.action === 'deleteReplied') {
      const sh = msgs_();
      if (sh.getLastRow() < 2) return out_({ ok: true, deleted: 0 });
      const first = allUsers_()[0][0], ids = [], rows = [];
      sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach((r, i) => {
        if ((r[9] || first) !== me[0] || r[6] !== 'تم الرد' || r[10] === 'نعم') return;
        ids.push(r); rows.push(i + 2);
      });
      if (ids.length) { buryRows_(ids, me[0]); rows.reverse().forEach(x => sh.deleteRow(x)); }
      return out_({ ok: true, deleted: ids.length });
    }
    return out_({ ok: false, error: 'طلب غير معروف.' });
  } finally { if (needLock) lock.releaseLock(); }
}

// بريد المستخدمة: يصل من سكربت تشغّله هي في حسابها (Gmail) فلا يُقرأ بريد أي حساب آخر
function mailIn_(me, items) {
  const sh = msgs_(), n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, 1, n, COLS).getValues() : [];
  const index = {};
  vals.forEach((r, i) => { index[r[0]] = i; });
  (Array.isArray(items) ? items : []).slice(0, 50).forEach(it => {
    if (!it || !it.id) return;
    const id = 'mail_' + me[0] + '_' + String(it.id).slice(0, 40);
    if (isGone_(id)) return;
    const when = it.date ? new Date(Number(it.date)) : new Date();
    const text = (String(it.subject || '') + '\n' + String(it.body || '')).slice(0, 1000);
    if (id in index) {
      const i = index[id], row = vals[i];
      if (row[9] !== me[0] || row[6] === DEL) return;
      if (it.fromMe) { if (row[6] !== 'تم الرد') sh.getRange(i + 2, 7).setValue('تم الرد'); }
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
  (d.entry || []).forEach(entry => (entry.changes || []).forEach(ch => {
    const v = ch.value || {};

    (v.messages || []).forEach(m => {
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
      if (ids.indexOf(m.id) >= 0 || isGone_(m.id)) return;
      const c = (v.contacts || []).find(x => x.wa_id === m.from);
      const name = (c && c.profile && c.profile.name) || m.from;
      let text = m.type === 'text' ? m.text.body : m.type === 'button' ? m.button.text : '[رسالة من نوع: ' + m.type + ']';
      const ts = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
      sh.appendRow([m.id, ts, enc_(name), enc_(m.from), '', enc_(text), 'جديدة', '', 'واتساب', owner[0], '']);
    });

    (v.message_echoes || []).forEach(m => {
      const to = norm_(m.to);
      if (!to || sh.getLastRow() < 2) return;
      sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach((r, i) => {
        if (r[8] === 'واتساب' && r[9] === owner[0] && r[6] !== 'تم الرد' && r[6] !== DEL && norm_(dec_(r[3])) === to) sh.getRange(i + 2, 7).setValue('تم الرد');
      });
    });
  }));
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
// عند الحذف يُمسح الصف كاملًا من الشيت. ولكي لا يعيد استيراد البريد (آخر 3 أيام) الرسالة نفسها،
// نحفظ بصمة مقتطعة (Hash) لمعرّفها فقط في خصائص السكربت (ليس في الشيت)، وتُنسى بعد 5 أيام.
function tb_(id) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'tb:' + id)).slice(0, 8); }
function tombs_(k) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(k || 'TOMB') || '[]'); } catch (e) { return []; }
}
function isGone_(id) { const h = tb_(id); return tombs_().some(x => x.slice(0, 8) === h); }
function bury_(ids, k, ttlDays, cap) {
  k = k || 'TOMB';
  const now = Math.floor(Date.now() / 60000), ttl = (ttlDays || 5) * 1440;
  const keep = tombs_(k).filter(x => now - parseInt(x.slice(8), 36) < ttl);
  ids.forEach(id => keep.push(tb_(id) + now.toString(36)));
  PropertiesService.getScriptProperties().setProperty(k, JSON.stringify(keep.slice(-(cap || 450))));
}
// إشعارات الواتساب القادمة من تطبيق الجوال ليس لها معرّف ثابت، فنمنع عودتها بحسب (المالكة + المرسل + النص) لمدة 12 ساعة بعد حذفها.
function cid_(owner, name, text) { return owner + '|' + name + '|' + text; }
function isGoneC_(c) { const h = tb_(c), now = Math.floor(Date.now() / 60000); return tombs_('TOMBC').some(x => x.slice(0, 8) === h && now - parseInt(x.slice(8), 36) < 720); }
function buryRows_(rows, owner) {   // rows: صفوف الشيت المحذوفة (مصفوفات)
  bury_(rows.map(r => r[0]));
  const c = rows.filter(r => r[8] === 'واتساب').map(r => cid_(owner, dec_(r[2]), dec_(r[5])));
  if (c.length) bury_(c, 'TOMBC', 0.5, 200);
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
      vals.forEach((r, i) => {
        if (r[9] === me[0] && r[8] === 'واتساب' && r[6] !== 'تم الرد' && r[6] !== DEL && norm_(dec_(r[3])) === phone) sh.getRange(i + 2, 7).setValue('تم الرد');
      });
      return;
    }
    if (vals.some(r => r[0] === m.id) || isGone_(m.id)) return;                       // منع التكرار
    const text = m.type === 'text' ? ((m.text && m.text.body) || '') : '[رسالة من نوع: ' + m.type + ']';
    const when = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
    sh.appendRow([m.id, when, enc_(m.from_name || phone), enc_(phone), '', enc_(text), 'جديدة', '', 'واتساب', me[0], '']);
  });
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
