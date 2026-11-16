// ===== الإعدادات =====
const WA_VERIFY_TOKEN = 'furqan-wa-2026';   // أي كلمة تختارينها، وتكتبينها نفسها في Meta
const MSG = 'الرسائل', USR = 'المستخدمون';
const COLS = 11;                              // عدد أعمدة جدول الرسائل
const DEL = 'محذوفة';                         // الرسالة المحذوفة: يُمسح محتواها كليًا ويبقى معرّفها فقط لمنع عودتها
const STATUSES_ = ['جديدة', 'تم الرد', 'لم يُرد عليها'];
// أي إيميل يصل إلى صندوق الوارد من أي مُرسِل (نستبعد فقط تصنيفي الترويج والاجتماعي)
const MAIL_QUERY = 'newer_than:3d in:inbox -category:promotions -category:social';

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
  const h = sh.getRange(1, 1, 1, COLS).getValues()[0];
  [[9, 'المصدر'], [10, 'المالكة'], [11, 'مؤرشفة']].forEach(x => { if (!h[x[0] - 1]) sh.getRange(1, x[0]).setValue(x[1]); });
  return sh;
}
function users_() {
  const sh = sheet_(USR, ['المعرف', 'الاسم', 'الجوال', 'البريد', 'التجزئة', 'الملح', 'الرمز', 'التاريخ']);
  sh.getRange('C:C').setNumberFormat('@');
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

function allUsers_() { const sh = users_(); return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(); }
function userByToken_(t) { return t ? allUsers_().find(r => r[6] === t) : null; }
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
    return u ? out_({ ok: true, user: { name: u[1], phone: u[2], email: u[3], mail: mailAddr_(u[0]) } }) : out_({ ok: false, error: 'الصفحة غير موجودة' });
  }
  return out_({ ok: true });
}

function doPost(e) {
  const raw = e.postData.contents;
  let d;
  if (raw.indexOf('wa§§') === 0) {                         // إشعار واتساب من تطبيق الجوال
    const p = raw.split('§§');
    d = { action: 'waNotify', token: p[1], title: p[2], text: p.slice(3).join('§§') };
  } else d = JSON.parse(raw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (e.parameter && e.parameter.wh) { handleWhapi_(e.parameter.wh, d); return out_({ ok: true }); }   // Whapi.Cloud
    if (d.object === 'whatsapp_business_account') { handleWhatsApp_(d, e.parameter && e.parameter.k); return out_({ ok: true }); }

    // ---- إنشاء حساب ----
    if (d.action === 'register') {
      const name = String(d.name || '').trim(), phone = norm_(d.phone), email = String(d.email || '').trim().toLowerCase(), pw = String(d.password || '');
      if (!name || phone.length < 11 || email.indexOf('@') < 1 || pw.length < 6)
        return out_({ ok: false, error: 'يرجى إدخال الاسم ورقم جوال صحيح وبريد صحيح وكلمة مرور لا تقل عن 6 أحرف.' });
      const dup = allUsers_().some(r => norm_(r[2]) === phone || String(r[3]).toLowerCase() === email);
      if (dup) return out_({ ok: false, error: 'رقم الجوال أو البريد مسجّل مسبقًا، يرجى تسجيل الدخول.' });
      const id = Utilities.getUuid().replace(/-/g, '').slice(0, 10), salt = Utilities.getUuid(), token = Utilities.getUuid();
      users_().appendRow([id, name, phone, email, hash_(pw, salt), salt, token, new Date()]);
      return out_({ ok: true, token, user: { id, name, phone, email } });
    }

    // ---- تسجيل الدخول ----
    if (d.action === 'login') {
      const idf = String(d.identifier || '').trim().toLowerCase(), n = norm_(idf);
      const u = allUsers_().find(r => String(r[3]).toLowerCase() === idf || (n && norm_(r[2]) === n));
      if (!u || hash_(String(d.password || ''), u[5]) !== u[4]) return out_({ ok: false, error: 'بيانات الدخول غير صحيحة.' });
      return out_({ ok: true, token: u[6], user: pub_(u) });
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
    const me = userByToken_(d.token);
    if (!me) return out_({ ok: false, error: 'انتهت الجلسة، يرجى تسجيل الدخول.', auth: false });

    if (d.action === 'waNotify') {
      const title = String(d.title || '').trim(), text = String(d.text || '').trim();
      if (!title || !text) return out_({ ok: true });
      if (/^(whatsapp|واتساب)/i.test(title) || /\d+\s*(new messages|رسائل جديدة|رسالة جديدة)/i.test(text)) return out_({ ok: true });   // إشعارات التجميع
      const sh = msgs_(), n = sh.getLastRow() - 1;
      const recent = n > 0 ? sh.getRange(Math.max(2, n - 18), 1, Math.min(n, 20), COLS).getValues() : [];
      if (recent.some(r => r[9] === me[0] && r[6] !== DEL && dec_(r[2]) === title && dec_(r[5]) === text)) return out_({ ok: true });   // منع التكرار
      sh.appendRow([Utilities.getUuid(), new Date(), enc_(title), '', '', enc_(text), 'جديدة', '', 'واتساب', me[0], '']);
      return out_({ ok: true });
    }

    if (d.action === 'addManual') {
      if (!d.name) return out_({ ok: false, error: 'الاسم مطلوب.' });
      msgs_().appendRow([Utilities.getUuid(), new Date(), enc_(String(d.name).slice(0, 200)), enc_(norm_(d.phone)), '', enc_(String(d.message || '').slice(0, 5000)), 'جديدة', '', 'واتساب', me[0], '']);
      return out_({ ok: true });
    }

    if (d.action === 'list') {
      const first = allUsers_()[0][0];
      const sh = msgs_(), wantArchived = !!d.archived;
      let archivedCount = 0;
      const rows = [];
      if (sh.getLastRow() >= 2) {
        sh.getRange(2, 1, sh.getLastRow() - 1, COLS).getValues().forEach(r => {
          if ((r[9] || first) !== me[0] || r[6] === DEL) return;       // الرسائل القديمة بلا مالكة تعود لأول حساب
          const isArch = r[10] === 'نعم';
          if (isArch) archivedCount++;
          if (isArch === wantArchived) rows.push(rowOut_(r));            // فك التشفير للمطلوب عرضه فقط
        });
      }
      rows.reverse();
      return out_({ ok: true, rows, archivedCount });
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
      // مسح كامل للمحتوى؛ يبقى المعرّف فقط حتى لا يعيد استيراد البريد/الواتساب الرسالة نفسها
      sh.getRange(f.row, 2, 1, COLS - 1).setValues([[new Date(), '', '', '', '', DEL, '', f.vals[8], f.vals[9], '']]);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'طلب غير معروف.' });
  } finally { lock.releaseLock(); }
}

function handleWhatsApp_(d, key) {
  const owner = key ? userByToken_(key) : null;          // الرابط الخاص بكل حساب يحمل مفتاحه (?k=...)
  if (!owner) return;
  const sh = msgs_();
  (d.entry || []).forEach(entry => (entry.changes || []).forEach(ch => {
    const v = ch.value || {};

    (v.messages || []).forEach(m => {
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
      if (ids.indexOf(m.id) >= 0) return;
      const c = (v.contacts || []).find(x => x.wa_id === m.from);
      const name = (c && c.profile && c.profile.name) || m.from;
      let text = m.type === 'text' ? m.text.body : m.type === 'button' ? m.button.text : '[رسالة من نوع: ' + m.type + ']';
      const ts = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
      sh.appendRow([m.id, ts, enc_(name), enc_(m.from), '', enc_(text), 'جديدة', '', 'واتساب', owner[0], '']);
      notify_(owner[3], name, m.from, '', text, 'واتساب');
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

function notify_(to, name, phone, email, message, source) {
  try {
    MailApp.sendEmail(to, 'رسالة جديدة (' + source + ') من ' + name,
      'الاسم: ' + name + '\nالجوال: ' + (phone || '-') + '\nالبريد: ' + (email || '-') + '\n\n' + message);
  } catch (err) { Logger.log('تعذّر إرسال إشعار البريد: ' + err); }
}


// ===== استيراد البريد تلقائيًا (Gmail) =====
// لكل مستخدمة عنوان استقبال خاص بها: (بريد الحساب الذي نشر السكربت) مع +المعرّف، مثل: inbox+a1b2c3d4e5@gmail.com
// أي إيميل يُرسل لهذا العنوان (أو يُحوَّل إليه من بريدها) يظهر في صندوقها هي فقط.
// الرسائل التي لا تخص أي مستخدمة لا تُستورد أبدًا.
// شغّلي الدالة setupEmailTrigger مرة واحدة فقط لتفعيل الاستيراد كل 5 دقائق.
function setupEmailTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'importEmails').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importEmails').timeBased().everyMinutes(5).create();
  importEmails();
}

function importEmails() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { importEmails_(); } finally { lock.releaseLock(); }
}

function mailText_(t, last) {
  return (t.getFirstMessageSubject() + '\n' + last.getPlainBody()).slice(0, 1000);
}

function baseMail_() {
  return (Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '').toLowerCase();
}
// عنوان الاستقبال الخاص بالمستخدمة
function mailAddr_(id) {
  const me = baseMail_(), at = me.lastIndexOf('@');
  return at < 0 ? '' : me.slice(0, at) + '+' + id + '@' + me.slice(at + 1);
}

// يحدد صاحبة الرسالة: عنوان الاستقبال (+المعرّف) في To/Cc، أو في ترويسة Delivered-To (للتحويل والنسخة المخفية)
function routeMail_(last, users, meEmail) {
  const hay = (last.getTo() + ' ' + last.getCc()).toLowerCase();
  let m = hay.match(/\+([0-9a-f]{10})@/);
  if (!m) {
    try {
      const raw = String(last.getRawContent()), end = raw.search(/\r?\n\r?\n/), head = raw.slice(0, end > 0 ? end : 20000).toLowerCase();
      m = head.match(/^(?:delivered-to|x-original-to|x-forwarded-to):[^\n]*\+([0-9a-f]{10})@/m);
    } catch (err) {}
  }
  if (m) return users.find(r => r[0] === m[1]) || null;
  if (hay.indexOf(meEmail) >= 0) return users.find(r => String(r[3]).toLowerCase() === meEmail) || null;   // أُرسلت مباشرة لبريد الحساب نفسه
  return null;
}

function importEmails_() {
  const users = allUsers_();
  if (!users.length) { Logger.log('لا يوجد أي حساب مسجّل في الموقع بعد.'); return; }
  const meEmail = baseMail_();
  if (!meEmail) { Logger.log('تعذّر معرفة بريد الحساب الذي نشر السكربت.'); return; }
  const mine = [meEmail].concat(GmailApp.getAliases().map(a => a.toLowerCase()));
  const cache = CacheService.getScriptCache();

  const sh = msgs_();
  purge_(sh);
  const n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, 1, n, COLS).getValues() : [];
  const index = {};
  vals.forEach((r, i) => { index[r[0]] = i; });

  const threads = GmailApp.search(MAIL_QUERY, 0, 50);
  let added = 0, skipped = 0;
  threads.forEach(t => {
    const msgs = t.getMessages(), last = msgs[msgs.length - 1];
    const lastFrom = last.getFrom().toLowerCase();
    const fromMe = mine.some(a => lastFrom.indexOf(a) >= 0);
    const id = 'mail_' + t.getId();

    if (id in index) {
      const i = index[id], row = vals[i];
      if (row[6] === DEL) return;                                    // محذوفة: لا تعود
      if (fromMe && row[6] !== 'تم الرد') sh.getRange(i + 2, 7).setValue('تم الرد');
      else if (!fromMe && (row[6] === 'تم الرد' || row[10] === 'نعم') && last.getDate() > new Date(row[1])) {   // وصل رد جديد بعد ردّك أو بعد الأرشفة
        sh.getRange(i + 2, 7).setValue('جديدة');
        sh.getRange(i + 2, 2).setValue(last.getDate());
        sh.getRange(i + 2, 6).setValue(enc_(mailText_(t, last)));
        sh.getRange(i + 2, 11).setValue('');
      }
      return;
    }
    if (fromMe) return;
    if (cache.get('skip_' + t.getId())) return;                      // سبق أن تبيّن أنها لا تخص أي مستخدمة
    const owner = routeMail_(last, users, meEmail);
    if (!owner) { cache.put('skip_' + t.getId(), '1', 21600); skipped++; return; }
    const from = last.getFrom();
    const name = from.replace(/<.*>/, '').replace(/"/g, '').trim() || from;
    const mail = (from.match(/<(.+?)>/) || [])[1] || from;
    sh.appendRow([id, last.getDate(), enc_(name), '', enc_(mail), enc_(mailText_(t, last)), 'جديدة', '', 'البريد', owner[0], '']);
    added++;
  });
  Logger.log('البريد: ' + meEmail + ' | سلاسل مطابقة: ' + threads.length + ' | أُضيف: ' + added + ' | لا تخص أحدًا: ' + skipped);
}

// تنظيف: حذف صفوف الرسائل المحذوفة نهائيًا بعد 7 أيام (لا تحمل أي محتوى أصلًا)
function purge_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, COLS).getValues(), cutoff = Date.now() - 7 * 86400000;
  for (let i = n - 1; i >= 0; i--) {
    if (vals[i][6] === DEL && new Date(vals[i][1]).getTime() < cutoff) sh.deleteRow(i + 2);
  }
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
  const me = userByToken_(token);
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
    if (vals.some(r => r[0] === m.id)) return;                       // منع التكرار
    const text = m.type === 'text' ? ((m.text && m.text.body) || '') : '[رسالة من نوع: ' + m.type + ']';
    const when = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
    sh.appendRow([m.id, when, enc_(m.from_name || phone), enc_(phone), '', enc_(text), 'جديدة', '', 'واتساب', me[0], '']);
  });
}


// ===== فحص الاستيراد: شغّليها من المحرر ثم افتحي «سجل التنفيذ» (Execution log) =====
// تبيّن لكل رسالة من آخر 3 أيام: لمن ستُنسب؟ وهل ستُستورد؟ وإن لا فما السبب.
function diagnoseEmail() {
  const users = allUsers_(), meEmail = baseMail_();
  Logger.log('حساب Google الذي نشر السكربت (صندوق البريد الذي يُقرأ): ' + meEmail);
  users.forEach(r => Logger.log('مستخدمة: ' + r[1] + ' ← عنوان استقبالها: ' + mailAddr_(r[0])));
  const mine = [meEmail].concat(GmailApp.getAliases().map(a => a.toLowerCase()));
  const sh = msgs_(), n = sh.getLastRow() - 1;
  const state = {};
  (n > 0 ? sh.getRange(2, 1, n, COLS).getValues() : []).forEach(r => { state[r[0]] = r[6]; });
  const passed = {};
  GmailApp.search(MAIL_QUERY, 0, 50).forEach(t => { passed[t.getId()] = 1; });
  const all = GmailApp.search('newer_than:3d', 0, 20);
  Logger.log('رسائل آخر 3 أيام: ' + all.length + ' | تمرّ من الفلتر: ' + Object.keys(passed).length);
  all.forEach(t => {
    const ms = t.getMessages(), last = ms[ms.length - 1], from = last.getFrom().toLowerCase();
    let why;
    if (!passed[t.getId()]) why = 'مستبعدة بالفلتر (تصنيف ترويج/اجتماعي، أو ليست في الوارد)';
    else if (mine.some(a => from.indexOf(a) >= 0)) why = 'آخر رسالة فيها منك أنتِ فلا تُستورد';
    else if (state['mail_' + t.getId()] === DEL) why = 'سبق أن حُذفت من الموقع';
    else if (('mail_' + t.getId()) in state) why = 'مستوردة سابقًا (الحالة: ' + state['mail_' + t.getId()] + ')';
    else {
      const o = routeMail_(last, users, meEmail);
      why = o ? 'ستُنسب إلى: ' + o[1] + ' وتُستورد' : 'لا تخص أي مستخدمة (لم تُرسل إلى عنوان استقبال أحد) فلن تُستورد';
    }
    Logger.log('• ' + String(t.getFirstMessageSubject()).slice(0, 40) + ' | من: ' + last.getFrom().slice(0, 40) + ' → ' + why);
  });
}
