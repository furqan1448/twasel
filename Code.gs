// ===== الإعدادات =====
const WA_VERIFY_TOKEN = 'furqan-wa-2026';   // أي كلمة تختارينها، وتكتبينها نفسها في Meta
const MSG = 'الرسائل', USR = 'المستخدمون';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, header) {
  let sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function msgs_() {
  const sh = sheet_(MSG, ['المعرف', 'التاريخ', 'الاسم', 'الجوال', 'البريد', 'الرسالة', 'الحالة', 'ملاحظة', 'المصدر', 'المالكة']);
  if (!sh.getRange(1, 9).getValue()) sh.getRange(1, 9).setValue('المصدر');
  if (!sh.getRange(1, 10).getValue()) sh.getRange(1, 10).setValue('المالكة');
  sh.getRange('D:D').setNumberFormat('@');
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

function doGet(e) {
  const p = e.parameter || {};
  if (p['hub.mode'] === 'subscribe') {
    return ContentService.createTextOutput(p['hub.verify_token'] === WA_VERIFY_TOKEN ? p['hub.challenge'] : 'forbidden');
  }
  if (p.action === 'profile') {
    const u = userById_(p.u);
    return u ? out_({ ok: true, user: { name: u[1], phone: u[2], email: u[3] } }) : out_({ ok: false, error: 'الصفحة غير موجودة' });
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
      msgs_().appendRow([Utilities.getUuid(), new Date(), d.name, d.phone || '', d.email || '', d.message, 'جديدة', '', 'النموذج', owner[0]]);
      notify_(owner[3], d.name, d.phone, d.email, d.message, 'النموذج');
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
      const recent = n > 0 ? sh.getRange(Math.max(2, n - 18), 1, Math.min(n, 20), 10).getValues() : [];
      if (recent.some(r => r[9] === me[0] && r[2] === title && r[5] === text)) return out_({ ok: true });   // منع التكرار
      sh.appendRow([Utilities.getUuid(), new Date(), title, '', '', text, 'جديدة', '', 'واتساب', me[0]]);
      return out_({ ok: true });
    }

    if (d.action === 'addManual') {
      if (!d.name) return out_({ ok: false, error: 'الاسم مطلوب.' });
      msgs_().appendRow([Utilities.getUuid(), new Date(), d.name, norm_(d.phone), '', d.message || '', 'جديدة', '', 'واتساب', me[0]]);
      return out_({ ok: true });
    }

    if (d.action === 'list') {
      const first = allUsers_()[0][0];
      const sh = msgs_();
      const rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues()
        .filter(r => (r[9] || first) === me[0])               // الرسائل القديمة بلا مالكة تعود لأول حساب
        .map(r => ({ id: r[0], date: r[1], name: r[2], phone: r[3], email: r[4], message: r[5], status: r[6], note: r[7], source: r[8] || 'النموذج' }))
        .reverse();
      return out_({ ok: true, rows });
    }

    if (d.action === 'update') {
      const first = allUsers_()[0][0];
      const sh = msgs_();
      if (sh.getLastRow() < 2) return out_({ ok: false, error: 'الرسالة غير موجودة.' });
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
      const i = vals.findIndex(r => r[0] === d.id && (r[9] || first) === me[0]);
      if (i < 0) return out_({ ok: false, error: 'الرسالة غير موجودة.' });
      if (d.status !== undefined) sh.getRange(i + 2, 7).setValue(d.status);
      if (d.note !== undefined) sh.getRange(i + 2, 8).setValue(d.note);
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
      sh.appendRow([m.id, ts, name, m.from, '', text, 'جديدة', '', 'واتساب', owner[0]]);
      notify_(owner[3], name, m.from, '', text, 'واتساب');
    });

    (v.message_echoes || []).forEach(m => {
      const to = norm_(m.to);
      if (!to || sh.getLastRow() < 2) return;
      sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues().forEach((r, i) => {
        if (norm_(r[3]) === to && r[8] === 'واتساب' && r[9] === owner[0] && r[6] !== 'تم الرد') sh.getRange(i + 2, 7).setValue('تم الرد');
      });
    });
  }));
}

function notify_(to, name, phone, email, message, source) {
  try {
    MailApp.sendEmail(to, 'رسالة جديدة (' + source + ') من ' + name,
      'الاسم: ' + name + '\nالجوال: ' + (phone || '-') + '\nالبريد: ' + (email || '-') + '\n\n' + message);
  } catch (err) {}
}


// ===== استيراد البريد تلقائيًا (Gmail) =====
// يقرأ بريد الحساب الذي نشر السكربت فقط، ويُنسب الرسائل لحساب مسجَّل بالبريد نفسه.
// شغّلي الدالة setupEmailTrigger مرة واحدة فقط لتفعيل الاستيراد كل 5 دقائق.
function setupEmailTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'importEmails').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importEmails').timeBased().everyMinutes(5).create();
  importEmails();
}

function importEmails() {
  const meEmail = Session.getEffectiveUser().getEmail().toLowerCase();
  const owner = allUsers_().find(r => String(r[3]).toLowerCase() === meEmail);
  if (!owner) return;                                   // سجّلي في الموقع بالبريد نفسه أولًا
  const sh = msgs_();
  const n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, 1, n, 10).getValues() : [];
  const index = {};
  vals.forEach((r, i) => { index[r[0]] = i; });

  GmailApp.search('newer_than:3d -category:promotions -category:social -category:updates -category:forums', 0, 50).forEach(t => {
    const msgs = t.getMessages(), last = msgs[msgs.length - 1];
    const fromMe = last.getFrom().toLowerCase().indexOf(meEmail) >= 0;
    const id = 'mail_' + t.getId();

    if (id in index) {
      const i = index[id], row = vals[i];
      if (fromMe && row[6] !== 'تم الرد') sh.getRange(i + 2, 7).setValue('تم الرد');
      else if (!fromMe && row[6] === 'تم الرد' && last.getDate() > new Date(row[1])) {   // وصل رد جديد بعد ردّك
        sh.getRange(i + 2, 7).setValue('جديدة'); sh.getRange(i + 2, 2).setValue(last.getDate());
      }
      return;
    }
    if (fromMe) return;
    const from = last.getFrom();
    const name = from.replace(/<.*>/, '').replace(/"/g, '').trim() || from;
    const mail = (from.match(/<(.+?)>/) || [])[1] || from;
    const text = (t.getFirstMessageSubject() + '\n' + last.getPlainBody()).slice(0, 1000);
    sh.appendRow([id, last.getDate(), name, '', mail, text, 'جديدة', '', 'البريد', owner[0]]);
  });
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
    const vals = n > 0 ? sh.getRange(2, 1, n, 10).getValues() : [];

    if (m.from_me) {                                                 // ردّك من الجوال: تُعلَّم المحادثة «تم الرد»
      vals.forEach((r, i) => {
        if (r[9] === me[0] && r[8] === 'واتساب' && norm_(r[3]) === phone && r[6] !== 'تم الرد') sh.getRange(i + 2, 7).setValue('تم الرد');
      });
      return;
    }
    if (vals.some(r => r[0] === m.id)) return;                       // منع التكرار
    const text = m.type === 'text' ? ((m.text && m.text.body) || '') : '[رسالة من نوع: ' + m.type + ']';
    const when = m.timestamp ? new Date(Number(m.timestamp) * 1000) : new Date();
    sh.appendRow([m.id, when, m.from_name || phone, phone, '', text, 'جديدة', '', 'واتساب', me[0]]);
  });
}
