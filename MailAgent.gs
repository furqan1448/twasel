// ===== سكربت ربط البريد (مشروع واحد تنشرينه أنتِ مرة واحدة) =====
// النشر: نشر ← عملية نشر جديدة ← تطبيق ويب
//   تنفيذ باسم: «المستخدم الذي يصل إلى تطبيق الويب»   ·   من يمكنه الوصول: «أي شخص لديه حساب Google»
// ثم انسخي رابط النشر (ينتهي بـ /exec) وضعيه في index.html مكان MAIL_LINK.
// أي تعديل لاحق على هذا الكود يسري على جميع المربوطات تلقائيًا دون أن تفعل المستخدمات شيئًا.

function cfg_() {
  const P = PROP_();
  return { key: P.getProperty('KEY'), api: P.getProperty('API'), site: P.getProperty('SITE') || '' };
}

const PROP_ = () => PropertiesService.getUserProperties();

function post_(api, payload) {
  try {
    const r = UrlFetchApp.fetch(api, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: payload, muteHttpExceptions: true });
    return JSON.parse(r.getContentText());
  } catch (e) { return null; }
}

// إن تغيّر رابط النشر يقرؤه من الموقع تلقائيًا، فلا تحتاجين تعديل شيء
function refresh_(c) {
  try {
    const html = UrlFetchApp.fetch(c.site + (c.site.indexOf('?') < 0 ? '?' : '&') + 't=' + Date.now(), { muteHttpExceptions: true }).getContentText();
    const m = html.match(/const API = '(https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec)'/);
    if (m) { PROP_().setProperty('API', m[1]); return m[1]; }
  } catch (e) {}
  return '';
}

// خفيف: يفحص الرسائل الجديدة فقط، ويرسلها للموقع، فيصلك البريد خلال دقيقة تقريبًا
function sync() {
  const c = cfg_(), P = PROP_();
  if (!c.key || !c.api) return { ok: false };
  const last = Number(P.getProperty('LAST') || 0), now = Date.now();
  const me = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const items = GmailApp.search('newer_than:3d (in:inbox OR in:sent) -category:promotions -category:social', 0, 30)
    .filter(t => !last || t.getLastMessageDate().getTime() > last - 600000)
    .map(t => {
      const ms = t.getMessages(), m = ms[ms.length - 1], from = m.getFrom();
      return { id: t.getId(), date: m.getDate().getTime(), from: from, subject: t.getFirstMessageSubject(),
               body: m.getPlainBody().slice(0, 1000), fromMe: !!me && from.toLowerCase().indexOf(me) >= 0 };
    });
  if (!items.length && now - Number(P.getProperty('PING') || 0) < 1800000) return { ok: true };
  const payload = JSON.stringify({ action: 'mailIn', token: c.key, items: items });
  let r = post_(c.api, payload);
  if (!r && c.site) { const a = refresh_(c); if (a) r = post_(a, payload); }
  if (r && r.ok) { P.setProperty('LAST', String(now)); P.setProperty('PING', String(now)); }
  return r || { ok: false };
}

function page_(title, text) {
  return HtmlService.createHtmlOutput('<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:420px;margin:12vh auto;padding:24px;text-align:center;line-height:2">' +
    '<h2 style="color:#8C1A2C">' + title + '</h2><p>' + text + '</p></div>').setTitle(title);
}

function doGet(e) {
  const p = e.parameter || {}, P = PROP_();
  if (p.off) {
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t)); P.deleteAllProperties();
    return page_('تم إيقاف الربط', 'لن تُرسَل رسائل بريدك إلى الموقع بعد الآن.');
  }
  if (!p.k || !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(p.api || ''))
    return page_('رابط غير صحيح', 'افتحي رابط الربط من زر «ربط البريد» داخل الموقع.');
  P.setProperties({ KEY: p.k, API: p.api, SITE: /^https:\/\//.test(p.site || '') ? p.site : '', LAST: '0', PING: '0' });
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const r = sync();
  if (!r || !r.ok) {
    P.deleteAllProperties();
    return page_('تعذّر الربط', r && r.auth === false ? 'مفتاح الربط غير صحيح، افتحي رابط الربط من الموقع من جديد.' : 'تعذّر الاتصال بالموقع، حاولي مرة أخرى بعد قليل.');
  }
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(1).create();
  return page_('تم ربط بريدك ✓', 'ستصل رسائلك إلى الموقع خلال دقيقة تقريبًا. يمكنك إغلاق هذه الصفحة.');
}
