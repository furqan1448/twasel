// ===== سكربت ربط البريد (مشروع واحد تنشرينه أنتِ مرة واحدة) =====
// النشر: نشر ← عملية نشر جديدة ← تطبيق ويب
//   تنفيذ باسم: «المستخدم الذي يصل إلى تطبيق الويب»   ·   من يمكنه الوصول: «أي شخص لديه حساب Google»
// ثم انسخي رابط النشر (ينتهي بـ /exec) وضعيه في index.html مكان MAIL_LINK.

function cfg_() {
  const P = PROP_();
  return { key: P.getProperty('KEY'), api: P.getProperty('API'), site: P.getProperty('SITE') || '' };
}

const PROP_ = () => PropertiesService.getUserProperties();

function post_(api, payload) {
  try {
    const r = UrlFetchApp.fetch(api, {
      method: 'post',
      contentType: 'text/plain;charset=utf-8',
      payload: payload,
      followRedirects: true,
      muteHttpExceptions: true
    });
    const code = r.getResponseCode();
    const txt = r.getContentText();
    try {
      const j = JSON.parse(txt);
      j._http = code;
      return j;
    } catch (je) {
      return { ok: false, _http: code, _raw: txt.slice(0, 180) };
    }
  } catch (e) {
    return { ok: false, _err: String(e) };
  }
}

function getFallback_(api, key, items) {
  try {
    const b64 = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(items)).getBytes());
    const url = api + '?action=mailIn&k=' + encodeURIComponent(key) + '&items=' + encodeURIComponent(b64);
    if (url.length > 7500) return null;
    const r = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true });
    return JSON.parse(r.getContentText());
  } catch (e) { return null; }
}

function refresh_(c) {
  try {
    const html = UrlFetchApp.fetch(c.site + (c.site.indexOf('?') < 0 ? '?' : '&') + 't=' + Date.now(), { muteHttpExceptions: true }).getContentText();
    const m = html.match(/const API = '(https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec)'/);
    if (m) { PROP_().setProperty('API', m[1]); return m[1]; }
  } catch (e) {}
  return '';
}

function sync(forceAll) {
  const c = cfg_(), P = PROP_();
  if (!c.key || !c.api) return { ok: false, reason: 'missing_cfg' };
  const last = forceAll ? 0 : Number(P.getProperty('LAST') || 0), now = Date.now();
  const me = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const threads = GmailApp.search('newer_than:7d (in:inbox OR in:sent)', 0, 25);
  const items = threads
    .filter(t => !last || t.getLastMessageDate().getTime() > last - 600000)
    .map(t => {
      const ms = t.getMessages(), m = ms[ms.length - 1], from = m.getFrom();
      return {
        id: t.getId(),
        date: m.getDate().getTime(),
        from: from,
        subject: t.getFirstMessageSubject(),
        body: m.getPlainBody().slice(0, 800),
        fromMe: !!me && from.toLowerCase().indexOf(me) >= 0
      };
    });
  if (!forceAll && !items.length && now - Number(P.getProperty('PING') || 0) < 1800000) return { ok: true, count: 0 };
  const payload = JSON.stringify({ action: 'mailIn', token: c.key, items: items });
  let r = post_(c.api, payload);
  if ((!r || !r.ok) && r && r.auth !== false) {
    const g = getFallback_(c.api, c.key, items);
    if (g && g.ok) r = g;
  }
  if ((!r || !r.ok) && c.site) {
    const a = refresh_(c);
    if (a) r = post_(a, payload);
  }
  if (r && r.ok) {
    P.setProperty('LAST', String(now));
    P.setProperty('PING', String(now));
    r.synced = items.length;
    r.email = me;
  }
  return r || { ok: false };
}

function page_(title, text) {
  return HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:440px;margin:10vh auto;padding:24px;text-align:center;line-height:2;border:1px solid #ebe1e2;border-radius:16px">' +
    '<h2 style="color:#8C1A2C">' + title + '</h2><div>' + text + '</div></div>'
  ).setTitle(title);
}

function doGet(e) {
  const p = e.parameter || {}, P = PROP_();
  if (p.off) {
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
    P.deleteAllProperties();
    return page_('تم إيقاف الربط', 'لن تُرسَل رسائل بريدك إلى الموقع بعد الآن.');
  }
  if (!p.k || !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(p.api || '')) {
    return page_('رابط غير صحيح', 'افتحي رابط الربط من زر «ربط البريد» داخل الموقع.');
  }
  P.setProperties({ KEY: p.k, API: p.api, SITE: /^https:\/\//.test(p.site || '') ? p.site : '', LAST: '0', PING: '0' });
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const r = sync(true);
  if (!r || !r.ok) {
    P.deleteAllProperties();
    const detail = r && r.auth === false
      ? 'مفتاح الربط غير مطابق (تأكدي من تسجيل الدخول بنفس الحساب).'
      : 'تعذّر الاتصال بالموقع' + (r && r._http ? ' (رمز: ' + r._http + ')' : '') + '.';
    return page_('تعذّر الربط', detail);
  }
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(1).create();
  const site = P.getProperty('SITE');
  const em = r.email ? '<br><small style="color:#666">الحساب المربوط: ' + r.email + '</small>' : '';
  const cnt = '<br><small style="color:#1f6b45">تم فحص ومزامنة ' + (r.synced || 0) + ' محادثة من آخر 7 أيام.</small>';
  const backBtn = site ? '<p style="margin-top:18px"><a href="' + site + '" style="display:inline-block;padding:10px 24px;background:#8C1A2C;color:#fff;text-decoration:none;border-radius:8px;font-weight:bold">العودة إلى الصندوق</a></p>' : '';
  return page_('تم ربط بريدك ✓', 'ستصل أي رسالة جديدة إلى الموقع خلال دقيقة تقريبًا.' + em + cnt + backBtn);
}
