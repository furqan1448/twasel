// ===== سكربت ربط بريد كل مستخدمة الخاص (مشروع واحد تنشرينه أنتِ مرة واحدة) =====
// إعدادات النشر (مهمة جدًا لكي يرتبط بريد كل مستخدمة الخاص وليس بريد الموقع):
//   نشر ← عملية نشر جديدة ← تطبيق ويب
//   1) تنفيذ باسم (Execute as): «المستخدم الذي يصل إلى تطبيق الويب» (User accessing the web app)
//   2) من يمكنه الوصول (Who has access): «أي شخص لديه حساب Google» (Anyone with Google account)
// ثم انسخي رابط النشر (ينتهي بـ /exec) وضعيه في index.html مكان MAIL_LINK.

function cfg_() {
  const P = PROP_();
  return {
    key: P.getProperty('KEY'),
    api: P.getProperty('API'),
    site: P.getProperty('SITE') || '',
    regEmail: P.getProperty('REG_EMAIL') || ''
  };
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

function getFallback_(api, key, em, items) {
  try {
    const b64 = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(items)).getBytes());
    const url = api + '?action=mailIn&k=' + encodeURIComponent(key) + '&em=' + encodeURIComponent(em) + '&items=' + encodeURIComponent(b64);
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
  const me = (Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '').toLowerCase();

  // التحقق من أن البريد المفتوح في المتصفح يطابق البريد المسجّل في حساب المستخدمة بالموقع
  if (forceAll === 'force' && c.regEmail && me && c.regEmail !== me) {
    // force override allowed
  } else if (c.regEmail && me && c.regEmail !== me) {
    return { ok: false, emailMismatch: true, expected: c.regEmail, actual: me };
  }

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
  if (!forceAll && !items.length && now - Number(P.getProperty('PING') || 0) < 1800000) return { ok: true, count: 0, email: me };
  const payload = JSON.stringify({ action: 'mailIn', token: c.key, em: me, items: items });
  let r = post_(c.api, payload);
  if ((!r || !r.ok) && r && r.auth !== false && !r.emailMismatch) {
    const g = getFallback_(c.api, c.key, me, items);
    if (g && (g.ok || g.emailMismatch)) r = g;
  }
  if ((!r || !r.ok) && !r.emailMismatch && c.site) {
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

// المشغّل الزمني يمرّر كائن حدث للدالة، فنستخدم دالة وسيطة حتى لا تُعاد مزامنة كل الرسائل كل دقيقة
function tick() { sync(false); }

function page_(title, text) {
  return HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:460px;margin:10vh auto;padding:24px;text-align:center;line-height:2;border:1px solid #ebe1e2;border-radius:16px">' +
    '<h2 style="color:#8C1A2C">' + title + '</h2><div>' + text + '</div></div>'
  ).setTitle(title);
}

function doGet(e) {
  const p = e.parameter || {}, P = PROP_();
  if (p.off) {
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
    P.deleteAllProperties();
    return page_('تم إيقاف الربط', 'لن تُرسَل رسائل بريدك الخاص إلى الموقع بعد الآن.');
  }
  if (!p.k || !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(p.api || '')) {
    return page_('رابط غير صحيح', 'افتحي رابط الربط من زر «ربط البريد» داخل حسابك في الموقع.');
  }
  const regEmail = String(p.em || '').trim().toLowerCase();
  P.setProperties({
    KEY: p.k,
    API: p.api,
    SITE: /^https:\/\//.test(p.site || '') ? p.site : '',
    REG_EMAIL: regEmail,
    LAST: '0',
    PING: '0'
  });
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const r = sync(p.force ? 'force' : true);
  if (!r || !r.ok) {
    P.deleteAllProperties();
    if (r && r.emailMismatch) {
      const targetEmail = r.expected || regEmail;
      const retryUrl = ScriptApp.getService().getUrl() + '?k=' + encodeURIComponent(p.k) + '&api=' + encodeURIComponent(p.api) + '&site=' + encodeURIComponent(p.site || '') + '&em=' + encodeURIComponent(targetEmail);
      return page_(
        'هذا ليس البريد المسجّل في حسابك',
        'حسابك في الموقع مسجّل بالبريد: <b style="color:#8C1A2C;direction:ltr">' + targetEmail + '</b><br>' +
        'بينما المتصفح مفتوح بحساب مختلف (<b style="direction:ltr">' + (r.actual || '') + '</b>).<br><br>' +
        'ارجعي إلى الموقع وافتحي «ربط البريد» في نافذة خاصة (Incognito/Private) وسجّلي الدخول بالبريد المسجّل فقط.'
      );
    }
    const detail = r && r.auth === false
      ? 'مفتاح الربط غير مطابق (تأكدي من تسجيل الدخول بنفس الحساب).'
      : 'تعذّر الاتصال بالموقع' + (r && r._http ? ' (رمز: ' + r._http + ')' : '') + '.';
    const dbg = String(JSON.stringify(r || {}) + ' | API=' + (p.api || '')).slice(0, 400).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return page_('تعذّر الربط', detail + '<br><br><small dir="ltr" style="color:#999;word-break:break-all">' + dbg + '</small>');
  }
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
  const site = P.getProperty('SITE');
  const em = r.email ? '<br><small style="color:#666">تم ربط بريدك الخاص: <b style="direction:ltr">' + r.email + '</b></small>' : '';
  const cnt = '<br><small style="color:#1f6b45">تمت مزامنة ' + (r.synced || 0) + ' محادثة من بريدك الخاص.</small>';
  const backBtn = site ? '<p style="margin-top:18px"><a href="' + site + '" target="_top" style="display:inline-block;padding:10px 24px;background:#8C1A2C;color:#fff;text-decoration:none;border-radius:8px;font-weight:bold">العودة إلى الصندوق</a></p>' : '';
  return page_('تم ربط بريدك الخاص بنجاح ✓', 'ستصل رسائل بريدك الخاص إلى صندوقك في الموقع خلال دقيقة تقريبًا.' + em + cnt + backBtn);
}
