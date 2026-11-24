/* Service Worker: يستقبل تنبيهات الدفع (Web Push) من الخادم ويعرضها حتى والموقع مغلق */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

function idb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('twasel-sw', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function kv(k) {
  try {
    const db = await idb();
    return await new Promise(res => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(null); });
  } catch (e) { return null; }
}

self.addEventListener('push', e => e.waitUntil((async () => {
  let items = [], test = false, msg = null;
  try {
    const api = await kv('api'), tok = await kv('tok'), mode = await kv('mode');
    if (api && tok) {
      // نعيد المحاولة حتى 3 مرات (بينها 1.2 ثانية) إن لم تصل التفاصيل بعد، كي لا يظهر نص عام بدل عنوان التذكير
      for (let a = 0; a < 3; a++) {
        try {
          const r = await fetch(api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'remDue', token: tok, mode: mode === 'all' ? 'all' : 'rem' }) });
          const j = await r.json();
          if (j && j.ok) { items = j.rows || []; test = !!j.test; msg = j.msg || null; if (items.length || test || msg) break; }
        } catch (err) {}
        if (a < 2) await new Promise(res => setTimeout(res, 1200));
      }
    }
  } catch (err) {}
  const base = self.registration.scope, icon = 'icon-192.png';
  if (msg) {   // رسالة جديدة (للأجهزة المفعّل عليها «جميع الرسائل»)
    const src = msg.source ? ' عبر ' + msg.source : '';
    const one = msg.n <= 1;
    await self.registration.showNotification(one ? 'رسالة جديدة' : msg.n + ' رسائل جديدة', {
      body: one ? (msg.name ? 'من ' + msg.name + src : 'وصلتك رسالة جديدة') : 'وصلتك رسائل جديدة، افتحي صِلة لعرضها.',
      icon, tag: 'msg-new', renotify: true,
      data: { url: base + (one && msg.id ? '#msg=' + encodeURIComponent(msg.id) : '#inbox') }
    });
    if (!items.length) return;
  }
  if (test && !items.length) {
    return self.registration.showNotification('تنبيه تجريبي ✓', { body: 'التنبيهات تعمل على هذا الجهاز.', icon, tag: 'rem-test', data: { url: base } });
  }
  if (!items.length) {   // يجب إظهار إشعار دائمًا (شرط آبل)، حتى لو تعذّر جلب التفاصيل
    return self.registration.showNotification('لديك تنبيه جديد', { body: 'افتحي صِلة لعرض التفاصيل.', icon, tag: 'rem-generic', data: { url: base + '#rem' } });
  }
  await Promise.all(items.slice(0, 5).map(r => self.registration.showNotification('تذكير: ' + r.title, {
    body: (r.name ? r.name + ' · ' : '') + 'حان موعد المتابعة',
    icon, tag: 'rem-' + r.id, renotify: true,
    data: { url: base + (r.msgId ? '#msg=' + encodeURIComponent(r.msgId) : '#rem') }
  })));
})()));

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.url) || self.registration.scope;
  e.waitUntil((async () => {
    const cs = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of cs) {
      if ('focus' in c) { await c.focus(); c.postMessage({ type: 'open', url }); return; }
    }
    await self.clients.openWindow(url);
  })());
});
