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
  let items = [], test = false;
  try {
    const api = await kv('api'), tok = await kv('tok');
    if (api && tok) {
      const r = await fetch(api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'remDue', token: tok }) });
      const j = await r.json();
      if (j && j.ok) { items = j.rows || []; test = !!j.test; }
    }
  } catch (err) {}
  const base = self.registration.scope, icon = 'icon-192.png';
  if (test && !items.length) {
    return self.registration.showNotification('تنبيه تجريبي ✓', { body: 'تنبيهات التذكير تعمل على جوالك.', icon, tag: 'rem-test', data: { url: base } });
  }
  if (!items.length) {   // يجب إظهار إشعار دائمًا (شرط آبل)، حتى لو تعذّر جلب التفاصيل
    return self.registration.showNotification('لديك تذكير مستحق', { body: 'افتحي صِلة لعرض التذكيرات.', icon, tag: 'rem-generic', data: { url: base + '#rem' } });
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
