/* Service Worker – האפליקציה נפתחת גם בלי קליטה. קריאות ל-Apps Script לא עוברות כאן (POST לדומיין אחר). */
const CACHE = 'zikui-v1.0.0';
const SHELL = [
  './', 'index.html', 'app.css', 'app.js', 'scanner.js', 'demo.js',
  'manifest.webmanifest', 'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png',
  'vendor/zxing-browser.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// stale-while-revalidate: מגיש מהמטמון מיד, ומעדכן ברקע לפעם הבאה
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(caches.open(CACHE).then(async cache => {
    const hit = await cache.match(req, { ignoreSearch: true });
    const net = fetch(req).then(res => { if (res.ok) cache.put(req, res.clone()); return res; }).catch(() => null);
    return hit || (await net) || (req.mode === 'navigate' ? cache.match('index.html') : Response.error());
  }));
});
