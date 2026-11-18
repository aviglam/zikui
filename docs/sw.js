/* Service Worker – האפליקציה נפתחת גם בלי קליטה. קריאות ל-Apps Script לא עוברות כאן (POST לדומיין אחר). */
const CACHE = 'zikui-v1.0.2';
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
// רשת קודם (כדי שעדכונים יגיעו מיד), ואם אין תשובה תוך 3 שניות או אין קליטה – מהמטמון
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(caches.open(CACHE).then(async cache => {
    const fromCache = async () =>
      (await cache.match(req, { ignoreSearch: true })) ||
      (req.mode === 'navigate' ? await cache.match('index.html') : null);
    const net = fetch(req, { cache: 'no-cache' })
      .then(res => { if (res.ok) cache.put(req, res.clone()); return res; });
    const timeout = new Promise(r => setTimeout(r, 3000, null));
    try {
      const res = await Promise.race([net, timeout]);
      if (res) return res;
      return (await fromCache()) || (await net);
    } catch (_) {
      return (await fromCache()) || Response.error();
    }
  }));
});
