// Офлайн-кеш: застосунок відкривається без інтернету; запити до Google не кешуються.
const CACHE = 'trener-v3';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || /^script\.google(usercontent)?\.com$/.test(url.hostname)) return;
  const isPage = req.mode === 'navigate';
  e.respondWith(
    caches.match(isPage ? './index.html' : req, { ignoreSearch: isPage }).then(cached => {
      const net = fetch(req).then(res => {
        if (res && (res.ok || res.type === 'opaque')) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(isPage ? './index.html' : req, copy));
        }
        return res;
      }).catch(() => cached);
      return cached || net;
    })
  );
});
