// MSI SalesOps CRM — Service Worker
// Cache name includes version — changing it forces old cache eviction on next activate.
const CACHE = 'salesops-01.09.26.1';
const SHELL = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-apple.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  // Delete every cache that doesn't match the current version
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // External / dynamic APIs → always network, never cache
  if (url.hostname.includes('firebase') || url.hostname.includes('googleapis') ||
      url.hostname.includes('cloudflare') || url.hostname.includes('unpkg') ||
      url.hostname.includes('jsdelivr') || url.hostname.includes('fonts')) {
    return;
  }

  // index.html → NETWORK FIRST so the latest version loads when online.
  // Falls back to cache only when offline.
  if (url.pathname.endsWith('index.html') || url.pathname === '/' || url.pathname.endsWith('/')) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          if (res && res.status === 200) {
            const cloned = res.clone(); // clone before any async op consumes the body
            caches.open(CACHE).then(c => c.put(e.request, cloned));
          }
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // Other shell assets (icons, manifest) → cache first, fallback network
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      if (res && res.status === 200 && e.request.method === 'GET') {
        const cloned = res.clone(); // clone before any async op consumes the body
        caches.open(CACHE).then(c => c.put(e.request, cloned));
      }
      return res;
    }))
  );
});
