// MSI SalesOps CRM — Service Worker
const CACHE = 'salesops-v1';
const SHELL = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-apple.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Firebase / external API calls → always network
  if (url.hostname.includes('firebase') || url.hostname.includes('googleapis') ||
      url.hostname.includes('cloudflare') || url.hostname.includes('unpkg') ||
      url.hostname.includes('jsdelivr') || url.hostname.includes('fonts')) {
    return;
  }
  // Shell files → cache first, fallback network
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      if (res && res.status === 200 && e.request.method === 'GET') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }))
  );
});
