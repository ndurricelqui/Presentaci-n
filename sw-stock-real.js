const CACHE = 'stock-real-v2';

self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// HTML siempre desde red; Trello/Sheets nunca se cachean; el resto (fonts) sí
self.addEventListener('fetch', e => {
  if (e.request.url.includes('docs.google.com') || e.request.url.includes('spreadsheets.google.com')) return;
  const isHTML = e.request.destination === 'document' || e.request.url.endsWith('.html');
  if (isHTML) {
    e.respondWith(fetch(e.request));
    return;
  }
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(r => {
        const copy = r.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        return r;
      });
    })
  );
});
