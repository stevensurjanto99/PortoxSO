// Network-first: selalu ambil versi terbaru, pakai cache hanya saat offline.
const V = 'so-v4';
self.addEventListener('install', e => e.waitUntil(caches.open(V).then(c => c.addAll(['./', 'index.html', 'manifest.json']))));
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x))))));
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); return r; })
    .catch(() => caches.match(e.request)));
});
