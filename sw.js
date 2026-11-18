// קבלת סחורה – service worker: keeps the app shell available, always tries the network first,
// and receives price-list files shared from WhatsApp (Android share sheet).
const CACHE = 'kabala-v8';
const SHELL = ['./', 'index.html', 'manifest.json', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== 'kabala-share').map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method === 'POST' && u.origin === location.origin && u.pathname.endsWith('/share')) {
    e.respondWith((async () => {
      try {
        const fd = await e.request.formData();
        const f = fd.getAll('file').find(x => x && x.size) || null;
        if (f) {
          const c = await caches.open('kabala-share');
          await c.put('shared-file', new Response(f, {headers: {'Content-Type': f.type || 'application/octet-stream', 'X-Name': encodeURIComponent(f.name || 'מחירון')}}));
        }
      } catch (err) {}
      return Response.redirect(new URL('./?share=1', self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== 'GET' || u.origin !== location.origin) return; // API calls go straight to the network
  e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; })
    .catch(() => caches.match(e.request).then(r => r || caches.match('index.html'))));
});
