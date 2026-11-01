/* Service Worker: hält nur die App-Dateien vor. Daten (Microsoft Graph, Anmeldung) werden nie zwischengespeichert. */
const CACHE = 'ze-v1';
const FILES = ['./', 'index.html', 'styles.css', 'config.js', 'calc.js', 'app.js', 'lib/msal-browser.min.js', 'manifest.webmanifest', 'icons/icon-192.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  // Netzwerk zuerst (damit Updates sofort greifen), Cache nur als Rückfall
  e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; }).catch(() => caches.match(e.request)));
});
