/* Service worker: la app abre sin señal. Sube CACHE cuando cambies archivos. */
const CACHE = 'salida-vin-v2.0.5';
const SHELL = ['./', 'index.html', 'app.js', 'config.js', 'manifest.json', 'icons/icon-180.png', 'icons/icon-192.png', 'icons/icon-512.png'];
const LIBS = ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Datos de Supabase: siempre en vivo, nunca desde caché
  if (url.hostname.endsWith('.supabase.co') || url.pathname.includes('/rest/v1/') || url.pathname.includes('/storage/v1/')) return;
  // Librería de Excel: caché primero
  if (LIBS.includes(req.url)) {
    e.respondWith(caches.match(req).then(m => m || fetch(req).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); return r; })));
    return;
  }
  if (url.origin !== location.origin) return;
  // Archivos de la app: red primero (para recibir actualizaciones), caché si no hay señal
  e.respondWith(
    fetch(req).then(r => { if (r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); } return r; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(m => m || (req.mode === 'navigate' ? caches.match('index.html') : undefined)))
  );
});
