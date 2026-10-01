/* =====================================================================
   COMEDOR PETROCENO · Service Worker
   ---------------------------------------------------------------------
   Todo el app es LOCAL: no depende de ningun CDN para funcionar.
   Este SW cachea los archivos propios y sirve la app sin internet.
   ===================================================================== */
const CACHE = 'comedor-petroceno-v3';

const ASSETS = [
  './',
  './index.html',
  './app-xlsx.js',
  './manifest.json',
  './offline.html',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable.png',
  // Motor PDF embebido: el PDF tambien funciona sin internet
  './vendor/jspdf.umd.min.js',
  './vendor/jspdf.plugin.autotable.min.js'
];

/* ---------- Instalacion: cachea todo lo local ---------- */
self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    // addAll falla entero si un recurso falla; aqui vamos uno por uno.
    await Promise.all(ASSETS.map(u => cache.add(u).catch(() => {})));
    await self.skipWaiting();
  })());
});

/* ---------- Activacion: limpia caches viejos ---------- */
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

/* ---------- Boton para activar la actualizacion al instante ---------- */
self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

/* ---------- Peticiones ---------- */
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Navegacion: red primero, si no hay red servimos la app cacheada
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        return await fetch(req);
      } catch (e) {
        const c = await caches.match('./index.html');
        return c || caches.match('./offline.html') || new Response(
          '<h1>Sin conexión</h1><p>Abre la app desde el icono instalado.</p>',
          { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
        );
      }
    })());
    return;
  }

  // Otros origenes (p.ej. jsPDF): red primero, cache como respaldo
  if (url.origin !== self.location.origin) {
    event.respondWith((async () => {
      try { return await fetch(req); }
      catch (e) { return (await caches.match(req)) || Response.error(); }
    })());
    return;
  }

  // Mismo origen: cache primero, y refrescamos en segundo plano
  event.respondWith((async () => {
    const cached = await caches.match(req);
    const fetching = fetch(req).then(res => {
      if (res && res.status === 200 && res.type === 'basic') {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      }
      return res;
    }).catch(() => cached);
    return cached || fetching;
  })());
});
