// GALERIA TAPIR — service worker
// Solo lo necesario para que el navegador permita "Agregar a pantalla de inicio".
// No cachea fotos (siempre tienen que verse actualizadas), solo el shell básico.

var CACHE = 'tapir-galeria-v2';
var ARCHIVOS_BASE = ['./index.html', './logo_tapirmedia.svg', './manifest.json'];

self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(function (cache) { return cache.addAll(ARCHIVOS_BASE); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (nombres) {
      return Promise.all(
        nombres.filter(function (n) { return n !== CACHE; }).map(function (n) { return caches.delete(n); })
      );
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  // el documento principal (index.html) siempre se pide fresco a la red,
  // nunca debe quedar una version vieja pegada en cache
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request).catch(function () {
        return caches.match('./index.html').then(function (resp) {
          return resp || new Response('', { status: 503, statusText: 'Sin conexión' });
        });
      })
    );
    return;
  }

  // el resto (logo, manifest, etc.) si funciona con network-first + respaldo
  e.respondWith(
    fetch(e.request).catch(function () {
      return caches.match(e.request).then(function (resp) {
        return resp || new Response('', { status: 503, statusText: 'Sin conexión' });
      });
    })
  );
});
