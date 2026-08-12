// GALERIA TAPIR — service worker
// Solo lo necesario para que el navegador permita "Agregar a pantalla de inicio".
// No cachea fotos (siempre tienen que verse actualizadas), solo el shell básico.

var CACHE = 'tapir-galeria-v1';
var ARCHIVOS_BASE = ['./index.html', './logo_tapirmedia.svg', './manifest.json'];

self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(function (cache) { return cache.addAll(ARCHIVOS_BASE); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', function (e) {
  // network-first: siempre intenta traer la version mas nueva,
  // y si no hay conexion, usa lo que haya en cache como respaldo
  e.respondWith(
    fetch(e.request).catch(function () { return caches.match(e.request); })
  );
});
