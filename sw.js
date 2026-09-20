// Minimaler Service Worker – nötig, damit der Browser "Installieren" anbietet.
self.addEventListener('install', function (event) {
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  self.clients.claim();
});

// Einfaches Netzwerk-zuerst-Verhalten (keine echte Offline-Funktion nötig,
// da die App ohnehin eine Internetverbindung zum Google Sheet braucht).
self.addEventListener('fetch', function (event) {
  event.respondWith(
    fetch(event.request).catch(function () {
      return new Response('Keine Internetverbindung.', { status: 503 });
    })
  );
});
