const CACHE_NAME = 'cassetto-shell-v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon.svg'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(APP_SHELL);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Cache-first solo per le GET dello stesso origine (l'app shell).
// Le chiamate all'API Apps Script (POST, cross-origin) passano sempre dalla rete.
self.addEventListener('fetch', function(event) {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    caches.match(req).then(function(cached) {
      return cached || fetch(req).then(function(res) {
        const copy = res.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(req, copy);
        });
        return res;
      }).catch(function() {
        return caches.match('./index.html');
      });
    })
  );
});
