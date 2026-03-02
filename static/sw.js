const CACHE_NAME = "engitrack-v2";

self.addEventListener("install", function(event) {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(["/"]);
    }).catch(function() {})
  );
});

self.addEventListener("activate", function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  return self.clients.claim();
});

self.addEventListener("fetch", function(event) {
  if (event.request.method !== "GET") return;
  event.respondWith(
    fetch(event.request).catch(function() {
      return caches.match(event.request).then(function(r) {
        return r || new Response("Offline - please reconnect to use EngiTrack Combiner.", {
          status: 503,
          headers: { "Content-Type": "text/plain" }
        });
      });
    })
  );
});
