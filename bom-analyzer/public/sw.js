// This service worker exists only to remove itself.
//
// It used to cache the app shell, with index.html fetched network-first and
// app.js and styles.css served cache-first from a fixed cache name. Every
// release after the first therefore handed the browser a new index.html and an
// old app.js, and the app died on the first line that touched an element the
// old script had never heard of.
//
// Caching the shell bought nothing to begin with: this app is a front end for
// three supplier APIs and its own backend, so a shell that loads without a
// network has nothing to show. Rather than make the caching correct, the
// caching is gone — and this file stays behind to clean up after the copies
// already installed in people's browsers.

self.addEventListener('install', () => self.skipWaiting());

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.map((key) => caches.delete(key))))
      .then(() => self.registration.unregister())
      // Without a reload the page keeps whatever this worker already served it.
      .then(() => self.clients.matchAll({ type: 'window' }))
      .then((clients) => clients.forEach((client) => client.navigate(client.url)))
      .catch(() => {})
  );
});

// Everything goes straight to the network while this worker is still alive.
self.addEventListener('fetch', () => {});
