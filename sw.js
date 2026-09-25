self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open('grizzly-gains-v1').then((cache) => {
      return cache.addAll([
        './index.html',
        './icon.png',
        './small-icon.png',
        './manifest.json'
      ]);
    })
  );
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((response) => {
      return response || fetch(e.request);
    })
  );
});
