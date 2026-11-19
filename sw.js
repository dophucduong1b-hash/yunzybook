const CACHE_NAME = 'yunzybook-pwa-v1.6.2';

const PRECACHE_URLS = [
  './',
  './index.html',
  './styles.css',
  './manifest.webmanifest',
  './apple-touch-icon.png',
  './apple-touch-icon-precomposed.png',
  './js/app.js',
  './js/storage.js',
  './js/zip.js',
  './js/docx.js',
  './js/reader.js',
  './js/ui-helpers.js',
  './vendor/pdfjs/pdf.min.mjs',
  './vendor/pdfjs/pdf.worker.min.mjs',
  './icons/apple-touch-icon.png',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/favicon-32.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      Promise.all(
        PRECACHE_URLS.map((url) =>
          cache.add(new Request(url, { cache: 'reload' }))
        )
      )
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  const isNavigate = req.mode === 'navigate';
  const networkUpdate = fetch(req)
    .then((resp) => {
      if (resp && resp.status === 200 && (isNavigate || resp.type === 'basic')) {
        const copy = resp.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(isNavigate ? './index.html' : req, copy));
      }
      return resp;
    })
    .catch(() => null);

  event.waitUntil(networkUpdate);

  // Stale-While-Revalidate: instant cached response while refreshing cache in background
  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    const cached = isNavigate
      ? await cache.match('./index.html')
      : await cache.match(req, { ignoreSearch: true });
    if (cached) return cached;
    const fresh = await networkUpdate;
    return fresh || Response.error();
  })());
});
