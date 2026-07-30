const CACHE = 'kalkulator-v5';
const FILES = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  // Видаляємо всі старі кеші
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Не втручаємось у сторонні запити (аналітика, GA, ipapi тощо).
  // Хай браузер обробляє їх напряму — SW їх не бачить і не кешує.
  if (url.origin !== self.location.origin) return;

  // Обробляємо лише GET-запити нашого домену
  if (e.request.method !== 'GET') return;

  // Network First — спочатку мережа, потім кеш
  e.respondWith(
    fetch(e.request)
      .then(response => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(e.request))
  );
});
