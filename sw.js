// 로지판 Service Worker
const CACHE_NAME = 'logipan-v1';

self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(clients.claim());
});

self.addEventListener('fetch', e => {
  // 네트워크 우선, 실패하면 캐시
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});