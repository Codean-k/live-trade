// LIVE TRADE - Service Worker
// 캐시 전략:
// - index.html, manifest.json: network-first (항상 최신 시도, 실패 시 캐시)
// - scores.json, tickers.json: network-first (데이터는 최신 우선)
// - 그 외 정적 자원: cache-first

const CACHE_NAME = 'live-trade-v1';
const CORE_ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
];

// 설치: 핵심 파일 미리 캐시
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS))
  );
  self.skipWaiting();
});

// 활성화: 이전 버전 캐시 청소
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// fetch: network-first 전략
// 네트워크 응답이 성공하면 캐시 갱신, 실패하면 캐시 반환
self.addEventListener('fetch', (event) => {
  const { request } = event;

  // GET 요청만 처리 (POST 등은 그냥 통과)
  if (request.method !== 'GET') return;

  // chrome-extension 등 비표준 스킴 무시
  if (!request.url.startsWith('http')) return;

  event.respondWith(
    fetch(request)
      .then((response) => {
        // 성공 시 캐시에 저장 후 응답 반환
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, clone));
        }
        return response;
      })
      .catch(() => {
        // 네트워크 실패 시 캐시에서 찾아 반환
        return caches.match(request).then((cached) => {
          if (cached) return cached;
          // 캐시도 없으면 index.html 반환 (SPA 폴백)
          if (request.mode === 'navigate') {
            return caches.match('/index.html');
          }
        });
      })
  );
});
