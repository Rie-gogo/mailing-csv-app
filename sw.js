/**
 * Service Worker v6 — Excel → CSV 変換ツール（外字チェック機能追加）
 * Network First 戦略：オンライン時は常に最新取得、オフライン時はキャッシュ返却
 */
'use strict';

const CACHE = 'excel2csv-v6';

const PRECACHE = [
  'index.html',
  'css/style.css?v=6',
  'js/app.js?v=6',
  'js/gaiji.js?v=1',
  'manifest.json',
  'icons/icon.svg',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE).then(cache =>
      Promise.all(PRECACHE.map(url => {
        const req = url.startsWith('http') ? new Request(url, { mode: 'cors' }) : new Request(url);
        return fetch(req).then(r => r.ok ? cache.put(url, r) : null).catch(() => null);
      }))
    )
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const url = event.request.url;
  if (url.startsWith('blob:') || url.startsWith('data:') || url.startsWith('chrome-extension:')) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    try {
      // Network First：最新を取得してキャッシュ更新
      const res = await fetch(event.request);
      if (res && res.ok) cache.put(event.request, res.clone());
      return res;
    } catch {
      // オフライン時：キャッシュから返す
      const cached = await cache.match(event.request, { ignoreSearch: true });
      if (cached) return cached;
      const accept = event.request.headers.get('accept') || '';
      if (accept.includes('text/html')) {
        const fb = await cache.match('index.html');
        if (fb) return fb;
      }
      return new Response('Offline', { status: 503 });
    }
  })());
});
