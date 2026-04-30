/**
 * Service Worker v6 — Excel → CSV 変換ツール
 *
 * 【戦略】Stale While Revalidate (SWR)
 * ① キャッシュがあれば即返す
 * ② 裏側でネットワークから最新版を取得してキャッシュを更新
 * ③ 次回起動時に新しい内容が反映される
 */
'use strict';

const CACHE = 'excel2csv-v6';

const PRECACHE = [
  'index.html',
  'css/style.css?v=5',
  'js/app.js?v=6',
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
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const url = event.request.url;
  if (url.startsWith('blob:') || url.startsWith('data:') || url.startsWith('chrome-extension:')) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE);

    const cached = await cache.match(event.request, { ignoreSearch: true });

    const fetchPromise = fetch(event.request)
      .then(res => {
        if (res && res.ok) cache.put(event.request, res.clone());
        return res;
      })
      .catch(() => null);

    if (cached) return cached;

    const res = await fetchPromise;
    if (res) return res;

    const accept = event.request.headers.get('accept') || '';
    if (accept.includes('text/html')) {
      const fb = await cache.match('index.html');
      if (fb) return fb;
    }
    return new Response('Offline', { status: 503 });
  })());
});
