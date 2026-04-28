/**
 * Service Worker v7 — Excel → CSV 変換ツール（外字チェック機能追加）
 * Network First 戦略：オンライン時は常に最新取得、オフライン時はキャッシュ返却
 * GitHub Pages 対応：相対パスでキャッシュ登録
 */
'use strict';

const CACHE = 'excel2csv-v7';

// self.location.pathname からベースパスを取得（例: /mailing-csv-app/sw.js → /mailing-csv-app/）
const BASE = self.location.pathname.replace(/sw\.js$/, '');

const PRECACHE_PATHS = [
  '',           // index.html
  'index.html',
  'css/style.css',
  'js/app.js',
  'js/gaiji.js',
  'manifest.json',
  'icons/icon.svg',
];

const PRECACHE_EXTERNAL = [
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE).then(cache => {
      const localReqs = PRECACHE_PATHS.map(p => new Request(BASE + p));
      const extReqs   = PRECACHE_EXTERNAL.map(u => new Request(u, { mode: 'cors' }));
      return Promise.all([...localReqs, ...extReqs].map(req =>
        fetch(req).then(r => r.ok ? cache.put(req, r) : null).catch(() => null)
      ));
    })
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
      // HTMLリクエストにはindex.htmlをフォールバック
      const accept = event.request.headers.get('accept') || '';
      if (accept.includes('text/html')) {
        const fb = await cache.match(BASE + 'index.html');
        if (fb) return fb;
      }
      return new Response('Offline', { status: 503 });
    }
  })());
});
