/**
 * PWA用PNGアイコンをCanvasで動的生成し、
 * Cache Storage に保存する
 * → インストール時にPNGアイコンが要求されても応答できる
 */
(function () {
  'use strict';

  function drawIcon(size) {
    const canvas = document.createElement('canvas');
    canvas.width = size;
    canvas.height = size;
    const ctx = canvas.getContext('2d');
    const s = size;

    // 角丸背景
    const r = s * 0.18;
    ctx.beginPath();
    ctx.moveTo(r, 0);
    ctx.lineTo(s - r, 0);
    ctx.quadraticCurveTo(s, 0, s, r);
    ctx.lineTo(s, s - r);
    ctx.quadraticCurveTo(s, s, s - r, s);
    ctx.lineTo(r, s);
    ctx.quadraticCurveTo(0, s, 0, s - r);
    ctx.lineTo(0, r);
    ctx.quadraticCurveTo(0, 0, r, 0);
    ctx.closePath();
    ctx.fillStyle = '#2563eb';
    ctx.fill();

    // テーブルグリッド
    const pad = s * 0.14;
    const gw = s - pad * 2;
    const gh = s - pad * 2;
    const rows = 5, cols = 3;
    const cw = gw / cols;
    const rh = gh / rows;

    // セル背景
    ctx.fillStyle = 'rgba(255,255,255,0.12)';
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        ctx.fillRect(
          pad + col * cw + 2,
          pad + row * rh + 2,
          cw - 4,
          rh - 4
        );
      }
    }

    // ヘッダー行（1行目）を明るく
    ctx.fillStyle = 'rgba(255,255,255,0.55)';
    for (let col = 0; col < cols; col++) {
      ctx.fillRect(pad + col * cw + 2, pad + 2, cw - 4, rh - 4);
    }

    // 下矢印（ダウンロードアイコン）
    const lw = s * 0.055;
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = lw;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';

    const ax = s * 0.5;
    const ay1 = s * 0.52;
    const ay2 = s * 0.76;
    const aw = s * 0.14;

    ctx.beginPath();
    ctx.moveTo(ax, ay1);
    ctx.lineTo(ax, ay2);
    ctx.stroke();

    ctx.beginPath();
    ctx.moveTo(ax - aw, ay2 - aw * 0.9);
    ctx.lineTo(ax, ay2);
    ctx.lineTo(ax + aw, ay2 - aw * 0.9);
    ctx.stroke();

    // 底線
    ctx.beginPath();
    ctx.moveTo(ax - aw * 1.4, ay2 + lw);
    ctx.lineTo(ax + aw * 1.4, ay2 + lw);
    ctx.stroke();

    return canvas;
  }

  async function saveIconToCache(canvas, path, cacheName) {
    return new Promise(resolve => {
      canvas.toBlob(async blob => {
        if (!blob) { resolve(); return; }
        try {
          const cache = await caches.open(cacheName);
          const response = new Response(blob, {
            headers: { 'Content-Type': 'image/png' }
          });
          await cache.put(path, response);
          console.log('[Icons] Cached:', path);
        } catch (e) {
          console.warn('[Icons] Cache failed:', path, e);
        }
        resolve();
      }, 'image/png');
    });
  }

  // キャッシュ名は SW と揃える
  const CACHE_NAME = 'excel2csv-v3';

  window.generateAndCacheIcons = async function () {
    if (!('caches' in window)) return;
    try {
      const c192 = drawIcon(192);
      const c512 = drawIcon(512);
      await Promise.all([
        saveIconToCache(c192, 'icons/icon-192.png', CACHE_NAME),
        saveIconToCache(c512, 'icons/icon-512.png', CACHE_NAME),
      ]);
    } catch (e) {
      console.warn('[Icons] Generation failed:', e);
    }
  };
})();
