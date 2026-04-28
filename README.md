# Excel → CSV 変換ツール（オフライン対応 PWA）

## 概要

Excel ファイル（.xlsx / .xls）をブラウザ内だけで CSV に変換する
**完全オフライン動作 PWA アプリ**です。

- ✅ **先頭ゼロ保持**: セルの「表示文字列 (w)」を優先取得 → 0012 → "0012"
- ✅ **全セルをダブルクォートで囲む**: `"値"` 形式で出力
- ✅ **ファイルがサーバーへ送信されない**: 100% ブラウザ内処理
- ✅ **ドラッグ&ドロップ + ファイル選択** どちらも対応
- ✅ **複数シート対応**: 全シートまとめて変換・ダウンロード可能
- ✅ **文字コード選択**: UTF-8(BOM付き) / UTF-8 / Shift-JIS

---

## 動作環境

| ブラウザ | 対応状況 |
|---|---|
| Chrome / Edge | ✅ 完全対応（PWAインストールも可） |
| Firefox | ✅ 動作可（一部PWA機能制限あり） |
| Safari (iOS/macOS) | ✅ 動作可 |

---

## ファイル構成

```
index.html           メイン画面
css/style.css        スタイルシート
js/app.js            変換ロジック（SheetJS 使用）
sw.js                Service Worker（オフライン対応）
manifest.json        PWAマニフェスト
icons/icon.svg       アプリアイコン（SVG）
```

---

## 主な機能

### 変換オプション

| オプション | 説明 |
|---|---|
| 全シートを変換 | 複数シートをまとめて変換 |
| 空行をスキップ | 全セルが空の行を除外 |
| 文字コード | UTF-8 BOM付き / UTF-8 / Shift-JIS |

### 先頭ゼロの保持方法

SheetJS の `cellText: true` オプションで読み込み、セルの `w`（書式済み表示文字列）を優先して取得します。
例：Excel で `0001234` と表示されているセルは `"0001234"` として出力されます。

---

## オフライン対応の仕組み

Service Worker（`sw.js`）が以下のファイルをキャッシュします：

- `index.html`
- `css/style.css`
- `js/app.js`
- `manifest.json`
- `https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js` (SheetJS)

初回アクセス後はネット未接続でも完全に動作します。

---

## 開発スタック

- **SheetJS (xlsx 0.18.5)**: Excel 読み込み・パース
- **Service Worker API**: オフラインキャッシュ
- **Web Streams API / TextEncoder**: 文字コード変換
- **File API / Blob API**: ファイル読み込み・ダウンロード

---

## 既知の制限

- **Shift-JIS 出力**: 一部ブラウザ（Safari 等）では UTF-8 BOM付きにフォールバック
- **PWAアイコン（PNG）**: 現在は SVG のみ。PNG が必要な場合は `icons/generate-icons.html` を開いてダウンロード

---

## 今後の改善案

- [ ] ZIP で全シートをまとめてダウンロード（JSZip 統合）
- [ ] CSV プレビューでの列幅調整
- [ ] 変換ログ表示
- [ ] ドラッグで列を並び替え
