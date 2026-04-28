/**
 * Excel → CSV 変換ツール v4
 * ・複数ファイル同時対応
 * ・変換後ファイル名 = 元のExcelファイル名.csv
 * ・全セルをダブルクォートで囲む
 * ・先頭ゼロ保持（SheetJS の w プロパティ優先）
 */
'use strict';

// ─── 状態 ───────────────────────────────────
// results: [{ excelName, baseName, sheets:[{sheetName,csvText}], error }]
let results = [];
let previewIndex = 0; // 現在プレビュー中のインデックス（results 配列と対応）

// ─── DOM ────────────────────────────────────
const dropZone    = document.getElementById('drop-zone');
const fileInput   = document.getElementById('file-input');
const processing  = document.getElementById('processing');
const procText    = document.getElementById('processing-text');
const errorBox    = document.getElementById('error-box');
const errorSpan   = document.getElementById('error-text');
const resultPanel = document.getElementById('result-panel');
const summaryEl   = document.getElementById('summary');
const previewTabs = document.getElementById('preview-tabs');
const previewNote = document.getElementById('preview-note');
const previewCont = document.getElementById('preview-table-container');
const dlList      = document.getElementById('download-list');
const resetBtn    = document.getElementById('reset-btn');

// オプション
const optAllSheets = () => document.getElementById('opt-all-sheets').checked;
const optSkipEmpty = () => document.getElementById('opt-skip-empty').checked;
const optEncoding  = () => document.getElementById('opt-encoding').value;

// ─── ドラッグ＆ドロップ ──────────────────────
dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', e => { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const files = Array.from(e.dataTransfer.files);
  if (files.length) startConvert(files);
});
dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });

// ─── ファイル選択 ────────────────────────────
fileInput.addEventListener('change', e => {
  const files = Array.from(e.target.files);
  if (files.length) startConvert(files);
  fileInput.value = '';
});

// ─── リセット ───────────────────────────────
resetBtn.addEventListener('click', () => {
  results = [];
  previewIndex = 0;
  resultPanel.classList.add('hidden');
  errorBox.classList.add('hidden');
  dropZone.classList.remove('hidden');
  document.querySelector('.options-panel').classList.remove('hidden');
});

// ─── メイン変換処理 ──────────────────────────
async function startConvert(allFiles) {
  // Excelだけ抽出
  const excelFiles = allFiles.filter(f => /\.(xlsx|xls)$/i.test(f.name));
  if (!excelFiles.length) {
    showError('Excelファイル（.xlsx / .xls）が含まれていません。');
    return;
  }

  hideError();
  setProcessing(true, `0 / ${excelFiles.length} 件処理中...`);
  results = [];

  for (let i = 0; i < excelFiles.length; i++) {
    const f = excelFiles[i];
    setProcessing(true, `${i + 1} / ${excelFiles.length} 件処理中：${f.name}`);
    try {
      const r = await convertFile(f);
      results.push(r);
    } catch (err) {
      results.push({ excelName: f.name, baseName: stripExt(f.name), sheets: [], error: err.message });
    }
  }

  setProcessing(false);
  previewIndex = 0;
  renderResults();
}

// ─── 1ファイル変換 ───────────────────────────
async function convertFile(file) {
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, { type: 'array', cellText: true, cellDates: true, raw: false });

  const targetSheets = optAllSheets() ? wb.SheetNames : [wb.SheetNames[0]];
  const sheets = [];

  for (const name of targetSheets) {
    const ws = wb.Sheets[name];
    if (ws) sheets.push({ sheetName: name, csvText: wsToCsv(ws) });
  }

  return { excelName: file.name, baseName: stripExt(file.name), sheets, error: null };
}

// ─── シート→CSV ──────────────────────────────
function wsToCsv(ws) {
  if (!ws['!ref']) return '';
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows  = [];

  for (let R = range.s.r; R <= range.e.r; R++) {
    const cells = [];
    let hasVal  = false;

    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      const txt  = cellText(cell);
      if (txt !== '') hasVal = true;
      cells.push(txt);
    }

    if (optSkipEmpty() && !hasVal) continue;
    rows.push(cells.map(v => `"${v.replace(/"/g, '""')}"`).join(','));
  }

  return rows.join('\r\n');
}

// セルの表示文字列取得（先頭ゼロ保持・日付書式修正）
function cellText(cell) {
  if (!cell) return '';

  // ① cell.w（Excelの表示書式文字列）が存在する場合は最優先で使用
  //    ただし SheetJS のデフォルト日付書式 "M/D/YY" / "M/D/YYYY" は
  //    Excel 本来の形式ではないため YYYY/M/D 形式に変換する
  if (cell.w != null) {
    const w = String(cell.w).trim();
    // SheetJS デフォルト日付書式パターン: M/D/YY または M/D/YYYY
    const matchShort = w.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);    // M/D/YY
    const matchLong  = w.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);    // M/D/YYYY
    if (matchShort) {
      // このセルが実際に日付型（t:'n' or t:'d'）の場合のみ変換
      // 通常の数値（例: 1/2/34）と区別するために型をチェック
      if (cell.t === 'n' || cell.t === 'd') {
        const y = parseInt(matchShort[3]) + 2000;
        return `${y}/${matchShort[1]}/${matchShort[2]}`;
      }
    }
    if (matchLong) {
      if (cell.t === 'n' || cell.t === 'd') {
        // M/D/YYYY → YYYY/M/D
        return `${matchLong[3]}/${matchLong[1]}/${matchLong[2]}`;
      }
    }
    // 上記以外はそのまま返す（Excel表示書式を尊重）
    return w;
  }

  // ② w がない場合
  // 日付型（cellDates:true で t:'d'、値はJavaScript Date）
  if (cell.t === 'd' && cell.v instanceof Date) {
    const dt = cell.v;
    if (!isNaN(dt.getTime())) {
      const y = dt.getFullYear();
      const m = dt.getMonth() + 1;
      const d = dt.getDate();
      return `${y}/${m}/${d}`;
    }
  }

  // ③ 数値型で w がない場合はシリアル値から変換を試みる
  if (cell.v != null) return String(cell.v);
  return '';
}

// ─── 結果画面描画 ────────────────────────────
function renderResults() {
  dropZone.classList.add('hidden');
  document.querySelector('.options-panel').classList.add('hidden');
  resultPanel.classList.remove('hidden');

  const ok  = results.filter(r => !r.error).length;
  const ng  = results.length - ok;

  // サマリー
  summaryEl.innerHTML = `
    <div class="file-info-item">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
      <span>${results.length} ファイル処理（成功 <strong>${ok}</strong> 件${ng ? ' / 失敗 ' + ng + ' 件' : ''}）</span>
    </div>`;

  // プレビュー切り替えタブ（複数ファイル時）
  if (results.length > 1) {
    previewTabs.innerHTML = results.map((r, i) => `
      <button class="sheet-tab ${i === 0 ? 'active' : ''}" data-i="${i}" title="${escH(r.excelName)}">
        ${escH(r.baseName)}${r.error ? ' ⚠' : ''}
      </button>`).join('');
    previewTabs.querySelectorAll('.sheet-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        previewIndex = +btn.dataset.i;
        previewTabs.querySelectorAll('.sheet-tab').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        renderPreview();
      });
    });
    previewTabs.classList.remove('hidden');
  } else {
    previewTabs.innerHTML = '';
    previewTabs.classList.add('hidden');
  }

  renderPreview();
  renderDownloadList();
}

// ─── プレビュー描画 ──────────────────────────
function renderPreview() {
  const r = results[previewIndex];
  if (!r) return;

  if (r.error) {
    previewNote.textContent = '';
    previewCont.innerHTML = `<p class="empty-preview" style="color:#dc2626">エラー: ${escH(r.error)}</p>`;
    return;
  }

  const csv   = r.sheets[0]?.csvText ?? '';
  const lines = csv.split('\r\n').filter(Boolean);
  previewNote.textContent = `全 ${lines.length} 行 / 先頭 ${Math.min(20, lines.length)} 行を表示`;

  if (!lines.length) {
    previewCont.innerHTML = '<p class="empty-preview">データがありません</p>';
    return;
  }

  const trs = lines.slice(0, 20).map(line => {
    const tds = parseLine(line).map(c => `<td>${escH(c)}</td>`).join('');
    return `<tr>${tds}</tr>`;
  });
  previewCont.innerHTML = `<table class="preview-table"><tbody>${trs.join('')}</tbody></table>`;
}

// ─── ダウンロード一覧描画 ────────────────────
function renderDownloadList() {
  const enc     = optEncoding();
  const okFiles = results.filter(r => !r.error);
  dlList.innerHTML = '';

  if (!okFiles.length) return;

  // ラベル
  const lbl = document.createElement('div');
  lbl.className = 'dl-section-header';
  lbl.textContent = `ダウンロード（${okFiles.length} ファイル）`;
  dlList.appendChild(lbl);

  // ファイルごとのカード
  results.forEach(r => {
    if (r.error) return;

    const card = document.createElement('div');
    card.className = 'dl-file-card';

    // ファイル名ラベル
    const label = document.createElement('div');
    label.className = 'dl-file-label';
    label.innerHTML = `
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
      <span>${escH(r.excelName)}</span>`;
    card.appendChild(label);

    // ボタン行
    const row = document.createElement('div');
    row.className = 'dl-btn-row';

    r.sheets.forEach(sheet => {
      // ファイル名決定：シートが複数ある場合は「basename_シート名.csv」
      const csvName = r.sheets.length > 1
        ? `${r.baseName}_${sanitize(sheet.sheetName)}.csv`
        : `${r.baseName}.csv`;

      const btn = document.createElement('button');
      btn.className = 'download-btn';
      btn.textContent = csvName;

      // ★クロージャで csvText と csvName を確実に束縛
      const csvTextSnapshot = sheet.csvText;
      const csvNameSnapshot = csvName;
      const encSnapshot     = enc;
      btn.addEventListener('click', () => dlCsv(csvTextSnapshot, csvNameSnapshot, encSnapshot));

      row.appendChild(btn);
    });

    card.appendChild(row);
    dlList.appendChild(card);
  });
}

// ─── CSV ダウンロード ────────────────────────
function dlCsv(csvText, filename, enc) {
  let data;
  if (enc === 'utf8bom') {
    const bom  = new Uint8Array([0xEF, 0xBB, 0xBF]);
    const body = new TextEncoder().encode(csvText);
    data = new Uint8Array(bom.length + body.length);
    data.set(bom); data.set(body, bom.length);
  } else {
    data = new TextEncoder().encode(csvText);
  }
  const blob = new Blob([data], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ─── CSV 1行パース ──────────────────────────
function parseLine(line) {
  const cells = []; let i = 0;
  while (i < line.length) {
    if (line[i] === '"') {
      let v = ''; i++;
      while (i < line.length) {
        if (line[i] === '"' && line[i+1] === '"') { v += '"'; i += 2; }
        else if (line[i] === '"') { i++; break; }
        else { v += line[i++]; }
      }
      cells.push(v);
      if (line[i] === ',') i++;
    } else {
      const end = line.indexOf(',', i);
      if (end < 0) { cells.push(line.slice(i)); break; }
      cells.push(line.slice(i, end)); i = end + 1;
    }
  }
  return cells;
}

// ─── ユーティリティ ──────────────────────────
function setProcessing(on, msg) {
  processing.classList.toggle('hidden', !on);
  if (msg) procText.textContent = msg;
}
function showError(msg) { errorSpan.textContent = msg; errorBox.classList.remove('hidden'); }
function hideError()    { errorBox.classList.add('hidden'); }
function escH(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function stripExt(name) { return name.replace(/\.[^/.]+$/, ''); }
function sanitize(name) { return name.replace(/[/\\?%*:|"<>]/g, '_'); }


/* ═══════════════════════════════════════════════════════════
   VLOOKUP 結合機能
   ① 基本リスト（マスタ）と ② 挿入先リストを読み込み、
   キー列で照合して①の値を②に転記し CSV で出力する
   ═══════════════════════════════════════════════════════════ */

// ─── 状態 ──────────────────────────────────────────────────
let vlDataA = null;  // { headers:[], rows:[[]] } 基本リスト
let vlDataB = null;  // { headers:[], rows:[[]], fileName } 挿入先
let vlColMappings = []; // [{ fromCol, toCol }]

// ─── DOM ───────────────────────────────────────────────────
const vlDropA    = document.getElementById('vl-drop-a');
const vlDropB    = document.getElementById('vl-drop-b');
const vlInputA   = document.getElementById('vl-input-a');
const vlInputB   = document.getElementById('vl-input-b');
const vlFileA    = document.getElementById('vl-file-a');
const vlFileB    = document.getElementById('vl-file-b');
const vlMapping  = document.getElementById('vl-mapping');
const vlKeyA     = document.getElementById('vl-key-a');
const vlKeyB     = document.getElementById('vl-key-b');
const vlColList  = document.getElementById('vl-col-list');
const vlAddCol   = document.getElementById('vl-add-col');
const vlRunBtn   = document.getElementById('vl-run-btn');
const vlResult   = document.getElementById('vl-result');
const vlPreview  = document.getElementById('vl-preview-table');
const vlDlWrap   = document.getElementById('vl-dl-btn-wrap');
const vlResetBtn = document.getElementById('vl-reset-btn');

// ─── ドロップ＆クリック ────────────────────────────────────
function setupVlDrop(dropEl, inputEl, side) {
  dropEl.addEventListener('dragover',  e => { e.preventDefault(); dropEl.classList.add('drag-over'); });
  dropEl.addEventListener('dragleave', e => { if (!dropEl.contains(e.relatedTarget)) dropEl.classList.remove('drag-over'); });
  dropEl.addEventListener('drop', e => {
    e.preventDefault(); dropEl.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) loadVlFile(f, side);
  });
  dropEl.addEventListener('click', () => inputEl.click());
  dropEl.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputEl.click(); } });
  inputEl.addEventListener('change', e => {
    const f = e.target.files[0];
    if (f) loadVlFile(f, side);
    inputEl.value = '';
  });
}
setupVlDrop(vlDropA, vlInputA, 'A');
setupVlDrop(vlDropB, vlInputB, 'B');

// ─── Excel ファイル読み込み ────────────────────────────────
async function loadVlFile(file, side) {
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    alert('Excelファイル（.xlsx/.xls）を選択してください。');
    return;
  }
  try {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellText: true, cellDates: true, raw: false });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const data = parseSheetToTable(ws);
    data.fileName = file.name;
    data.baseName = stripExt(file.name);

    if (side === 'A') {
      vlDataA = data;
      renderVlFileInfo(vlFileA, file.name, data.headers.length, data.rows.length, 'a');
      vlDropA.classList.add('vl-loaded');
      // ①を読み込んだ直後にヘッダーをヒント表示
      renderVlHeaderHint('a', data.headers);
    } else {
      vlDataB = data;
      renderVlFileInfo(vlFileB, file.name, data.headers.length, data.rows.length, 'b');
      vlDropB.classList.add('vl-loaded');
      // ②を読み込んだ直後にヘッダーをヒント表示
      renderVlHeaderHint('b', data.headers);
    }

    // 両方読み込み済みならマッピングUIを表示
    if (vlDataA && vlDataB) buildMappingUI();
  } catch (e) {
    alert('ファイルの読み込みに失敗しました: ' + e.message);
  }
}

// シートをテーブル形式（headers + rows）に変換
function parseSheetToTable(ws) {
  if (!ws['!ref']) return { headers: [], rows: [] };
  const range = XLSX.utils.decode_range(ws['!ref']);
  const headers = [];
  const rows = [];

  // 1行目をヘッダーとして取得
  for (let C = range.s.c; C <= range.e.c; C++) {
    const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
    headers.push(cellText(cell));
  }

  // 2行目以降をデータとして取得
  for (let R = range.s.r + 1; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      row.push(cellText(cell));
    }
    rows.push(row);
  }
  return { headers, rows };
}

// ファイル情報表示
function renderVlFileInfo(el, name, cols, rows, side) {
  el.classList.remove('hidden');
  const color = side === 'a' ? '#6366f1' : '#0891b2';
  el.innerHTML = `
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/>
    </svg>
    <strong>${escH(name)}</strong>
    <span class="vl-file-meta">${cols} 列 / ${rows} 行</span>`;
}

// ─── ファイル読み込み直後にヘッダー一覧をヒント表示 ────────
function renderVlHeaderHint(side, headers) {
  const infoEl = side === 'a' ? vlFileA : vlFileB;
  // 既存の file-info に追記
  const existing = infoEl.querySelector('.vl-header-hint');
  if (existing) existing.remove();

  const hint = document.createElement('div');
  hint.className = 'vl-header-hint';
  hint.innerHTML = `<span class="vl-hint-label">列名：</span>` +
    headers.map(h => `<span class="vl-header-tag">${escH(h)}</span>`).join('');
  infoEl.appendChild(hint);

  // もう片方がまだなら「もう一方のファイルを読み込んでください」案内
  const waitEl = document.getElementById(`vl-wait-${side}`);
  if (!waitEl) {
    const wait = document.createElement('p');
    wait.id = `vl-wait-${side}`;
    wait.className = 'vl-wait-msg';
    const other = side === 'a' ? '②挿入先リスト' : '①基本リスト';
    const bothReady = vlDataA && vlDataB;
    if (!bothReady) {
      wait.textContent = `✅ 読み込み完了！次に ${other} を読み込んでください。`;
      infoEl.appendChild(wait);
    }
  }
}

// ─── マッピングUI を構築 ────────────────────────────────────
function buildMappingUI() {
  vlMapping.classList.remove('hidden');
  vlResult.classList.add('hidden');

  // 待機メッセージを削除
  document.querySelectorAll('.vl-wait-msg').forEach(el => el.remove());

  // キー列セレクト（①）--- ヘッダー名でoption生成
  vlKeyA.innerHTML = '<option value="">キー列を選択（① ' + escH(vlDataA.fileName) + '）</option>' +
    vlDataA.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('');

  // キー列セレクト（②）--- ヘッダー名でoption生成
  vlKeyB.innerHTML = '<option value="">キー列を選択（② ' + escH(vlDataB.fileName) + '）</option>' +
    vlDataB.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('');

  // 転記列リスト初期化（1行追加）
  vlColList.innerHTML = '';
  vlColMappings = [];
  addColMapping();

  checkRunnable();
}

// 転記列マッピング1行追加
function addColMapping() {
  const idx = vlColMappings.length;
  vlColMappings.push({ fromCol: '', toCol: '' });

  const row = document.createElement('div');
  row.className = 'vl-col-row';
  row.dataset.idx = idx;
  row.innerHTML = `
    <div class="vl-select-wrap">
      <span class="vl-select-badge badge-a">①</span>
      <select class="vl-select vl-from-col" data-idx="${idx}">
        <option value="">転記元の列（①）</option>
        ${vlDataA.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('')}
      </select>
    </div>
    <div class="vl-arrow-small">→</div>
    <div class="vl-select-wrap">
      <span class="vl-select-badge badge-b">②</span>
      <select class="vl-select vl-to-col" data-idx="${idx}">
        <option value="">転記先の列（②）</option>
        ${vlDataB.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('')}
      </select>
    </div>
    <button class="vl-del-col" data-idx="${idx}" title="この行を削除">✕</button>
  `;

  row.querySelector('.vl-from-col').addEventListener('change', e => {
    vlColMappings[+e.target.dataset.idx].fromCol = e.target.value;
    checkRunnable();
  });
  row.querySelector('.vl-to-col').addEventListener('change', e => {
    vlColMappings[+e.target.dataset.idx].toCol = e.target.value;
    checkRunnable();
  });
  row.querySelector('.vl-del-col').addEventListener('click', () => {
    row.remove();
    checkRunnable();
  });

  vlColList.appendChild(row);
}

vlAddCol.addEventListener('click', addColMapping);

vlKeyA.addEventListener('change', checkRunnable);
vlKeyB.addEventListener('change', checkRunnable);

function checkRunnable() {
  // キーが両方選択済み＆転記列が1つ以上有効
  const keyOk = vlKeyA.value !== '' && vlKeyB.value !== '';
  const colRows = vlColList.querySelectorAll('.vl-col-row');
  let colOk = false;
  colRows.forEach(row => {
    const from = row.querySelector('.vl-from-col').value;
    const to   = row.querySelector('.vl-to-col').value;
    if (from !== '' && to !== '') colOk = true;
  });
  vlRunBtn.disabled = !(keyOk && colOk);
}

// ─── 転記実行 ─────────────────────────────────────────────
vlRunBtn.addEventListener('click', runVlookup);

function runVlookup() {
  const keyAIdx = +vlKeyA.value;
  const keyBIdx = +vlKeyB.value;

  // 転記列マッピングを収集
  const colRows = vlColList.querySelectorAll('.vl-col-row');
  const mappings = [];
  colRows.forEach(row => {
    const from = row.querySelector('.vl-from-col').value;
    const to   = row.querySelector('.vl-to-col').value;
    if (from !== '' && to !== '') mappings.push({ from: +from, to: +to });
  });

  if (!mappings.length) { alert('転記列を1つ以上設定してください。'); return; }

  // ①のキー → 行 のマップを作成
  // ※キーは trim() のみ。先頭ゼロを落とさないよう数値変換しない
  const masterMap = new Map();
  vlDataA.rows.forEach(row => {
    const key = String(row[keyAIdx] ?? '').trim();
    if (key !== '') masterMap.set(key, row);
  });

  // ②のデータをコピーして転記
  const newRows = vlDataB.rows.map(row => {
    const newRow = [...row];
    const key = String(row[keyBIdx] ?? '').trim();
    const masterRow = masterMap.get(key);
    if (masterRow) {
      mappings.forEach(({ from, to }) => {
        // 転記元の値をそのまま代入（先頭ゼロ保持）
        newRow[to] = masterRow[from] ?? '';
      });
    }
    return newRow;
  });

  // プレビュー表示
  const allRows = [vlDataB.headers, ...newRows];
  const previewRows = allRows.slice(0, 21);
  const trs = previewRows.map((row, ri) => {
    const tds = row.map(c => `<td>${escH(c)}</td>`).join('');
    return `<tr class="${ri === 0 ? 'preview-header-row' : ''}">${tds}</tr>`;
  });
  vlPreview.innerHTML = `<table class="preview-table"><tbody>${trs.join('')}</tbody></table>`;

  // ── CSV生成 ──────────────────────────────────────────────
  // ・ヘッダー行も含めて全セルをダブルクォートで囲む
  // ・セル内の " は "" にエスケープ
  // ・数値変換は一切しない（先頭ゼロ保持）
  const csvLines = [vlDataB.headers, ...newRows].map(row =>
    row.map(v => {
      const str = String(v ?? '');           // 必ず文字列化（数値変換しない）
      return `"${str.replace(/"/g, '""')}"`;  // ダブルクォートで囲む
    }).join(',')
  );
  const csvText = csvLines.join('\r\n');

  // ダウンロードボタン
  const enc = optEncoding();
  const csvName = `${vlDataB.baseName}_転記済み.csv`;
  vlDlWrap.innerHTML = '';
  const dlBtn = document.createElement('button');
  dlBtn.className = 'download-btn';
  dlBtn.innerHTML = `
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="7 10 12 15 17 10"/>
      <line x1="12" y1="15" x2="12" y2="3"/>
    </svg>
    ${escH(csvName)} をダウンロード
  `;
  dlBtn.addEventListener('click', () => dlCsv(csvText, csvName, enc));
  vlDlWrap.appendChild(dlBtn);

  vlResult.classList.remove('hidden');
  vlResult.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── リセット ─────────────────────────────────────────────
vlResetBtn.addEventListener('click', () => {
  vlDataA = null; vlDataB = null;
  vlFileA.classList.add('hidden'); vlFileB.classList.add('hidden');
  vlDropA.classList.remove('vl-loaded'); vlDropB.classList.remove('vl-loaded');
  vlMapping.classList.add('hidden');
  vlResult.classList.add('hidden');
  vlColList.innerHTML = '';
  vlColMappings = [];
  // ヒント・待機メッセージ削除
  document.querySelectorAll('.vl-header-hint, .vl-wait-msg').forEach(el => el.remove());
});
