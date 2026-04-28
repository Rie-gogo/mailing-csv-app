/**
 * gaiji.js — 外字チェック・変換機能
 * =====================================
 * 外字（Unicode私用領域 / Private Use Area）の検出と変換を行う。
 *
 * 【外字の判定根拠】
 *   参考: https://bunkyudo.co.jp/external-fonts-excel-powerquery-t-h/
 *   Unicode BMP 私用領域: U+E000 ～ U+F8FF (10進数: 57344 ～ 63743)
 *   補助私用領域A:        U+F0000 ～ U+FFFFF
 *   補助私用領域B:        U+100000 ～ U+10FFFF
 *
 * 【依存】
 *   - SheetJS (XLSX) — Excelファイルの読み込みに使用
 *   - app.js の cellText() / escH() / dlCsv() / stripExt() を共有
 */
'use strict';

/* ═══════════════════════════════════════════════════════════
   外字判定ロジック（コアモジュール）
   ─ この部分は他のアプリへの移植が容易なよう独立させています ─
   ═══════════════════════════════════════════════════════════ */

/**
 * Unicode 私用領域の範囲定義
 * @type {Array<[number, number]>}
 */
const GAIJI_RANGES = [
  [0xE000,   0xF8FF  ],  // BMP 私用領域 (57344 ～ 63743)
  [0xF0000,  0xFFFFF ],  // 補助私用領域A
  [0x100000, 0x10FFFF],  // 補助私用領域B
];

/**
 * 1文字が外字（Unicode私用領域）かどうかを判定する。
 * @param {string} char - 判定する1文字
 * @returns {boolean}
 */
function isGaiji(char) {
  if (!char) return false;
  const cp = char.codePointAt(0);
  return GAIJI_RANGES.some(([s, e]) => cp >= s && cp <= e);
}

/**
 * 文字列中の外字をすべて検出する。
 * @param {string} text - チェック対象の文字列
 * @returns {{ hasGaiji: boolean, chars: string[], positions: number[], codepoints: string[] }}
 */
function checkString(text) {
  const result = { hasGaiji: false, chars: [], positions: [], codepoints: [] };
  // サロゲートペアを正しく扱うため Array.from を使用
  const chars = Array.from(text);
  chars.forEach((ch, idx) => {
    if (isGaiji(ch)) {
      result.hasGaiji = true;
      result.chars.push(ch);
      result.positions.push(idx);
      result.codepoints.push(`U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`);
    }
  });
  return result;
}

/**
 * 変換テーブルを使って文字列中の外字を変換する。
 * @param {string} text - 変換対象の文字列
 * @param {Map<string, string>} convTable - 外字→変換先のMap
 * @param {string} fallback - テーブルに存在しない外字の置き換え先
 * @returns {string}
 */
function convertGaiji(text, convTable, fallback = '?') {
  return Array.from(text).map(ch => {
    if (!isGaiji(ch)) return ch;
    return convTable.has(ch) ? convTable.get(ch) : fallback;
  }).join('');
}

/* ═══════════════════════════════════════════════════════════
   UI / アプリ連携部分
   ═══════════════════════════════════════════════════════════ */

// ─── 状態 ──────────────────────────────────────────────────
let gaijiExcelData = null;  // { fileName, baseName, sheets: [{ sheetName, rows: [[]] }] }
let gaijiConvTable = new Map(); // 外字 → 変換先
let gaijiConvLoaded = false;    // 変換テーブルCSVが読み込まれているか

// ─── DOM ───────────────────────────────────────────────────
const gDropExcel   = document.getElementById('gaiji-drop-excel');
const gInputExcel  = document.getElementById('gaiji-input-excel');
const gFileExcel   = document.getElementById('gaiji-file-excel');
const gDropCsv     = document.getElementById('gaiji-drop-csv');
const gInputCsv    = document.getElementById('gaiji-input-csv');
const gFileCsv     = document.getElementById('gaiji-file-csv');
const gRunBtn      = document.getElementById('gaiji-run-btn');
const gProcessing  = document.getElementById('gaiji-processing');
const gProcText    = document.getElementById('gaiji-proc-text');
const gResult      = document.getElementById('gaiji-result');
const gSummaryBanner = document.getElementById('gaiji-summary-banner');
const gSheetTabs   = document.getElementById('gaiji-sheet-tabs');
const gDetectNote  = document.getElementById('gaiji-detect-note');
const gDetectTable = document.getElementById('gaiji-detect-table');
const gDlWrap      = document.getElementById('gaiji-dl-wrap');
const gResetBtn    = document.getElementById('gaiji-reset-btn');

// オプション
const gOptAllSheets = () => document.getElementById('gaiji-opt-all-sheets').checked;
const gOptConvert   = () => document.getElementById('gaiji-opt-convert').checked;
const gOptFallback  = () => document.getElementById('gaiji-opt-fallback').value;

// ─── ドロップ＆クリック（Excel） ──────────────────────────
gDropExcel.addEventListener('dragover',  e => { e.preventDefault(); gDropExcel.classList.add('drag-over'); });
gDropExcel.addEventListener('dragleave', e => { if (!gDropExcel.contains(e.relatedTarget)) gDropExcel.classList.remove('drag-over'); });
gDropExcel.addEventListener('drop', e => {
  e.preventDefault(); gDropExcel.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) loadGaijiExcel(f);
});
gDropExcel.addEventListener('click', () => gInputExcel.click());
gDropExcel.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); gInputExcel.click(); } });
gInputExcel.addEventListener('change', e => {
  const f = e.target.files[0];
  if (f) loadGaijiExcel(f);
  gInputExcel.value = '';
});

// ─── ドロップ＆クリック（変換テーブルCSV） ───────────────
gDropCsv.addEventListener('dragover',  e => { e.preventDefault(); gDropCsv.classList.add('drag-over'); });
gDropCsv.addEventListener('dragleave', e => { if (!gDropCsv.contains(e.relatedTarget)) gDropCsv.classList.remove('drag-over'); });
gDropCsv.addEventListener('drop', e => {
  e.preventDefault(); gDropCsv.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) loadGaijiCsv(f);
});
gDropCsv.addEventListener('click', () => gInputCsv.click());
gDropCsv.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); gInputCsv.click(); } });
gInputCsv.addEventListener('change', e => {
  const f = e.target.files[0];
  if (f) loadGaijiCsv(f);
  gInputCsv.value = '';
});

// ─── Excel ファイル読み込み ────────────────────────────────
async function loadGaijiExcel(file) {
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    alert('Excelファイル（.xlsx/.xls）を選択してください。');
    return;
  }
  try {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellText: true, cellDates: true, raw: false });

    const targetSheets = gOptAllSheets() ? wb.SheetNames : [wb.SheetNames[0]];
    const sheets = targetSheets.map(name => {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) return { sheetName: name, rows: [] };
      const range = XLSX.utils.decode_range(ws['!ref']);
      const rows = [];
      for (let R = range.s.r; R <= range.e.r; R++) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
          const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
          row.push(cellText(cell));
        }
        rows.push(row);
      }
      return { sheetName: name, rows };
    });

    gaijiExcelData = {
      fileName: file.name,
      baseName: stripExt(file.name),
      sheets,
    };

    // ファイル情報表示
    const totalRows = sheets.reduce((s, sh) => s + sh.rows.length, 0);
    gFileExcel.classList.remove('hidden');
    gFileExcel.innerHTML = `
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
      <strong>${escH(file.name)}</strong>
      <span class="vl-file-meta">${sheets.length} シート / ${totalRows} 行</span>`;
    gDropExcel.classList.add('vl-loaded');

    checkGaijiRunnable();
  } catch (e) {
    alert('ファイルの読み込みに失敗しました: ' + e.message);
  }
}

// ─── 変換テーブルCSV 読み込み ──────────────────────────────
async function loadGaijiCsv(file) {
  if (!/\.csv$/i.test(file.name)) {
    alert('CSVファイル（.csv）を選択してください。');
    return;
  }
  try {
    const text = await file.text();
    gaijiConvTable = parseConversionCsv(text);
    gaijiConvLoaded = true;

    gFileCsv.classList.remove('hidden');
    gFileCsv.innerHTML = `
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
      <strong>${escH(file.name)}</strong>
      <span class="vl-file-meta">${gaijiConvTable.size} 件登録</span>`;
    gDropCsv.classList.add('vl-loaded');

    checkGaijiRunnable();
  } catch (e) {
    alert('CSVの読み込みに失敗しました: ' + e.message);
  }
}

/**
 * 変換テーブルCSVをパースして Map に変換する。
 * フォーマット: 外字文字,変換先文字列（ヘッダーなし）
 * @param {string} csvText
 * @returns {Map<string, string>}
 */
function parseConversionCsv(csvText) {
  const map = new Map();
  const lines = csvText.split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    // カンマで最初の区切りのみ分割（変換先にカンマが含まれる場合を考慮）
    const idx = trimmed.indexOf(',');
    if (idx < 0) continue;
    const key = trimmed.slice(0, idx);
    const val = trimmed.slice(idx + 1);
    if (key) map.set(key, val);
  }
  return map;
}

// ─── 実行ボタンの活性化制御 ────────────────────────────────
function checkGaijiRunnable() {
  gRunBtn.disabled = !gaijiExcelData;
}

// ─── 実行ボタン ────────────────────────────────────────────
gRunBtn.addEventListener('click', runGaijiCheck);

async function runGaijiCheck() {
  if (!gaijiExcelData) return;

  gProcessing.classList.remove('hidden');
  gProcText.textContent = 'チェック中...';
  gResult.classList.add('hidden');

  // 非同期処理のため少し待機してUIを更新
  await new Promise(r => setTimeout(r, 30));

  try {
    const fallback = gOptFallback();
    const doConvert = gOptConvert();

    // 全シートの外字チェック結果を収集
    // sheetResults: [{ sheetName, detections: [{ row, col, value, gaijiChars, codepoints }], convertedRows }]
    const sheetResults = gaijiExcelData.sheets.map(sheet => {
      const detections = [];
      const convertedRows = [];

      sheet.rows.forEach((row, rowIdx) => {
        const convertedRow = [];
        row.forEach((cell, colIdx) => {
          const checked = checkString(cell);
          if (checked.hasGaiji) {
            detections.push({
              row: rowIdx + 1,
              col: colIdx + 1,
              value: cell,
              gaijiChars: checked.chars,
              codepoints: checked.codepoints,
            });
          }
          if (doConvert) {
            convertedRow.push(convertGaiji(cell, gaijiConvTable, fallback));
          }
        });
        if (doConvert) convertedRows.push(convertedRow);
      });

      return { sheetName: sheet.sheetName, detections, convertedRows };
    });

    gProcessing.classList.add('hidden');
    renderGaijiResult(sheetResults, doConvert, fallback);
  } catch (e) {
    gProcessing.classList.add('hidden');
    alert('チェック中にエラーが発生しました: ' + e.message);
  }
}

// ─── 結果描画 ──────────────────────────────────────────────
let gCurrentSheetIdx = 0;
let gLastSheetResults = null;
let gLastDoConvert = false;
let gLastFallback = '?';

function renderGaijiResult(sheetResults, doConvert, fallback) {
  gLastSheetResults = sheetResults;
  gLastDoConvert = doConvert;
  gLastFallback = fallback;
  gCurrentSheetIdx = 0;

  const totalDetections = sheetResults.reduce((s, r) => s + r.detections.length, 0);
  const hasAny = totalDetections > 0;

  // サマリーバナー
  if (hasAny) {
    gSummaryBanner.className = 'gaiji-summary-banner gaiji-banner-warn';
    gSummaryBanner.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
        <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
      </svg>
      <span>外字が <strong>${totalDetections}</strong> 件検出されました。${doConvert ? '変換済みCSVをダウンロードできます。' : ''}</span>`;
  } else {
    gSummaryBanner.className = 'gaiji-summary-banner gaiji-banner-ok';
    gSummaryBanner.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <polyline points="20 6 9 17 4 12"/>
      </svg>
      <span>外字は検出されませんでした。${doConvert ? '変換済みCSVをダウンロードできます。' : ''}</span>`;
  }

  // シートタブ（複数シート時）
  if (sheetResults.length > 1) {
    gSheetTabs.innerHTML = sheetResults.map((r, i) => {
      const badge = r.detections.length > 0 ? ` <span class="gaiji-tab-badge">${r.detections.length}</span>` : '';
      return `<button class="sheet-tab ${i === 0 ? 'active' : ''}" data-i="${i}">${escH(r.sheetName)}${badge}</button>`;
    }).join('');
    gSheetTabs.querySelectorAll('.sheet-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        gCurrentSheetIdx = +btn.dataset.i;
        gSheetTabs.querySelectorAll('.sheet-tab').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        renderGaijiDetectTable(sheetResults[gCurrentSheetIdx]);
      });
    });
    gSheetTabs.classList.remove('hidden');
  } else {
    gSheetTabs.innerHTML = '';
    gSheetTabs.classList.add('hidden');
  }

  renderGaijiDetectTable(sheetResults[0]);
  renderGaijiDownloadButtons(sheetResults, doConvert);

  gResult.classList.remove('hidden');
  gResult.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/**
 * 外字検出結果テーブルを描画する。
 */
function renderGaijiDetectTable(sheetResult) {
  const { detections } = sheetResult;

  if (!detections.length) {
    gDetectNote.textContent = '外字なし';
    gDetectTable.innerHTML = '<p class="empty-preview" style="color:#16a34a">このシートに外字は含まれていません</p>';
    return;
  }

  gDetectNote.textContent = `${detections.length} 件の外字を検出`;

  const rows = detections.map(d => {
    const cpList = d.codepoints.join(', ');
    const charList = d.gaijiChars.map(c =>
      `<span class="gaiji-char-badge" title="${escH(GaijiChecker_cpStr(c))}">${escH(GaijiChecker_cpStr(c))}</span>`
    ).join(' ');
    return `<tr>
      <td class="gaiji-td-num">${d.row}</td>
      <td class="gaiji-td-num">${d.col}</td>
      <td class="gaiji-td-val">${escH(d.value)}</td>
      <td class="gaiji-td-cp">${charList}</td>
      <td class="gaiji-td-cp">${escH(cpList)}</td>
    </tr>`;
  });

  gDetectTable.innerHTML = `
    <table class="preview-table gaiji-detect-tbl">
      <thead>
        <tr>
          <th>行</th><th>列</th><th>セルの値</th><th>外字文字</th><th>コードポイント</th>
        </tr>
      </thead>
      <tbody>${rows.join('')}</tbody>
    </table>`;
}

/** コードポイント文字列を返す */
function GaijiChecker_cpStr(char) {
  return `U+${char.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`;
}

/**
 * ダウンロードボタンを描画する。
 */
function renderGaijiDownloadButtons(sheetResults, doConvert) {
  gDlWrap.innerHTML = '';
  const enc = document.getElementById('opt-encoding').value;

  // 検出結果CSVのダウンロード
  const allDetections = sheetResults.flatMap(r =>
    r.detections.map(d => ({ sheet: r.sheetName, ...d }))
  );

  if (allDetections.length > 0) {
    const reportCsv = buildDetectionReportCsv(allDetections);
    const reportName = `${gaijiExcelData.baseName}_外字検出レポート.csv`;
    const reportBtn = document.createElement('button');
    reportBtn.className = 'download-btn gaiji-dl-btn';
    reportBtn.innerHTML = `
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
        <polyline points="7 10 12 15 17 10"/>
        <line x1="12" y1="15" x2="12" y2="3"/>
      </svg>
      外字検出レポート（${allDetections.length}件）をダウンロード`;
    reportBtn.addEventListener('click', () => dlCsv(reportCsv, reportName, enc));
    gDlWrap.appendChild(reportBtn);
  }

  // 変換済みCSVのダウンロード
  if (doConvert) {
    sheetResults.forEach(r => {
      if (!r.convertedRows.length) return;
      const csvText = r.convertedRows.map(row =>
        row.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')
      ).join('\r\n');

      const csvName = sheetResults.length > 1
        ? `${gaijiExcelData.baseName}_${sanitize(r.sheetName)}_外字変換済み.csv`
        : `${gaijiExcelData.baseName}_外字変換済み.csv`;

      const btn = document.createElement('button');
      btn.className = 'download-btn gaiji-dl-btn gaiji-dl-btn-convert';
      btn.innerHTML = `
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/>
          <line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        ${escH(csvName)} をダウンロード`;
      btn.addEventListener('click', () => dlCsv(csvText, csvName, enc));
      gDlWrap.appendChild(btn);
    });
  }

  // 何もない場合
  if (!gDlWrap.children.length) {
    gDlWrap.innerHTML = '<p class="gaiji-no-dl">外字が検出されなかったため、ダウンロードファイルはありません。</p>';
  }
}

/**
 * 外字検出レポートCSVを生成する。
 */
function buildDetectionReportCsv(detections) {
  const header = '"シート名","行","列","セルの値","外字コードポイント"';
  const rows = detections.map(d =>
    [d.sheet, d.row, d.col, d.value, d.codepoints.join(' ')].map(v =>
      `"${String(v ?? '').replace(/"/g, '""')}"`
    ).join(',')
  );
  return [header, ...rows].join('\r\n');
}

// ─── リセット ──────────────────────────────────────────────
gResetBtn.addEventListener('click', resetGaiji);

function resetGaiji() {
  gaijiExcelData = null;
  gaijiConvTable = new Map();
  gaijiConvLoaded = false;

  gFileExcel.classList.add('hidden');
  gFileCsv.classList.add('hidden');
  gDropExcel.classList.remove('vl-loaded');
  gDropCsv.classList.remove('vl-loaded');
  gResult.classList.add('hidden');
  gRunBtn.disabled = true;
  gDetectTable.innerHTML = '';
  gDlWrap.innerHTML = '';
}
