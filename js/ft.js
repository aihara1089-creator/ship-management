// ============================================================
// ft.js — FT搭載船リスト専用モジュール
// FleetTransfer管理表_MOL殿向け.xlsx を読み込んで表示
// ============================================================

'use strict';

// ============================================================
// STATE
// ============================================================
let ftAllRows   = [];   // 全行データ
let ftFiltered  = [];   // フィルター後
let ftOsFilter  = 'all';
let ftSearchQ   = '';
let ftPage      = 1;
const FT_PAGE_SIZE = 50;

// ============================================================
// EXCEL PARSER
// ============================================================
function parseFtExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  // "FT管理表" シートを優先、なければ先頭シート
  const targetSheet = workbook.SheetNames.find(n => n.includes('FT') || n.includes('管理')) 
                   || workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheet];

  // 日付型セルを yyyy/mm/dd に変換
  Object.keys(sheet).filter(k => !k.startsWith('!')).forEach(k => {
    const c = sheet[k];
    if (c && c.t === 'd' && c.v instanceof Date) {
      const d = c.v;
      const yyyy = d.getFullYear();
      const mm   = String(d.getMonth() + 1).padStart(2, '0');
      const dd   = String(d.getDate()).padStart(2, '0');
      c.t = 's';
      c.v = `${yyyy}/${mm}/${dd}`;
      c.w = c.v;
    }
  });

  const jsonRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false,
  });

  const rows = [];
  let currentOs = '';   // ▼ セクションで判断するOSデフォルト
  let no = 1;

  for (let i = 0; i < jsonRows.length; i++) {
    const row = jsonRows[i];
    const col0 = String(row[0] || '').trim();
    const col1 = String(row[1] || '').trim(); // Serial No.
    const col2 = String(row[2] || '').trim(); // OS
    const col3 = String(row[3] || '').trim(); // FT搭載日
    const col4 = String(row[4] || '').trim(); // 備考

    // ヘッダー行・タイトル行はスキップ
    if (!col0 || col0.startsWith('Fleet Transfer') || col0 === '船名 (Vessel Name)') continue;

    // ▼ セクション行（OSカテゴリ切替）
    if (col0.startsWith('▼')) {
      if (col0.includes('7')) currentOs = 'Windows7';
      else if (col0.includes('10')) currentOs = 'Windows10';
      else if (col0.includes('11')) currentOs = 'Windows11';
      continue;
    }

    // データ行
    const os = col2 || currentOs;
    if (!col0) continue; // 船名なしはスキップ

    // 搭載日の判定（予定かどうか）
    const isScheduled = !col3 || col4.includes('予定') || col3.includes('予定')
                     || /\d{4}\/\d{1,2}\/\d{1,2}-\d{1,2}/.test(col3)
                     || /^\d{4}\/\d{1,2}-\d{1,2}/.test(col3);

    rows.push({
      no:          no++,
      vesselName:  col0,
      serialNo:    col1,
      os:          os,
      installDate: col3,
      note:        col4,
      isScheduled: isScheduled,
    });
  }

  return rows;
}

// ============================================================
// KPI
// ============================================================
function renderFtKpi(rows) {
  const total     = rows.length;
  const win7      = rows.filter(r => r.os === 'Windows7').length;
  const win10     = rows.filter(r => r.os === 'Windows10' || r.os === 'Windows11').length;
  const scheduled = rows.filter(r => r.isScheduled).length;

  document.getElementById('ftKpiTotal').textContent     = total.toLocaleString();
  document.getElementById('ftKpiWin7').textContent      = win7.toLocaleString();
  document.getElementById('ftKpiWin10').textContent     = win10.toLocaleString();
  document.getElementById('ftKpiScheduled').textContent = scheduled.toLocaleString();
}

// ============================================================
// OS バッジ
// ============================================================
function ftOsBadge(os) {
  if (os === 'Windows7')  return `<span class="ft-os-badge ft-os-win7"><i class="fas fa-windows"></i> Win 7</span>`;
  if (os === 'Windows10') return `<span class="ft-os-badge ft-os-win10"><i class="fas fa-windows"></i> Win 10</span>`;
  if (os === 'Windows11') return `<span class="ft-os-badge ft-os-win11"><i class="fas fa-windows"></i> Win 11</span>`;
  return `<span class="ft-os-badge">${os}</span>`;
}

// ============================================================
// TABLE
// ============================================================
function renderFtTable() {
  const start = (ftPage - 1) * FT_PAGE_SIZE;
  const pageRows = ftFiltered.slice(start, start + FT_PAGE_SIZE);

  const tbody = document.getElementById('ftTableBody');
  if (!tbody) return;

  if (ftFiltered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:40px;color:var(--slate-400)">
      <i class="fas fa-search" style="font-size:2rem;display:block;margin-bottom:8px"></i>
      該当データなし</td></tr>`;
    document.getElementById('ftTableCount').textContent = '0 件表示';
    renderFtPagination();
    return;
  }

  tbody.innerHTML = pageRows.map(r => {
    const scheduledBadge = r.isScheduled
      ? `<span class="ft-scheduled-badge"><i class="fas fa-clock"></i> 予定</span>`
      : '';
    const dateCell = r.installDate
      ? `${r.installDate}${scheduledBadge}`
      : `<span class="ft-no-date">未定</span>`;
    const rowCls = r.isScheduled ? 'ft-row-scheduled' : '';

    return `<tr class="${rowCls}">
      <td class="ft-td-no">${r.no}</td>
      <td class="ft-td-name">${escFt(r.vesselName)}</td>
      <td class="ft-td-serial">${r.serialNo ? escFt(r.serialNo) : '<span style="color:var(--slate-300)">—</span>'}</td>
      <td class="ft-td-os">${ftOsBadge(r.os)}</td>
      <td class="ft-td-date">${dateCell}</td>
      <td class="ft-td-note">${r.note ? escFt(r.note) : ''}</td>
    </tr>`;
  }).join('');

  const showing = Math.min(start + FT_PAGE_SIZE, ftFiltered.length);
  document.getElementById('ftTableCount').textContent =
    `${start + 1}–${showing} / ${ftFiltered.length} 件`;

  renderFtPagination();
}

function escFt(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ============================================================
// PAGINATION
// ============================================================
function renderFtPagination() {
  const container = document.getElementById('ftPagination');
  if (!container) return;
  const totalPages = Math.ceil(ftFiltered.length / FT_PAGE_SIZE);
  if (totalPages <= 1) { container.innerHTML = ''; return; }

  let html = '';
  // 前へ
  html += `<button class="ft-page-btn" ${ftPage === 1 ? 'disabled' : ''} data-p="${ftPage - 1}">
    <i class="fas fa-chevron-left"></i></button>`;

  // ページ番号
  const range = 2;
  for (let p = 1; p <= totalPages; p++) {
    if (p === 1 || p === totalPages || Math.abs(p - ftPage) <= range) {
      html += `<button class="ft-page-btn ${p === ftPage ? 'active' : ''}" data-p="${p}">${p}</button>`;
    } else if (Math.abs(p - ftPage) === range + 1) {
      html += `<span class="ft-page-ellipsis">…</span>`;
    }
  }

  // 次へ
  html += `<button class="ft-page-btn" ${ftPage === totalPages ? 'disabled' : ''} data-p="${ftPage + 1}">
    <i class="fas fa-chevron-right"></i></button>`;

  container.innerHTML = html;
  container.querySelectorAll('.ft-page-btn[data-p]').forEach(btn => {
    btn.addEventListener('click', () => {
      ftPage = Number(btn.dataset.p);
      renderFtTable();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });
}

// ============================================================
// FILTER & SEARCH
// ============================================================
function applyFtFilters() {
  const q = ftSearchQ.toLowerCase();

  ftFiltered = ftAllRows.filter(r => {
    // OSフィルター
    if (ftOsFilter === 'Windows7'  && r.os !== 'Windows7')  return false;
    if (ftOsFilter === 'Windows10' && r.os !== 'Windows10' && r.os !== 'Windows11') return false;

    // テキスト検索
    if (q && !r.vesselName.toLowerCase().includes(q) && !r.serialNo.toLowerCase().includes(q)) return false;

    return true;
  });

  ftPage = 1;
  renderFtTable();
}

// ============================================================
// CSV EXPORT
// ============================================================
function exportFtCsv() {
  const header = ['No.', '船名', 'Serial No.', 'OS', 'FT搭載日', '備考'].join(',');
  const rows = ftFiltered.map(r =>
    [r.no, `"${r.vesselName.replace(/"/g,'""')}"`,
     `"${r.serialNo.replace(/"/g,'""')}"`,
     r.os,
     `"${r.installDate.replace(/"/g,'""')}"`,
     `"${r.note.replace(/"/g,'""')}"`
    ].join(',')
  );
  const blob = new Blob(['\uFEFF' + header + '\n' + rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `FT搭載船リスト_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ============================================================
// SHOW / HIDE
// ============================================================
function showFtDashboard(rows) {
  ftAllRows  = rows;
  ftFiltered = rows;
  ftPage     = 1;
  ftOsFilter = 'all';
  ftSearchQ  = '';

  // フィルターボタンをリセット
  document.querySelectorAll('#ftOsFilter .ft-filter-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.os === 'all');
  });
  document.getElementById('ftSearchInput').value = '';

  renderFtKpi(rows);
  renderFtTable();

  document.getElementById('ftUploadSection').classList.add('hidden');
  document.getElementById('ftDashboard').classList.remove('hidden');
}

function showFtUpload() {
  document.getElementById('ftDashboard').classList.add('hidden');
  document.getElementById('ftUploadSection').classList.remove('hidden');
  ftAllRows  = [];
  ftFiltered = [];
}

// ============================================================
// PAGE NAV（新造船スケジュール ⇔ FT搭載船リスト）
// ============================================================
function initPageNav() {
  const shiplistPage = document.getElementById('uploadSection')?.parentElement
                    || document.querySelector('.upload-section')?.parentElement;

  document.querySelectorAll('.page-nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const page = btn.dataset.page;

      // アクティブ切替
      document.querySelectorAll('.page-nav-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      if (page === 'ft') {
        // FTページを表示
        document.getElementById('uploadSection').classList.add('hidden');
        document.getElementById('dashboard').classList.add('hidden');
        document.getElementById('ftPage').classList.remove('hidden');
      } else {
        // 新造船スケジュールページを表示
        document.getElementById('ftPage').classList.add('hidden');
        // 既存のdashboard/uploadSectionの表示状態は app.js が管理しているのでそのまま復元
        const hasDashData = document.getElementById('dashboard') && 
                            !document.getElementById('dashboard').classList.contains('hidden-always');
        // app.jsのallDataが空かどうかで判断できないので、headerMetaのtotalCountで判断
        const totalEl = document.getElementById('totalCount');
        const hasData = totalEl && !totalEl.textContent.includes('—');
        if (hasData) {
          document.getElementById('dashboard').classList.remove('hidden');
        } else {
          document.getElementById('uploadSection').classList.remove('hidden');
        }
      }
    });
  });
}

// ============================================================
// FILE DROP / SELECT
// ============================================================
function initFtFileHandlers() {
  const dropZone   = document.getElementById('ftDropZone');
  const fileInput  = document.getElementById('ftFileInput');

  if (!dropZone || !fileInput) return;

  // ドラッグ＆ドロップ
  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) loadFtFile(file);
  });

  // クリック選択
  fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) loadFtFile(file);
  });

  // 「別ファイル」ボタン
  const btnBack = document.getElementById('ftBtnBack');
  if (btnBack) btnBack.addEventListener('click', showFtUpload);

  // CSV出力
  const btnExport = document.getElementById('ftBtnExport');
  if (btnExport) btnExport.addEventListener('click', exportFtCsv);

  // OSフィルターボタン
  document.querySelectorAll('#ftOsFilter .ft-filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#ftOsFilter .ft-filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      ftOsFilter = btn.dataset.os;
      applyFtFilters();
    });
  });

  // 検索
  const searchInput = document.getElementById('ftSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      ftSearchQ = searchInput.value.trim();
      applyFtFilters();
    });
  }
}

function loadFtFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const rows = parseFtExcel(new Uint8Array(e.target.result));
      if (rows.length === 0) {
        alert('データが読み込めませんでした。ファイルの形式を確認してください。');
        return;
      }
      showFtDashboard(rows);
    } catch (err) {
      console.error('FT Excel parse error:', err);
      alert('ファイルの読み込みに失敗しました: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  initPageNav();
  initFtFileHandlers();
});
