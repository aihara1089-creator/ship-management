// ============================================================
// ft.js — FT搭載船リスト専用モジュール
// FleetTransfer管理表_MOL殿向け.xlsx を読み込んで表示・編集
// ============================================================

'use strict';

// ============================================================
// CONSTANTS
// ============================================================
const FT_PAGE_SIZE      = 50;
const FT_STORE_KEY      = 'molFtEdit_v1';  // localStorage キー

// ============================================================
// STATE
// ============================================================
let ftAllRows   = [];   // 全行データ（Excel + 受注連動 auto 行）
let ftFiltered  = [];   // フィルター後
let ftOsFilter  = 'all';
let ftSourceFilter = 'all'; // 'all' | 'excel' | 'auto' | 'manual'
let ftSearchQ   = '';
let ftCurrentPage = 1;

// localStorage に保存する編集データ
// shape: { [normalizedVesselName]: { serialNo, installDate, os, note, editedAt } }
let ftEditStore = {};

// ============================================================
// PERSISTENCE
// ============================================================
function ftSaveStore() {
  try { localStorage.setItem(FT_STORE_KEY, JSON.stringify(ftEditStore)); } catch(e) {}
}
function ftLoadStore() {
  try {
    const raw = localStorage.getItem(FT_STORE_KEY);
    if (raw) ftEditStore = JSON.parse(raw);
  } catch(e) { ftEditStore = {}; }
}
// 船名を正規化してキーにする（大文字・スペース統一）
function ftNormalizeKey(name) {
  return String(name || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

// ============================================================
// 編集データの適用（Excelデータに編集オーバーレイを重ねる）
// ============================================================
function ftApplyEdits(rows) {
  return rows.map(r => {
    const key  = ftNormalizeKey(r.vesselName);
    const edit = ftEditStore[key];
    if (!edit) return r;
    return {
      ...r,
      serialNo:    edit.serialNo    !== undefined ? edit.serialNo    : r.serialNo,
      installDate: edit.installDate !== undefined ? edit.installDate : r.installDate,
      os:          edit.os          !== undefined ? edit.os          : r.os,
      note:        edit.note        !== undefined ? edit.note        : r.note,
      isScheduled: ftIsScheduled(
        edit.installDate !== undefined ? edit.installDate : r.installDate,
        edit.note        !== undefined ? edit.note        : r.note,
      ),
      edited: true,
    };
  });
}

function ftIsScheduled(installDate, note) {
  if (!installDate) return true;  // 日付未定 → 予定扱い
  const d = String(installDate);
  const n = String(note || '');

  // 文字列に「予定」が含まれる場合は明示的に予定
  if (n.includes('予定') || d.includes('予定')) return true;

  const todayMs = new Date(new Date().toDateString()).getTime(); // 今日0:00

  // 日付範囲形式（例: 2023/04/10-14 や 2023/04-05）
  // → ハイフン前の開始日を取り出して過去か判定
  const mRange = d.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})-/);
  if (mRange) {
    const installMs = new Date(Number(mRange[1]), Number(mRange[2]) - 1, Number(mRange[3])).getTime();
    return installMs > todayMs;
  }
  // 年月範囲形式（例: 2023/04-05）→ 開始月の月末で比較
  const mMonthRange = d.match(/^(\d{4})\/(\d{1,2})-/);
  if (mMonthRange) {
    const endOfMonth = new Date(Number(mMonthRange[1]), Number(mMonthRange[2]), 0).getTime();
    return endOfMonth > todayMs;
  }

  // 確定日付（yyyy/mm/dd）
  const m = d.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) {
    const installMs = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
    return installMs > todayMs;
  }

  // 年月のみ（yyyy/mm）→ 当月末で比較
  const m2 = d.match(/^(\d{4})\/(\d{1,2})$/);
  if (m2) {
    const endOfMonth = new Date(Number(m2[1]), Number(m2[2]), 0).getTime();
    return endOfMonth > todayMs;
  }

  // 解釈できない形式は予定扱い
  return true;
}

// ============================================================
// 受注済み連動：orderStatusStore から FT 自動エントリを生成
// ============================================================
function buildAutoFtRows(baseRows) {
  // app.js の orderStatusStore / allData をグローバルから参照
  const store   = (typeof orderStatusStore !== 'undefined') ? orderStatusStore : {};
  const vessels = (typeof allData          !== 'undefined') ? allData          : [];

  const autoRows = [];
  const existNames = new Set(baseRows.map(r => ftNormalizeKey(r.vesselName)));

  // ordered ステータスの船で、FTリストに未掲載のものを追加
  Object.entries(store).forEach(([key, rec]) => {
    if (rec.status !== 'ordered') return;

    // vesselName を allData から逆引き
    const vessel = vessels.find(v =>
      (v.VESSEL_UID || v.BUILDERS_VESSEL_NUMBER || v.VESSEL_NAME || '') === key
    );
    const name = vessel ? (vessel.VESSEL_NAME || key) : key;
    if (!name) return;

    const normName = ftNormalizeKey(name);
    if (existNames.has(normName)) return; // 既にExcel行がある → スキップ

    // 編集オーバーレイを適用
    const edit = ftEditStore[normName] || {};
    const installDate = edit.installDate || '';
    const note        = edit.note        || '受注済み（搭載予定）';
    const serialNo    = edit.serialNo    || '';
    const os          = edit.os          || 'Windows10';

    autoRows.push({
      vesselName:  name,
      serialNo:    serialNo,
      os:          os,
      installDate: installDate,
      note:        note,
      isScheduled: true,
      source:      'auto',   // 受注連動
      edited:      !!Object.keys(edit).length,
    });
  });

  return autoRows;
}

// ============================================================
// EXCEL PARSER
// ============================================================
function parseFtExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  const targetSheet = workbook.SheetNames.find(n => n.includes('FT') || n.includes('管理'))
                   || workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheet];

  // 日付型セルを yyyy/mm/dd に変換
  Object.keys(sheet).filter(k => !k.startsWith('!')).forEach(k => {
    const c = sheet[k];
    if (c && c.t === 'd' && c.v instanceof Date) {
      const d = c.v;
      c.t = 's';
      c.v = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
      c.w = c.v;
    }
  });

  const jsonRows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', raw:false });

  const rows = [];
  let currentOs = '';
  let no = 1;

  for (let i = 0; i < jsonRows.length; i++) {
    const row  = jsonRows[i];
    const col0 = String(row[0] || '').trim();
    const col1 = String(row[1] || '').trim();
    const col2 = String(row[2] || '').trim();
    const col3 = String(row[3] || '').trim();
    const col4 = String(row[4] || '').trim();

    if (!col0 || col0.startsWith('Fleet Transfer') || col0 === '船名 (Vessel Name)') continue;
    if (col0.startsWith('▼')) {
      if      (col0.includes('7'))  currentOs = 'Windows7';
      else if (col0.includes('10')) currentOs = 'Windows10';
      else if (col0.includes('11')) currentOs = 'Windows11';
      continue;
    }

    const os = col2 || currentOs;
    rows.push({
      no:          no++,
      vesselName:  col0,
      serialNo:    col1,
      os:          os,
      installDate: col3,
      note:        col4,
      isScheduled: ftIsScheduled(col3, col4),
      source:      'excel',
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
  const win10     = rows.filter(r => r.os !== 'Windows7').length;
  const scheduled = rows.filter(r => r.isScheduled).length;

  document.getElementById('ftKpiTotal').textContent     = total.toLocaleString();
  document.getElementById('ftKpiWin7').textContent      = win7.toLocaleString();
  document.getElementById('ftKpiWin10').textContent     = win10.toLocaleString();
  document.getElementById('ftKpiScheduled').textContent = scheduled.toLocaleString();
}

// ============================================================
// BADGE HELPERS
// ============================================================
function ftOsBadge(os) {
  if (os === 'Windows7')  return `<span class="ft-os-badge ft-os-win7"><i class="fas fa-windows"></i> Win 7</span>`;
  if (os === 'Windows10') return `<span class="ft-os-badge ft-os-win10"><i class="fas fa-windows"></i> Win 10</span>`;
  if (os === 'Windows11') return `<span class="ft-os-badge ft-os-win11"><i class="fas fa-windows"></i> Win 11</span>`;
  return `<span class="ft-os-badge">${escFt(os)}</span>`;
}

function ftSourceBadge(r) {
  if (r.source === 'auto')
    return `<span class="ft-source-badge ft-source-auto"><i class="fas fa-link"></i> 受注連動</span>`;
  if (r.edited)
    return `<span class="ft-source-badge ft-source-edited"><i class="fas fa-pen"></i> 編集済</span>`;
  return '';
}

function escFt(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ============================================================
// TABLE
// ============================================================
function renderFtTable() {
  const start    = (ftCurrentPage - 1) * FT_PAGE_SIZE;
  const pageRows = ftFiltered.slice(start, start + FT_PAGE_SIZE);
  const tbody    = document.getElementById('ftTableBody');
  if (!tbody) return;

  if (ftFiltered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="ft-empty-cell">
      <i class="fas fa-search"></i><br>該当データなし</td></tr>`;
    document.getElementById('ftTableCount').textContent = '0 件表示';
    renderFtPagination();
    return;
  }

  tbody.innerHTML = pageRows.map(r => {
    const scheduledBadge = r.isScheduled
      ? `<span class="ft-scheduled-badge"><i class="fas fa-clock"></i> 予定</span>` : '';
    const dateCell = r.installDate
      ? `${escFt(r.installDate)}${scheduledBadge}`
      : `<span class="ft-no-date">未定</span>${scheduledBadge}`;

    let rowCls = '';
    if (r.source === 'auto') rowCls = 'ft-row-auto';
    else if (r.isScheduled)  rowCls = 'ft-row-scheduled';

    const normKey = escFt(ftNormalizeKey(r.vesselName));

    return `<tr class="${rowCls} ft-row-clickable" data-key="${normKey}" title="クリックして編集">
      <td class="ft-td-no">${r.no || ''}</td>
      <td class="ft-td-name">
        ${escFt(r.vesselName)}
        ${ftSourceBadge(r)}
      </td>
      <td class="ft-td-serial">${r.serialNo ? escFt(r.serialNo) : '<span class="ft-dash">—</span>'}</td>
      <td class="ft-td-os">${ftOsBadge(r.os)}</td>
      <td class="ft-td-date">${dateCell}</td>
      <td class="ft-td-note">${r.note ? escFt(r.note) : ''}</td>
      <td class="ft-td-action">
        <button class="ft-edit-btn" data-key="${normKey}" title="編集">
          <i class="fas fa-pen"></i>
        </button>
      </td>
    </tr>`;
  }).join('');

  const showing = Math.min(start + FT_PAGE_SIZE, ftFiltered.length);
  document.getElementById('ftTableCount').textContent =
    `${start + 1}–${showing} / ${ftFiltered.length} 件`;

  // 編集ボタン・行クリックイベント
  tbody.querySelectorAll('.ft-edit-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      openFtEditModal(btn.dataset.key);
    });
  });
  tbody.querySelectorAll('.ft-row-clickable').forEach(tr => {
    tr.addEventListener('click', () => openFtEditModal(tr.dataset.key));
  });

  renderFtPagination();
}

// ============================================================
// PAGINATION
// ============================================================
function renderFtPagination() {
  const container = document.getElementById('ftPagination');
  if (!container) return;
  const totalPages = Math.ceil(ftFiltered.length / FT_PAGE_SIZE);
  if (totalPages <= 1) { container.innerHTML = ''; return; }

  let html = `<button class="ft-page-btn" ${ftCurrentPage===1?'disabled':''} data-p="${ftCurrentPage-1}">
    <i class="fas fa-chevron-left"></i></button>`;
  const range = 2;
  for (let p = 1; p <= totalPages; p++) {
    if (p===1 || p===totalPages || Math.abs(p-ftCurrentPage)<=range) {
      html += `<button class="ft-page-btn ${p===ftCurrentPage?'active':''}" data-p="${p}">${p}</button>`;
    } else if (Math.abs(p-ftCurrentPage)===range+1) {
      html += `<span class="ft-page-ellipsis">…</span>`;
    }
  }
  html += `<button class="ft-page-btn" ${ftCurrentPage===totalPages?'disabled':''} data-p="${ftCurrentPage+1}">
    <i class="fas fa-chevron-right"></i></button>`;

  container.innerHTML = html;
  container.querySelectorAll('.ft-page-btn[data-p]').forEach(btn => {
    btn.addEventListener('click', () => {
      ftCurrentPage = Number(btn.dataset.p);
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
    if (ftOsFilter === 'Windows7'  && r.os !== 'Windows7') return false;
    if (ftOsFilter === 'Windows10' && r.os === 'Windows7') return false;
    if (ftSourceFilter === 'auto'   && r.source !== 'auto')          return false;
    if (ftSourceFilter === 'excel'  && r.source !== 'excel')         return false;
    if (ftSourceFilter === 'edited' && !r.edited && r.source !== 'auto') return false;
    if (ftSourceFilter === 'scheduled' && !r.isScheduled)            return false;
    if (q && !r.vesselName.toLowerCase().includes(q) && !(r.serialNo||'').toLowerCase().includes(q)) return false;
    return true;
  });

  ftCurrentPage = 1;
  renderFtTable();
}

// ============================================================
// EDIT MODAL
// ============================================================
function openFtEditModal(normKey) {
  // normKey でデータ行を探す
  const row = ftAllRows.find(r => ftNormalizeKey(r.vesselName) === normKey);
  if (!row) return;

  const edit = ftEditStore[normKey] || {};
  // 現在の表示値（編集オーバーレイ適用後）
  const cur = {
    serialNo:    edit.serialNo    !== undefined ? edit.serialNo    : (row.serialNo    || ''),
    installDate: edit.installDate !== undefined ? edit.installDate : (row.installDate || ''),
    os:          edit.os          !== undefined ? edit.os          : (row.os          || 'Windows10'),
    note:        edit.note        !== undefined ? edit.note        : (row.note        || ''),
  };

  const overlay = document.getElementById('ftEditOverlay');
  if (!overlay) return;

  overlay.querySelector('.ft-edit-vessel-name').textContent = row.vesselName;
  overlay.querySelector('#ftEditSerial').value      = cur.serialNo;
  overlay.querySelector('#ftEditDate').value        = cur.installDate;
  overlay.querySelector('#ftEditNote').value        = cur.note;

  // OSセレクト
  const osSelect = overlay.querySelector('#ftEditOs');
  osSelect.value = cur.os;

  // sourceバッジ
  overlay.querySelector('.ft-edit-source-info').innerHTML =
    row.source === 'auto'
      ? `<span class="ft-source-badge ft-source-auto"><i class="fas fa-link"></i> 受注連動エントリ</span>`
      : `<span class="ft-source-badge" style="background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0">Excelデータ</span>`;

  // 保存ボタン
  overlay.querySelector('#ftEditSave').onclick = () => {
    const newSerial = overlay.querySelector('#ftEditSerial').value.trim();
    const newDate   = overlay.querySelector('#ftEditDate').value.trim();
    const newOs     = overlay.querySelector('#ftEditOs').value;
    const newNote   = overlay.querySelector('#ftEditNote').value.trim();

    ftEditStore[normKey] = {
      serialNo:    newSerial,
      installDate: newDate,
      os:          newOs,
      note:        newNote,
      editedAt:    new Date().toISOString(),
    };
    ftSaveStore();

    // 全行に再適用して再描画
    ftRefreshAllRows();
    closeFtEditModal();

    // トースト通知（app.jsのtoast関数があれば使う）
    if (typeof toast === 'function') toast('FTデータを保存しました', 'success');
  };

  // リセットボタン（編集を元に戻す）
  const resetBtn = overlay.querySelector('#ftEditReset');
  if (resetBtn) {
    resetBtn.onclick = () => {
      if (!confirm('編集内容をリセットしてExcel/受注元のデータに戻しますか？')) return;
      delete ftEditStore[normKey];
      ftSaveStore();
      ftRefreshAllRows();
      closeFtEditModal();
      if (typeof toast === 'function') toast('編集をリセットしました', 'info');
    };
    resetBtn.disabled = !ftEditStore[normKey];
  }

  overlay.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
}

function closeFtEditModal() {
  const overlay = document.getElementById('ftEditOverlay');
  if (overlay) overlay.classList.add('hidden');
  document.body.style.overflow = '';
}

// ============================================================
// REFRESH（編集オーバーレイ + 受注連動を再構築）
// ============================================================
function ftRefreshAllRows() {
  // Excel由来の行（sourceが'excel'のもの）を取り出して編集適用
  const excelRows  = ftAllRows.filter(r => r.source === 'excel');
  const applied    = ftApplyEdits(excelRows);

  // 受注連動行を再構築
  const autoRows   = buildAutoFtRows(applied);

  // 番号を振り直して結合
  const combined = [...applied, ...autoRows].map((r, i) => ({ ...r, no: i + 1 }));
  ftAllRows = combined;

  applyFtFilters();
  renderFtKpi(ftAllRows);
}

// ============================================================
// CSV EXPORT
// ============================================================
function exportFtCsv() {
  const header = ['No.', '船名', 'Serial No.', 'OS', 'FT搭載日', '備考', 'ソース'].join(',');
  const rows   = ftFiltered.map(r =>
    [r.no || '',
     `"${(r.vesselName||'').replace(/"/g,'""')}"`,
     `"${(r.serialNo  ||'').replace(/"/g,'""')}"`,
     r.os || '',
     `"${(r.installDate||'').replace(/"/g,'""')}"`,
     `"${(r.note      ||'').replace(/"/g,'""')}"`,
     r.source === 'auto' ? '受注連動' : 'Excel',
    ].join(',')
  );
  const blob = new Blob(['\uFEFF'+header+'\n'+rows.join('\n')], { type:'text/csv;charset=utf-8;' });
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
function showFtDashboard(excelRows) {
  ftLoadStore(); // 編集データをロード

  // 編集オーバーレイを適用
  const applied  = ftApplyEdits(excelRows);
  // 受注連動行を追加
  const autoRows = buildAutoFtRows(applied);
  // 番号を振り直して結合
  const combined = [...applied, ...autoRows].map((r, i) => ({ ...r, no: i + 1 }));

  ftAllRows   = combined;
  ftFiltered  = combined;
  ftCurrentPage = 1;
  ftOsFilter  = 'all';
  ftSourceFilter = 'all';
  ftSearchQ   = '';

  // UI リセット
  document.querySelectorAll('#ftOsFilter .ft-filter-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.os === 'all'));
  document.querySelectorAll('#ftSourceFilter .ft-filter-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.src === 'all'));
  document.getElementById('ftSearchInput').value = '';

  renderFtKpi(ftAllRows);
  renderFtTable();

  document.getElementById('ftUploadSection').classList.add('hidden');
  document.getElementById('ftDashboard').classList.remove('hidden');
}

function showFtUpload() {
  document.getElementById('ftDashboard').classList.add('hidden');
  document.getElementById('ftUploadSection').classList.remove('hidden');
}

// ============================================================
// PAGE NAV
// ============================================================
function initPageNav() {
  document.querySelectorAll('.page-nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const page = btn.dataset.page;
      document.querySelectorAll('.page-nav-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      if (page === 'ft') {
        document.getElementById('uploadSection').classList.add('hidden');
        document.getElementById('dashboard').classList.add('hidden');
        document.getElementById('ftPage').classList.remove('hidden');
        // FTページに切り替えたとき受注連動を最新化
        if (ftAllRows.length > 0) ftRefreshAllRows();
      } else {
        document.getElementById('ftPage').classList.add('hidden');
        const totalEl  = document.getElementById('totalCount');
        const hasData  = totalEl && !totalEl.textContent.includes('—');
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
  const dropZone  = document.getElementById('ftDropZone');
  const fileInput = document.getElementById('ftFileInput');
  if (!dropZone || !fileInput) return;

  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) loadFtFile(file);
  });
  fileInput.addEventListener('change', e => {
    const file = e.target.files[0]; if (file) loadFtFile(file);
  });

  const btnBack   = document.getElementById('ftBtnBack');
  const btnExport = document.getElementById('ftBtnExport');
  if (btnBack)   btnBack.addEventListener('click', showFtUpload);
  if (btnExport) btnExport.addEventListener('click', exportFtCsv);

  // OS フィルター
  document.querySelectorAll('#ftOsFilter .ft-filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#ftOsFilter .ft-filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      ftOsFilter = btn.dataset.os;
      applyFtFilters();
    });
  });

  // ソース フィルター
  document.querySelectorAll('#ftSourceFilter .ft-filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#ftSourceFilter .ft-filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      ftSourceFilter = btn.dataset.src;
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

  // 編集モーダル：閉じるボタン（ヘッダーのX と フッターのキャンセル 両方）
  document.getElementById('ftEditClose')?.addEventListener('click', closeFtEditModal);
  document.getElementById('ftEditClose2')?.addEventListener('click', closeFtEditModal);
  document.getElementById('ftEditOverlay')?.addEventListener('click', e => {
    if (e.target === document.getElementById('ftEditOverlay')) closeFtEditModal();
  });
  // ESCキー
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeFtEditModal();
  });
}

function loadFtFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const rawRows = parseFtExcel(new Uint8Array(e.target.result));
      if (rawRows.length === 0) {
        alert('データが読み込めませんでした。ファイルの形式を確認してください。');
        return;
      }
      showFtDashboard(rawRows);
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
  ftLoadStore();
  initPageNav();
  initFtFileHandlers();
});
