/**
 * MOL 船舶管理リスト — 共有サーバー
 * - 受注状態を JSON ファイルで永続化（全ユーザー共有）
 * - アップロードされた CSV データも保存・共有
 * - 静的ファイル配信 (index.html / css / js)
 */
'use strict';

const express = require('express');
const cors    = require('cors');
const fs      = require('fs');
const path    = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

// データ保存先
const DATA_DIR        = path.join(__dirname, 'data');
const ORDER_STATUS_FILE = path.join(DATA_DIR, 'order_status.json');
const CSV_DATA_FILE     = path.join(DATA_DIR, 'csv_data.json');   // { csv: '...', updatedAt, updatedBy }

// data/ ディレクトリがなければ作成
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ファイルを読む汎用ヘルパー
function readJSON(file, def) {
  try {
    if (fs.existsSync(file)) return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch(e) {}
  return def;
}
function writeJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8');
}

// ミドルウェア
app.use(cors());
app.use(express.json({ limit: '10mb' }));   // CSV テキストは大きくなり得る
app.use(express.static(path.join(__dirname)));  // index.html / css / js を配信

// ============================================================
// API: 受注状態
// ============================================================

// GET /api/order-status  → 全受注状態を返す
app.get('/api/order-status', (req, res) => {
  const data = readJSON(ORDER_STATUS_FILE, {});
  res.json({ ok: true, data });
});

// POST /api/order-status  body: { key, status, quoteDate, orderedDate, note }
app.post('/api/order-status', (req, res) => {
  const { key, status, quoteDate, orderedDate, note } = req.body;
  if (!key) return res.status(400).json({ ok: false, error: 'key is required' });

  const store = readJSON(ORDER_STATUS_FILE, {});
  store[key] = {
    status:      status      || 'other',
    quoteDate:   quoteDate   || '',
    orderedDate: orderedDate || '',
    note:        note        || '',
    updatedAt:   new Date().toISOString(),
  };
  writeJSON(ORDER_STATUS_FILE, store);
  res.json({ ok: true, data: store[key] });
});

// DELETE /api/order-status/:key  → 1件削除
app.delete('/api/order-status/:key', (req, res) => {
  const key   = decodeURIComponent(req.params.key);
  const store = readJSON(ORDER_STATUS_FILE, {});
  delete store[key];
  writeJSON(ORDER_STATUS_FILE, store);
  res.json({ ok: true });
});

// ============================================================
// API: CSV データ（最後にアップロードされたCSVを全員で共有）
// ============================================================

// GET /api/csv  → 保存済み CSV テキストと更新情報を返す
app.get('/api/csv', (req, res) => {
  const data = readJSON(CSV_DATA_FILE, null);
  if (!data) return res.json({ ok: true, data: null });
  res.json({ ok: true, data });
});

// POST /api/csv  body: { csv, updatedBy }
app.post('/api/csv', (req, res) => {
  const { csv, updatedBy } = req.body;
  if (!csv) return res.status(400).json({ ok: false, error: 'csv is required' });
  const payload = {
    csv,
    updatedBy:  updatedBy || '不明',
    updatedAt:  new Date().toISOString(),
  };
  writeJSON(CSV_DATA_FILE, payload);
  res.json({ ok: true, updatedAt: payload.updatedAt });
});

// ============================================================
// 起動
// ============================================================
app.listen(PORT, '0.0.0.0', () => {
  console.log(`MOL Shiplist Server running on http://0.0.0.0:${PORT}`);
});
