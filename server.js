/**
 * MOL 船舶管理リスト — 共有サーバー (MongoDB版)
 * - 受注状態をMongoDBで永続化（全ユーザー共有）
 * - アップロードされた CSV データも保存・共有
 * - 静的ファイル配信 (index.html / css / js)
 */
'use strict';

const express  = require('express');
const cors     = require('cors');
const mongoose = require('mongoose');
const path     = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;
const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://aihara1089_db_user:YWk4CgmUjq3VRDYv@cluster0.otnzlqa.mongodb.net/ship_management?appName=Cluster0';

// ─── ミドルウェア ───────────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// ─── MongoDBスキーマ ────────────────────────────────────────

// 受注状態（keyごとに1レコード）
const orderStatusSchema = new mongoose.Schema({
  key:         { type: String, unique: true, required: true },
  status:      { type: String, default: 'other' },
  quoteDate:   { type: String, default: '' },
  orderedDate: { type: String, default: '' },
  note:        { type: String, default: '' },
  updatedAt:   { type: String, default: '' },
}, { versionKey: false });

// CSVデータ（1件のみ保存）
const csvDataSchema = new mongoose.Schema({
  _id:       { type: String, default: 'csv' },
  csv:       String,
  updatedBy: String,
  updatedAt: String,
}, { versionKey: false });

const OrderStatus = mongoose.model('OrderStatus', orderStatusSchema);
const CsvData     = mongoose.model('CsvData',     csvDataSchema);

// ============================================================
// API: 受注状態
// ============================================================

// GET /api/order-status → 全受注状態を返す
app.get('/api/order-status', async (req, res) => {
  try {
    const records = await OrderStatus.find({}).lean();
    const data = {};
    records.forEach(r => {
      const { key, ...rest } = r;
      delete rest._id;
      data[key] = rest;
    });
    res.json({ ok: true, data });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// POST /api/order-status  body: { key, status, quoteDate, orderedDate, note }
app.post('/api/order-status', async (req, res) => {
  try {
    const { key, status, quoteDate, orderedDate, note } = req.body;
    if (!key) return res.status(400).json({ ok: false, error: 'key is required' });

    const record = {
      status:      status      || 'other',
      quoteDate:   quoteDate   || '',
      orderedDate: orderedDate || '',
      note:        note        || '',
      updatedAt:   new Date().toISOString(),
    };
    await OrderStatus.findOneAndUpdate(
      { key },
      { key, ...record },
      { upsert: true, new: true }
    );
    res.json({ ok: true, data: record });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// DELETE /api/order-status/:key → 1件削除
app.delete('/api/order-status/:key', async (req, res) => {
  try {
    const key = decodeURIComponent(req.params.key);
    await OrderStatus.findOneAndDelete({ key });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ============================================================
// API: CSV データ
// ============================================================

// GET /api/csv → 保存済みCSVテキストと更新情報を返す
app.get('/api/csv', async (req, res) => {
  try {
    const data = await CsvData.findById('csv').lean();
    if (!data) return res.json({ ok: true, data: null });
    delete data._id;
    res.json({ ok: true, data });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// POST /api/csv  body: { csv, updatedBy }
app.post('/api/csv', async (req, res) => {
  try {
    const { csv, updatedBy } = req.body;
    if (!csv) return res.status(400).json({ ok: false, error: 'csv is required' });
    const payload = {
      csv,
      updatedBy:  updatedBy || '不明',
      updatedAt:  new Date().toISOString(),
    };
    await CsvData.findByIdAndUpdate('csv', payload, { upsert: true });
    res.json({ ok: true, updatedAt: payload.updatedAt });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ============================================================
// 起動
// ============================================================
mongoose.connect(MONGO_URI)
  .then(() => {
    console.log('✅ MongoDB接続成功');
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`🚢 MOL Shiplist Server running on http://0.0.0.0:${PORT}`);
    });
  })
  .catch(e => {
    console.error('❌ MongoDB接続失敗:', e.message);
    process.exit(1);
  });
