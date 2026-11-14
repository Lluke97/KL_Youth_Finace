const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const multer = require('multer');
const db = require('./db');
const { syncCafe } = require('./lib/cafeSync');
const sheetsSync = require('./lib/sheetsSync');

// Purchase/expense categories specific to youth-group spending (distinct
// from the generic manual-entry form, which is mainly for 회비/income).
const PURCHASE_CATEGORIES = ['식비·다과', '교통비', '행사비', '비품·소모품', '후원·지원금', '기타'];

// Receipt photos live on the same Railway volume as the JSON data, so they
// survive redeploys. Optional — a purchase can be logged with no photo.
const RECEIPTS_DIR = path.join(__dirname, 'data', 'receipts');
if (!fs.existsSync(RECEIPTS_DIR)) fs.mkdirSync(RECEIPTS_DIR, { recursive: true });
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, RECEIPTS_DIR),
    filename: (req, file, cb) => cb(null, Date.now().toString(36) + Math.random().toString(36).slice(2, 8) + path.extname(file.originalname || '').slice(0, 10))
  }),
  limits: { fileSize: 8 * 1024 * 1024 },
  fileFilter: (req, file, cb) => cb(null, /^image\//.test(file.mimetype))
});

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;
const ADMIN_USER = process.env.ADMIN_USER || 'admin';
const ADMIN_PASS = process.env.ADMIN_PASS || 'changeme123';

if (ADMIN_PASS === 'changeme123') {
  console.warn('[youth-budget] WARNING: using the default admin password. Set ADMIN_USER / ADMIN_PASS env vars before going live.');
}

const ADMIN_ACCOUNTS = [{ user: ADMIN_USER, pass: ADMIN_PASS }];
if (process.env.ADMIN_USER2 && process.env.ADMIN_PASS2) {
  ADMIN_ACCOUNTS.push({ user: process.env.ADMIN_USER2, pass: process.env.ADMIN_PASS2 });
}

// Toggle: set REQUIRE_ADMIN_AUTH=false (Railway variable) to open /admin and
// /api/admin/* with no login — e.g. while only Luke/집사님 are using it and
// the login prompt is just friction. Set it back to anything else (or
// remove it) to require the admin login again; no code change needed.
const REQUIRE_ADMIN_AUTH = process.env.REQUIRE_ADMIN_AUTH !== 'false';

function safeEqual(a, b) {
  return crypto.timingSafeEqual(Buffer.from(String(a).padEnd(128)), Buffer.from(String(b).padEnd(128)));
}

// ---- same small HTTP Basic Auth + brute-force lockout as cafe-pos ----
const FAILED_ATTEMPT_LIMIT = 5;
const LOCKOUT_MS = 15 * 60 * 1000;
const failedAttempts = new Map();

function clientIp(req) {
  return (req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.socket.remoteAddress || 'unknown';
}

function requireAdmin(req, res, next) {
  const ip = clientIp(req);
  const record = failedAttempts.get(ip);
  if (record && record.lockedUntil && record.lockedUntil > Date.now()) {
    const waitMin = Math.ceil((record.lockedUntil - Date.now()) / 60000);
    return res.status(429).send(`Too many failed attempts. Try again in about ${waitMin} minute(s).`);
  }

  const header = req.headers.authorization || '';
  const [scheme, encoded] = header.split(' ');
  if (scheme === 'Basic' && encoded) {
    const decoded = Buffer.from(encoded, 'base64').toString('utf8');
    const sep = decoded.indexOf(':');
    const user = decoded.slice(0, sep);
    const pass = decoded.slice(sep + 1);
    const matched = ADMIN_ACCOUNTS.some((acc) => safeEqual(user, acc.user) && safeEqual(pass, acc.pass));
    if (matched) {
      failedAttempts.delete(ip);
      return next();
    }
  }

  const next_count = (record && record.count ? record.count : 0) + 1;
  if (next_count >= FAILED_ATTEMPT_LIMIT) {
    failedAttempts.set(ip, { count: next_count, lockedUntil: Date.now() + LOCKOUT_MS });
  } else {
    failedAttempts.set(ip, { count: next_count, lockedUntil: 0 });
  }

  res.set('WWW-Authenticate', 'Basic realm="Youth Budget Admin"');
  return res.status(401).send('Authentication required.');
}

function maybeRequireAdmin(req, res, next) {
  if (!REQUIRE_ADMIN_AUTH) return next();
  return requireAdmin(req, res, next);
}

// Asia/Kuala_Lumpur "today", same convention as cafe-pos.
const MY_TZ = 'Asia/Kuala_Lumpur';
const myDateFormatter = new Intl.DateTimeFormat('en-CA', { timeZone: MY_TZ, year: 'numeric', month: '2-digit', day: '2-digit' });
function todayStr() {
  return myDateFormatter.format(new Date());
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// One-time fix: an early cafe-pos test pushed a settlement with all-zero
// cash/TnG/bank via the webhook, which is now the "latest" settlement and
// displaced the correct real balance on the dashboard. Guarded by an env
// var so it only runs when explicitly asked, and idempotent via a fixed id
// so a restart or redeploy never applies it twice.
if (process.env.FIX_SETTLEMENT_20261114 === '1') {
  const __settlements = db.getSettlements();
  const __marker = 'settle_fix20261114';
  if (!__settlements.some((s) => s.id === __marker)) {
    __settlements.push({
      id: __marker,
      date: '2026-11-14',
      cash: 1003.7,
      tng: 1000,
      bank: 15438.6,
      total: 1003.7 + 1000 + 15438.6,
      source: 'manual',
      memo: '실제 잔액 수동 보정 (현금 1,003.70 / 터치앤고 1,000 / 은행 15,438.60)',
      createdAt: new Date().toISOString()
    });
    db.saveSettlements(__settlements);
    console.log('[youth-budget] applied FIX_SETTLEMENT_20261114 migration');
  }
}

// ---------------- Public pages + read-only API (dashboard is open to all) ----------------
app.use(express.static(path.join(__dirname, 'public')));
app.use('/receipts', express.static(RECEIPTS_DIR));

app.get('/api/purchase-categories', (req, res) => {
  res.json(PURCHASE_CATEGORIES);
});

app.get('/api/transactions', (req, res) => {
  let txns = db.getTransactions();
  const { from, to, source, type, tag, category } = req.query;
  if (from) txns = txns.filter((t) => t.date >= from);
  if (to) txns = txns.filter((t) => t.date <= to);
  if (source) txns = txns.filter((t) => t.source === source);
  if (type) txns = txns.filter((t) => t.type === type);
  if (tag) txns = txns.filter((t) => t.tag === tag);
  if (category) txns = txns.filter((t) => t.category === category);
  txns.sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt));
  res.json(txns);
});

app.get('/api/summary', (req, res) => {
  const txns = db.getTransactions();
  let balance = 0;
  const bySource = {}; // source -> { income, expense }
  const byTag = {}; // tag -> { income, expense }  (cafe-ops / youth-general / needs-review)
  const byCategory = {}; // category -> { income, expense }
  const monthly = {}; // YYYY-MM -> { income, expense }

  txns.forEach((t) => {
    const signedAmount = t.type === 'income' ? t.amount : -t.amount;
    balance += signedAmount;

    if (!bySource[t.source]) bySource[t.source] = { income: 0, expense: 0 };
    bySource[t.source][t.type] += t.amount;

    const tagKey = t.tag || 'untagged';
    if (!byTag[tagKey]) byTag[tagKey] = { income: 0, expense: 0 };
    byTag[tagKey][t.type] += t.amount;

    const catKey = t.category || '기타';
    if (!byCategory[catKey]) byCategory[catKey] = { income: 0, expense: 0 };
    byCategory[catKey][t.type] += t.amount;

    const month = t.date.slice(0, 7);
    if (!monthly[month]) monthly[month] = { month, income: 0, expense: 0 };
    monthly[month][t.type] += t.amount;
  });

  const monthlyList = Object.values(monthly).sort((a, b) => a.month.localeCompare(b.month));
  const needsReviewCount = txns.filter((t) => t.tag === 'needs-review').length;

  res.json({ balance, bySource, byTag, byCategory, monthly: monthlyList, needsReviewCount, transactionCount: txns.length });
});

// Settlements (실제 현금/터치앤고/은행 잔액) — read-only, same access as the
// rest of the dashboard.
app.get('/api/settlements', (req, res) => {
  const settlements = db.getSettlements().slice().sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  res.json(settlements);
});

// Cafe-pos pushes a completed settlement here automatically. This is NOT
// behind the admin login — cafe-pos is a different server calling in, not a
// browser — it's behind its own shared-secret header instead.
app.post('/api/webhook/cafe-settlement', (req, res) => {
  const secret = process.env.SETTLEMENT_WEBHOOK_SECRET;
  if (!secret) return res.status(503).json({ error: 'SETTLEMENT_WEBHOOK_SECRET not configured on this server' });
  if (req.headers['x-webhook-secret'] !== secret) {
    return res.status(401).json({ error: 'invalid webhook secret' });
  }
  const { date, cash, tng, bank, memo } = req.body || {};
  if (typeof cash !== 'number' || typeof tng !== 'number') {
    return res.status(400).json({ error: 'cash and tng must be numbers' });
  }
  const bankVal = typeof bank === 'number' ? bank : null;
  const entry = {
    id: 'settle_' + uid(),
    date: date || todayStr(),
    cash,
    tng,
    bank: bankVal,
    total: cash + tng + (bankVal || 0),
    source: 'cafe-pos',
    memo: memo || '카페 정산 자동 반영',
    createdAt: new Date().toISOString()
  };
  const list = db.getSettlements();
  list.push(entry);
  db.saveSettlements(list);
  res.status(201).json(entry);
});

// ---------------- Admin-only API ----------------
app.use('/api/admin', maybeRequireAdmin);

app.post('/api/admin/transactions', (req, res) => {
  const { date, type, source, tag, category, item, amount, memo, author } = req.body || {};
  if (!date || (type !== 'income' && type !== 'expense') || !category || typeof amount !== 'number' || amount <= 0) {
    return res.status(400).json({ error: 'date, type (income|expense), category and a positive amount are required' });
  }
  const entry = {
    id: 'tx_' + uid(),
    date,
    type,
    source: source || 'manual',
    tag: tag || null,
    category,
    item: item || category,
    amount,
    memo: memo || '',
    author: author || '',
    externalId: null,
    createdAt: new Date().toISOString()
  };
  const txns = db.getTransactions();
  txns.push(entry);
  db.saveTransactions(txns);
  res.status(201).json(entry);
});

// Dedicated purchase/expense entry: preset categories + optional receipt
// photo, instead of the free-text category in the generic form above.
app.post('/api/admin/purchases', upload.single('receipt'), (req, res) => {
  const { date, category, vendor, item, amount, memo, author } = req.body || {};
  const amountNum = parseFloat(amount);
  if (!date || !PURCHASE_CATEGORIES.includes(category) || !item || Number.isNaN(amountNum) || amountNum <= 0) {
    return res.status(400).json({ error: `date, category (${PURCHASE_CATEGORIES.join('|')}), item and a positive amount are required` });
  }
  const entry = {
    id: 'tx_' + uid(),
    date,
    type: 'expense',
    source: 'manual',
    tag: 'youth-general',
    category,
    item,
    vendor: vendor || '',
    amount: amountNum,
    memo: memo || '',
    author: author || '',
    receiptUrl: req.file ? `/receipts/${req.file.filename}` : null,
    externalId: null,
    createdAt: new Date().toISOString()
  };
  const txns = db.getTransactions();
  txns.push(entry);
  db.saveTransactions(txns);
  res.status(201).json(entry);
});

app.put('/api/admin/transactions/:id', (req, res) => {
  const txns = db.getTransactions();
  const idx = txns.findIndex((t) => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'not found' });
  const fields = ['date', 'type', 'source', 'tag', 'category', 'item', 'vendor', 'amount', 'memo', 'author'];
  fields.forEach((f) => { if (req.body && req.body[f] !== undefined) txns[idx][f] = req.body[f]; });
  db.saveTransactions(txns);
  res.json(txns[idx]);
});

app.delete('/api/admin/transactions/:id', (req, res) => {
  const txns = db.getTransactions().filter((t) => t.id !== req.params.id);
  db.saveTransactions(txns);
  res.json({ ok: true });
});

// Pull new orders/expenses from cafe-pos.
// Manual settlement entry (e.g. seeding today's actual balances by hand).
app.post('/api/admin/settlements', (req, res) => {
  const { date, cash, tng, bank, memo } = req.body || {};
  if (typeof cash !== 'number' || typeof tng !== 'number' || typeof bank !== 'number') {
    return res.status(400).json({ error: 'cash, tng, bank must all be numbers' });
  }
  const entry = {
    id: 'settle_' + uid(),
    date: date || todayStr(),
    cash,
    tng,
    bank,
    total: cash + tng + bank,
    source: 'manual',
    memo: memo || '',
    createdAt: new Date().toISOString()
  };
  const list = db.getSettlements();
  list.push(entry);
  db.saveSettlements(list);
  res.status(201).json(entry);
});

app.delete('/api/admin/settlements/:id', (req, res) => {
  const list = db.getSettlements().filter((s) => s.id !== req.params.id);
  db.saveSettlements(list);
  res.json({ ok: true });
});

app.post('/api/admin/sync-cafe', async (req, res) => {
  const baseUrl = process.env.CAFE_BASE_URL || 'https://jusumcafe.up.railway.app';
  const user = process.env.CAFE_ADMIN_USER;
  const pass = process.env.CAFE_ADMIN_PASS;
  if (!user || !pass) {
    return res.status(400).json({ error: 'CAFE_ADMIN_USER / CAFE_ADMIN_PASS 환경변수가 설정되어 있지 않습니다 (cafe-pos의 ADMIN_USER/ADMIN_PASS와 동일한 값).' });
  }
  try {
    const settings = db.getSettings();
    const { newTxns, lastOrderAt, lastExpenseCreatedAt } = await syncCafe({ baseUrl, user, pass, since: settings.cafeSync });

    const txns = db.getTransactions();
    const existingIds = new Set(txns.map((t) => t.externalId).filter(Boolean));
    const toAdd = newTxns.filter((t) => !existingIds.has(t.externalId));
    toAdd.forEach((t) => txns.push(t));
    db.saveTransactions(txns);

    settings.cafeSync = { lastOrderAt, lastExpenseCreatedAt };
    db.saveSettings(settings);

    res.json({ added: toAdd.length, needsReview: toAdd.filter((t) => t.tag === 'needs-review').length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/admin/export-sheet', async (req, res) => {
  if (!sheetsSync.isConfigured()) {
    return res.status(400).json({ error: '구글 시트 연동이 설정되어 있지 않습니다 (GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_JSON 필요).' });
  }
  try {
    const rows = await sheetsSync.exportTransactions(db.getTransactions());
    res.json({ exported: rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/admin/sheets-status', (req, res) => {
  res.json({ configured: sheetsSync.isConfigured() });
});

// Admin page itself needs auth before the browser even gets the HTML
// (unless REQUIRE_ADMIN_AUTH=false — see above).
app.get('/admin', maybeRequireAdmin, (req, res) => {
  res.sendFile(path.join(__dirname, 'public-admin', 'admin.html'));
});
app.use('/admin-assets', maybeRequireAdmin, express.static(path.join(__dirname, 'public-admin')));

app.listen(PORT, () => {
  console.log(`[youth-budget] listening on http://localhost:${PORT}`);
  console.log(`[youth-budget] dashboard: http://localhost:${PORT}/`);
  console.log(`[youth-budget] admin:     http://localhost:${PORT}/admin  (user: ${ADMIN_USER})`);
});
