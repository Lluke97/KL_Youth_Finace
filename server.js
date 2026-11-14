const express = require('express');
const path = require('path');
const crypto = require('crypto');
const db = require('./db');
const { syncCafe } = require('./lib/cafeSync');
const sheetsSync = require('./lib/sheetsSync');

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

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// ---------------- Public pages + read-only API (dashboard is open to all) ----------------
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/transactions', (req, res) => {
  let txns = db.getTransactions();
  const { from, to, source, type, tag } = req.query;
  if (from) txns = txns.filter((t) => t.date >= from);
  if (to) txns = txns.filter((t) => t.date <= to);
  if (source) txns = txns.filter((t) => t.source === source);
  if (type) txns = txns.filter((t) => t.type === type);
  if (tag) txns = txns.filter((t) => t.tag === tag);
  txns.sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt));
  res.json(txns);
});

app.get('/api/summary', (req, res) => {
  const txns = db.getTransactions();
  let balance = 0;
  const bySource = {}; // source -> { income, expense }
  const byTag = {}; // tag -> { income, expense }  (cafe-ops / youth-general / needs-review)
  const monthly = {}; // YYYY-MM -> { income, expense }

  txns.forEach((t) => {
    const signedAmount = t.type === 'income' ? t.amount : -t.amount;
    balance += signedAmount;

    if (!bySource[t.source]) bySource[t.source] = { income: 0, expense: 0 };
    bySource[t.source][t.type] += t.amount;

    const tagKey = t.tag || 'untagged';
    if (!byTag[tagKey]) byTag[tagKey] = { income: 0, expense: 0 };
    byTag[tagKey][t.type] += t.amount;

    const month = t.date.slice(0, 7);
    if (!monthly[month]) monthly[month] = { month, income: 0, expense: 0 };
    monthly[month][t.type] += t.amount;
  });

  const monthlyList = Object.values(monthly).sort((a, b) => a.month.localeCompare(b.month));
  const needsReviewCount = txns.filter((t) => t.tag === 'needs-review').length;

  res.json({ balance, bySource, byTag, monthly: monthlyList, needsReviewCount, transactionCount: txns.length });
});

// ---------------- Admin-only API ----------------
app.use('/api/admin', requireAdmin);

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

app.put('/api/admin/transactions/:id', (req, res) => {
  const txns = db.getTransactions();
  const idx = txns.findIndex((t) => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'not found' });
  const fields = ['date', 'type', 'source', 'tag', 'category', 'item', 'amount', 'memo', 'author'];
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

// Admin page itself needs auth before the browser even gets the HTML.
app.get('/admin', requireAdmin, (req, res) => {
  res.sendFile(path.join(__dirname, 'public-admin', 'admin.html'));
});
app.use('/admin-assets', requireAdmin, express.static(path.join(__dirname, 'public-admin')));

app.listen(PORT, () => {
  console.log(`[youth-budget] listening on http://localhost:${PORT}`);
  console.log(`[youth-budget] dashboard: http://localhost:${PORT}/`);
  console.log(`[youth-budget] admin:     http://localhost:${PORT}/admin  (user: ${ADMIN_USER})`);
});
