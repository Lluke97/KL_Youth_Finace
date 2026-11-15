const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const multer = require('multer');
const XLSX = require('xlsx');
const db = require('./db');
const { syncCafe } = require('./lib/cafeSync');
const sheetsSync = require('./lib/sheetsSync');

// Purchase/expense categories specific to youth-group spending (distinct
// from the generic manual-entry form, which is mainly for 회비/income).
const PURCHASE_CATEGORIES = ['식비·다과', '교통비', '행사비', '비품·소모품', '후원·지원금', '기타'];

// Income categories: 회비 (membership dues) and 후원/헌금 (outside
// sponsorship/offerings) are the two main real-world sources besides the
// cafe, so they get their own preset dropdown instead of free-text.
const INCOME_CATEGORIES = ['회비', '후원금', '헌금', '기타수입'];

// How a purchase/expense was actually paid for.
const PAYMENT_METHODS = ['현금', '카드', '터치앤고'];

// How money came in, for income entries.
const RECEIVE_CHANNELS = ['현금', '터치앤고', '계좌이체'];

// Every transaction's real-world "channel" (cash / Touch 'n Go / bank) is
// normalized to one of these three buckets, used to auto-roll the actual
// balance breakdown forward from the last manual settlement instead of
// requiring a fresh manual count every time something is logged. A card
// payment/receipt is treated as hitting the bank account.
function paymentToChannel(payment) {
  if (payment === '현금') return 'cash';
  if (payment === '터치앤고') return 'tng';
  if (payment === '카드') return 'bank';
  return null;
}
function channelToBucket(channel) {
  if (channel === '현금') return 'cash';
  if (channel === '터치앤고') return 'tng';
  if (channel === '계좌이체') return 'bank';
  return null;
}
const CHANNEL_LABEL = { cash: '현금', tng: '터치앤고', bank: '계좌/카드' };

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

// Bulk upload (엑셀/CSV): kept in memory, never written to disk — it's
// parsed once and discarded, unlike receipt photos above.
const uploadBulk = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

const BULK_TEMPLATE_HEADER = ['날짜', '구분', '카테고리', '항목', '금액', '메모', '작성자'];

// Accepts an Excel serial date number, a Date, or a string like
// "2026-11-14" / "2026/11/14" and normalizes to YYYY-MM-DD. Same approach
// as scripts/import-excel.js, duplicated here since that's a standalone
// CLI script, not a module other code can require from.
function parseFlexibleDate(raw) {
  let d;
  if (typeof raw === 'number') {
    const parsed = XLSX.SSF.parse_date_code(raw);
    d = new Date(parsed.y, parsed.m - 1, parsed.d);
  } else if (raw instanceof Date) {
    d = raw;
  } else {
    d = new Date(String(raw).trim().replace(/\./g, '-').replace(/\//g, '-'));
  }
  if (Number.isNaN(d.getTime())) return null;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// A small hand-rolled CSV parser instead of handing .csv files to the XLSX
// library's sheet reader: XLSX auto-sniffs date-looking strings in CSV into
// Excel serial numbers, and that conversion picks up a timezone-dependent
// fractional offset (e.g. "2026-11-01" -> 46327.333...) that then parses
// back to the wrong calendar day. Plain-text CSV parsing keeps every cell
// exactly as typed, which parseFlexibleDate's string branch handles correctly.
function parseCsvRows(text) {
  const lines = text.split(/\r\n|\n|\r/).filter((l) => l.length > 0);
  if (!lines.length) return [];
  function parseLine(line) {
    const out = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (inQuotes) {
        if (c === '"') {
          if (line[i + 1] === '"') { cur += '"'; i++; } else inQuotes = false;
        } else cur += c;
      } else if (c === '"') inQuotes = true;
      else if (c === ',') { out.push(cur); cur = ''; } else cur += c;
    }
    out.push(cur);
    return out;
  }
  const header = parseLine(lines[0]).map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const cells = parseLine(line);
    const obj = {};
    header.forEach((h, idx) => { obj[h] = cells[idx] !== undefined ? cells[idx].trim() : ''; });
    return obj;
  });
}

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

// Same toggle idea, for the simple shared passcode that now gates the whole
// site (not just /admin) — this is sensitive financial data.
const REQUIRE_SITE_AUTH = process.env.REQUIRE_SITE_AUTH !== 'false';
const SITE_PASS = process.env.SITE_PASS || 'changeme';
if (SITE_PASS === 'changeme') {
  console.warn('[youth-budget] WARNING: using the default site password. Set SITE_PASS env var before going live.');
}

function safeEqual(a, b) {
  return crypto.timingSafeEqual(Buffer.from(String(a).padEnd(128)), Buffer.from(String(b).padEnd(128)));
}

function isAdminAccount(user, pass) {
  return ADMIN_ACCOUNTS.some((acc) => safeEqual(user, acc.user) && safeEqual(pass, acc.pass));
}

// ---- same small HTTP Basic Auth + brute-force lockout as cafe-pos, used
// for both the site-wide passcode and the admin login (each gets its own
// lockout bookkeeping so a failed attempt on one doesn't lock out the other) ----
const FAILED_ATTEMPT_LIMIT = 5;
const LOCKOUT_MS = 15 * 60 * 1000;

function clientIp(req) {
  return (req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.socket.remoteAddress || 'unknown';
}

function makeAuthGate({ realm, failMessage, matchFn }) {
  const failedAttempts = new Map();
  return function gate(req, res, next) {
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
      if (matchFn(user, pass)) {
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

    res.set('WWW-Authenticate', `Basic realm="${realm}"`);
    return res.status(401).send(failMessage);
  };
}

const requireAdmin = makeAuthGate({
  realm: 'Jusumjusum Finance Admin',
  failMessage: 'Authentication required.',
  matchFn: (user, pass) => isAdminAccount(user, pass)
});

// Any username works — only the shared passcode (or a full admin login)
// matters. ADMIN_ACCOUNTS also pass here so an admin who doesn't know the
// shared passcode can still get in with their own login.
const requireSiteAccess = makeAuthGate({
  realm: 'Jusumjusum Finance',
  failMessage: '비밀번호가 필요합니다. 문의: 회장',
  matchFn: (user, pass) => safeEqual(pass, SITE_PASS) || isAdminAccount(user, pass)
});

function maybeRequireAdmin(req, res, next) {
  if (!REQUIRE_ADMIN_AUTH) return next();
  return requireAdmin(req, res, next);
}

function maybeRequireSiteAccess(req, res, next) {
  // cafe-pos calls this server-to-server with its own shared-secret header
  // (see the route below) — it can't answer a browser login prompt, so it's
  // excluded from the site passcode gate.
  if (req.path === '/api/webhook/cafe-settlement') return next();
  if (!REQUIRE_SITE_AUTH) return next();
  return requireSiteAccess(req, res, next);
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

// One-time fix: the 2026-11-14 settlement above was taken before the
// RM348 "토요모임 저녁 · 돈까스랩" purchase was paid for by card, so it no
// longer matches the ledger (which already counts that expense). Adds a
// follow-up settlement with the card payment subtracted from the bank
// figure. Guarded by an env var + fixed id, same pattern as the fix above.
if (process.env.SETTLE_CARDFIX_20261114 === '1') {
  const __settlements2 = db.getSettlements();
  const __marker2 = 'settle_cardfix20261114';
  if (!__settlements2.some((s) => s.id === __marker2)) {
    const cash = 1003.7, tng = 1000, bank = 15438.6 - 348;
    __settlements2.push({
      id: __marker2,
      date: '2026-11-14',
      cash,
      tng,
      bank,
      total: cash + tng + bank,
      source: 'manual',
      memo: '토요모임 저녁(돈까스랩) RM348 카드 결제 반영 (은행 15,438.60 → 15,090.60)',
      createdAt: new Date().toISOString()
    });
    db.saveSettlements(__settlements2);
    console.log('[youth-budget] applied SETTLE_CARDFIX_20261114 migration');
  }
}

// ---------------- Site-wide passcode gate (sensitive financial data) ----------------
app.use(maybeRequireSiteAccess);

// ---------------- Public pages + read-only API (dashboard is open to anyone past the gate above) ----------------
app.use(express.static(path.join(__dirname, 'public')));
app.use('/receipts', express.static(RECEIPTS_DIR));

app.get('/api/purchase-categories', (req, res) => {
  res.json(PURCHASE_CATEGORIES);
});

app.get('/api/income-categories', (req, res) => {
  res.json(INCOME_CATEGORIES);
});

app.get('/api/payment-methods', (req, res) => {
  res.json(PAYMENT_METHODS);
});

app.get('/api/receive-channels', (req, res) => {
  res.json(RECEIVE_CHANNELS);
});

// Auto-rolled actual-balance breakdown: start from the most recent manual
// settlement (the last time someone actually counted cash/TnG/bank) and
// add/subtract every transaction logged since then by its channel
// (cash/tng/bank), instead of requiring a fresh manual count every time.
// Transactions without a channel (most cafe-synced expenses, since
// cafe-pos doesn't record a payment method for those) can't be placed in
// any bucket and are called out separately so the gap is visible instead
// of silently wrong.
function computeBalanceBreakdown() {
  // Only a real physical-count baseline (manual entry or the cafe-pos
  // settlement webhook) is used as the anchor — never an "auto" row, which
  // is just a snapshot of this same calculation at an earlier moment. If an
  // auto row were used as the anchor, deleting or editing a transaction
  // that happened *before* that snapshot would have no way to flow through
  // (it would already be "baked into" the frozen baseline forever). Always
  // anchoring on the last real count and re-summing every transaction that
  // still exists since then means a delete/edit anywhere is reflected
  // immediately, with no stale numbers to clean up.
  const settlements = db.getSettlements().filter((s) => s.source !== 'auto').sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  const baseline = settlements[0] || null;
  const baselineCreatedAt = baseline ? baseline.createdAt : null;
  const txns = db.getTransactions();

  let cash = 0, tng = 0, bank = 0, unattributedCount = 0, unattributedAmount = 0;
  txns.forEach((t) => {
    if (baselineCreatedAt && t.createdAt <= baselineCreatedAt) return; // already reflected in the baseline count
    const signed = t.type === 'income' ? t.amount : -t.amount;
    if (t.channel === 'cash') cash += signed;
    else if (t.channel === 'tng') tng += signed;
    else if (t.channel === 'bank') bank += signed;
    else { unattributedCount += 1; unattributedAmount += signed; }
  });

  const baseCash = baseline ? baseline.cash : 0;
  const baseTng = baseline ? baseline.tng : 0;
  const baseBank = baseline ? (baseline.bank || 0) : 0;

  return {
    baseline: baseline ? { date: baseline.date, cash: baseCash, tng: baseTng, bank: baseBank } : null,
    deltas: { cash, tng, bank },
    current: {
      cash: baseCash + cash,
      tng: baseTng + tng,
      bank: baseBank + bank,
      total: baseCash + cash + (baseTng + tng) + (baseBank + bank)
    },
    unattributedCount,
    unattributedAmount
  };
}

// Called right after a channel-tagged income/expense is saved, so the 정산
// (settlement) history gets a new row reflecting the up-to-date cash/TnG/
// bank split automatically — instead of requiring a fresh manual count each
// time. A transaction with no channel can't be placed in a bucket, so it's
// left out (same as in computeBalanceBreakdown) and no row is added for it.
function recordAutoSettlement(txn) {
  if (!txn || !txn.channel) return;
  const b = computeBalanceBreakdown();
  const entry = {
    id: 'settle_' + uid(),
    date: txn.date,
    cash: b.current.cash,
    tng: b.current.tng,
    bank: b.current.bank,
    total: b.current.total,
    source: 'auto',
    txnId: txn.id,
    memo: `자동 반영: ${txn.item || txn.category || ''} (${txn.type === 'income' ? '+' : '-'}${txn.amount})`,
    createdAt: new Date().toISOString()
  };
  const list = db.getSettlements();
  list.push(entry);
  db.saveSettlements(list);
}

app.get('/api/balance-breakdown', (req, res) => {
  res.json(computeBalanceBreakdown());
});

// A ready-to-fill CSV so bulk upload isn't a guessing game about column
// names/order. Opens straight into Excel/Numbers/Google Sheets.
app.get('/api/admin/bulk-template.csv', maybeRequireAdmin, (req, res) => {
  const sample = [
    BULK_TEMPLATE_HEADER.join(','),
    ['2026-11-14', '수입', '회비', '11월 회비 - 홍길동', '30', '', 'Luke'].join(','),
    ['2026-11-14', '지출', '행사비', '수련회 간식', '120.50', '다이소', 'Luke'].join(',')
  ].join('\n');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="youth-budget-template.csv"');
  res.send('﻿' + sample); // BOM so Excel opens Korean text as UTF-8, not mojibake
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
  // 현금/은행은 카페가 보고한 절대 금액을 그대로 새 기준으로 삼지만, 터치앤고는
  // 다르다 — 카페의 터치앤고 지갑 잔액을 그대로 가져오면 안 되고, 카페 정산에서
  // "이번에 얼마나 늘었는지"(증가액)만 파이낸스의 현재 터치앤고 잔액에 더해야
  // 한다. 카페 지갑과 파이낸스가 보는 터치앤고는 같은 숫자가 아니기 때문.
  const { date, cash, tngIncrease, bank, memo } = req.body || {};
  if (typeof cash !== 'number' || typeof tngIncrease !== 'number') {
    return res.status(400).json({ error: 'cash and tngIncrease must be numbers' });
  }
  const bankVal = typeof bank === 'number' ? bank : null;
  const currentTng = computeBalanceBreakdown().current.tng;
  const tng = currentTng + tngIncrease;
  const entry = {
    id: 'settle_' + uid(),
    date: date || todayStr(),
    cash,
    tng,
    bank: bankVal,
    total: cash + tng + (bankVal || 0),
    source: 'cafe-pos',
    memo: memo || `카페 정산 자동 반영 (터치앤고 +${tngIncrease})`,
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
  const { date, type, source, tag, category, item, amount, memo, author, channel } = req.body || {};
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
    channel: ['cash', 'tng', 'bank'].includes(channel) ? channel : null,
    externalId: null,
    createdAt: new Date().toISOString()
  };
  const txns = db.getTransactions();
  txns.push(entry);
  db.saveTransactions(txns);
  recordAutoSettlement(entry);
  res.status(201).json(entry);
});

// Dedicated purchase/expense entry: preset categories + optional receipt
// photo, instead of the free-text category in the generic form above.
// Shared by the admin purchase form and the public quick-entry page below -
// anyone can add a new expense (like cafe-pos's entry screen), but editing
// or deleting one is admin-only (see PUT/DELETE /api/admin/transactions/:id).
function createPurchaseEntry(req, res) {
  const { date, category, vendor, item, amount, memo, author, payment } = req.body || {};
  const amountNum = parseFloat(amount);
  if (!date || !PURCHASE_CATEGORIES.includes(category) || !item || Number.isNaN(amountNum) || amountNum <= 0 || !PAYMENT_METHODS.includes(payment)) {
    return res.status(400).json({ error: `date, category (${PURCHASE_CATEGORIES.join('|')}), item, a positive amount, and payment (${PAYMENT_METHODS.join('|')}) are required` });
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
    payment,
    channel: paymentToChannel(payment),
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
  recordAutoSettlement(entry);
  res.status(201).json(entry);
}

app.post('/api/admin/purchases', upload.single('receipt'), createPurchaseEntry);

// Public quick-entry — no admin login needed to add an expense, same as
// cafe-pos's entry screen. Only adding is open; editing/deleting stays
// behind /admin.
app.post('/api/entries/purchase', upload.single('receipt'), createPurchaseEntry);

// Dedicated income entry: preset categories (회비/후원금/헌금/기타수입)
// instead of the free-text category in the generic form above. 회비 is
// filed under source "membership" so it still groups with the existing
// membership-fee history; everything else is "manual".
app.post('/api/admin/income', (req, res) => {
  const { date, category, payer, amount, memo, author, channel } = req.body || {};
  const amountNum = parseFloat(amount);
  if (!date || !INCOME_CATEGORIES.includes(category) || Number.isNaN(amountNum) || amountNum <= 0 || !RECEIVE_CHANNELS.includes(channel)) {
    return res.status(400).json({ error: `date, category (${INCOME_CATEGORIES.join('|')}), a positive amount, and channel (${RECEIVE_CHANNELS.join('|')}) are required` });
  }
  const entry = {
    id: 'tx_' + uid(),
    date,
    type: 'income',
    source: category === '회비' ? 'membership' : 'manual',
    tag: 'youth-general',
    category,
    item: payer ? `${category} - ${payer}` : category,
    payer: payer || '',
    receiveChannel: channel,
    channel: channelToBucket(channel),
    amount: amountNum,
    memo: memo || '',
    author: author || '',
    externalId: null,
    createdAt: new Date().toISOString()
  };
  const txns = db.getTransactions();
  txns.push(entry);
  db.saveTransactions(txns);
  recordAutoSettlement(entry);
  res.status(201).json(entry);
});

// Bulk upload: .xlsx or .csv with columns 날짜/구분/카테고리/항목/금액/메모/작성자
// (see /api/admin/bulk-template.csv). Every row is validated independently —
// bad rows are reported but never silently dropped, and nothing is saved
// unless at least one row is valid, so a half-garbled file doesn't leave a
// partial, confusing mess in the ledger.
app.post('/api/admin/bulk-upload', uploadBulk.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '파일이 없습니다 (file 필드로 .xlsx 또는 .csv 업로드)' });

  const isCsv = /\.csv$/i.test(req.file.originalname || '') || req.file.mimetype === 'text/csv';
  let rows;
  try {
    if (isCsv) {
      const text = req.file.buffer.toString('utf8').replace(/^﻿/, '');
      rows = parseCsvRows(text);
    } else {
      const wb = XLSX.read(req.file.buffer, { type: 'buffer', codepage: 65001, cellDates: false });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    }
  } catch (e) {
    return res.status(400).json({ error: '파일을 읽을 수 없습니다: ' + e.message });
  }
  if (!rows.length) return res.status(400).json({ error: '비어 있는 파일입니다' });

  const TYPE_MAP = { '수입': 'income', '지출': 'expense', income: 'income', expense: 'expense' };
  const valid = [];
  const errors = [];

  rows.forEach((row, i) => {
    const rowNum = i + 2; // +1 for header row, +1 for 1-based display
    const rawDate = row['날짜'] ?? row['date'];
    const rawType = String(row['구분'] ?? row['type'] ?? '').trim();
    const category = String(row['카테고리'] ?? row['category'] ?? '').trim();
    const rawAmount = row['금액'] ?? row['amount'];
    const item = String(row['항목'] ?? row['item'] ?? '').trim();
    const memo = String(row['메모'] ?? row['memo'] ?? '').trim();
    const author = String(row['작성자'] ?? row['author'] ?? '').trim();

    const date = parseFlexibleDate(rawDate);
    const type = TYPE_MAP[rawType];
    const amount = parseFloat(rawAmount);

    if (!date) return errors.push({ row: rowNum, error: `날짜를 읽을 수 없음: "${rawDate}"` });
    if (!type) return errors.push({ row: rowNum, error: `구분은 "수입" 또는 "지출"이어야 함: "${rawType}"` });
    if (!category) return errors.push({ row: rowNum, error: '카테고리가 비어 있음' });
    if (Number.isNaN(amount) || amount <= 0) return errors.push({ row: rowNum, error: `금액이 올바르지 않음: "${rawAmount}"` });

    valid.push({
      id: 'tx_' + uid(),
      date,
      type,
      source: category === '회비' ? 'membership' : 'manual',
      tag: 'youth-general',
      category,
      item: item || category,
      amount,
      memo,
      author,
      externalId: null,
      createdAt: new Date().toISOString()
    });
  });

  if (valid.length) {
    const txns = db.getTransactions();
    db.saveTransactions(txns.concat(valid));
  }
  res.json({ inserted: valid.length, errorCount: errors.length, errors: errors.slice(0, 50) });
});

app.put('/api/admin/transactions/:id', (req, res) => {
  const txns = db.getTransactions();
  const idx = txns.findIndex((t) => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'not found' });
  const fields = ['date', 'type', 'source', 'tag', 'category', 'item', 'vendor', 'payer', 'payment', 'channel', 'amount', 'memo', 'author'];
  fields.forEach((f) => { if (req.body && req.body[f] !== undefined) txns[idx][f] = req.body[f]; });
  db.saveTransactions(txns);
  res.json(txns[idx]);
});

app.delete('/api/admin/transactions/:id', (req, res) => {
  const txns = db.getTransactions().filter((t) => t.id !== req.params.id);
  db.saveTransactions(txns);
  // Remove the 정산 row that was auto-created for this transaction too, so
  // the settlement history doesn't keep a stale entry for something that
  // no longer exists. (The live balance total is unaffected either way —
  // see computeBalanceBreakdown — this is just keeping the history honest.)
  const settlements = db.getSettlements();
  const kept = settlements.filter((s) => s.txnId !== req.params.id);
  if (kept.length !== settlements.length) db.saveSettlements(kept);
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

app.put('/api/admin/settlements/:id', (req, res) => {
  const list = db.getSettlements();
  const idx = list.findIndex((s) => s.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'not found' });
  const { date, cash, tng, bank, memo } = req.body || {};
  if (date !== undefined) list[idx].date = date;
  if (cash !== undefined) list[idx].cash = parseFloat(cash);
  if (tng !== undefined) list[idx].tng = parseFloat(tng);
  if (bank !== undefined) list[idx].bank = bank === '' || bank === null ? null : parseFloat(bank);
  if (memo !== undefined) list[idx].memo = memo;
  list[idx].total = list[idx].cash + list[idx].tng + (list[idx].bank || 0);
  db.saveSettlements(list);
  res.json(list[idx]);
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

    // One settlement row for the whole sync, not one per order, so the
    // 정산 history doesn't get flooded by a busy day at the cafe.
    if (toAdd.some((t) => t.channel)) {
      const b = computeBalanceBreakdown();
      const list = db.getSettlements();
      list.push({
        id: 'settle_' + uid(),
        date: todayStr(),
        cash: b.current.cash,
        tng: b.current.tng,
        bank: b.current.bank,
        total: b.current.total,
        source: 'auto',
        memo: `자동 반영: 카페 동기화 (${toAdd.length}건)`,
        createdAt: new Date().toISOString()
      });
      db.saveSettlements(list);
    }

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

// Excel export: no Google Cloud setup needed, just a direct file
// download, split into one sheet per source so it reads like separate
// "항목별" tabs — 전체내역 / 구매지출 / 수입 / 카페동기화 / 정산.
app.get('/api/admin/export-xlsx', (req, res) => {
  const txns = db.getTransactions().slice().sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));
  const settlements = db.getSettlements().slice().sort((a, b) => a.createdAt.localeCompare(b.createdAt));

  const TX_HEADER = ['날짜', '구분', '소스', '분류', '카테고리', '항목', '거래처', '결제/수령', '금액', '메모', '입력자', '영수증'];
  const toRow = (t) => ([
    t.date,
    t.type === 'income' ? '수입' : '지출',
    t.source,
    t.tag || '',
    t.category,
    t.item,
    t.vendor || t.payer || '',
    t.payment || t.receiveChannel || '',
    t.type === 'income' ? t.amount : -t.amount,
    t.memo || '',
    t.author || '',
    t.receiptUrl || ''
  ]);

  const purchases = txns.filter((t) => t.type === 'expense' && t.source === 'manual');
  const income = txns.filter((t) => t.type === 'income');
  const cafe = txns.filter((t) => t.source === 'cafe');

  const wb = XLSX.utils.book_new();
  const addSheet = (name, rows) => {
    const sheet = XLSX.utils.aoa_to_sheet([TX_HEADER, ...rows.map(toRow)]);
    XLSX.utils.book_append_sheet(wb, sheet, name);
  };
  addSheet('전체내역', txns);
  addSheet('구매지출', purchases);
  addSheet('수입', income);
  addSheet('카페동기화', cafe);

  const settleHeader = ['날짜', '소스', '현금', '터치앤고', '은행', '합계', '메모'];
  const settleRows = settlements.map((s) => [s.date, s.source, s.cash, s.tng, s.bank == null ? '' : s.bank, s.total, s.memo || '']);
  const settleSheet = XLSX.utils.aoa_to_sheet([settleHeader, ...settleRows]);
  XLSX.utils.book_append_sheet(wb, settleSheet, '정산');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = `주섬주섬청년부_예산_${todayStr()}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
  res.send(buf);
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
