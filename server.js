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

// 카페 재정과 청년부 재정을 완전히 분리된 두 계좌로 다룬다. 거래 데이터엔
// 이미 source(cafe/membership/manual/carryover)가 있으니 그대로 재사용:
// source==='cafe'는 카페 자신의 서랍/계좌에서 실제로 나간(들어온) 돈이고,
// 그 외는 청년부 자체 회계다. 카페 돈으로 청년모임 식사 등을 "직접" 결제한
// 지출(태그 youth-general)도 실제로는 카페 쪽 현금/계좌에서 빠져나간
// 것이므로 카페 계좌에서 그대로 차감한다 — 그 지출의 용도(카페운영 vs
// 청년부지원)는 기존 "구분별 요약" 표가 별도로 보여준다. 카페에서 청년부로
// 실제 현금/계좌이전이 일어난 경우는 /api/admin/transfer로 명시적으로
// account를 지정해 기록하므로, 거기서 저장된 값이 항상 우선한다.
function txnAccount(t) {
  if (t && (t.account === 'cafe' || t.account === 'youth')) return t.account;
  return t && t.source === 'cafe' ? 'cafe' : 'youth';
}
function settlementAccount(s) {
  if (s && (s.account === 'cafe' || s.account === 'youth')) return s.account;
  return s && s.source === 'cafe-pos' ? 'cafe' : 'youth';
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

function fmtSigned(n) {
  const v = Math.round((n || 0) * 100) / 100;
  return (v >= 0 ? '+' : '') + v;
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

// One-time fix: before today's fix, the cafe-settlement webhook took
// cafe-pos's own cash/bank numbers (its own drawer/wallet count) and used
// them as an absolute OVERWRITE of Finance's cash/bank, instead of adding
// the increase the way it already correctly did for TnG. Every settlement
// the webhook pushed under that old code is missing a `cashIncrease` field
// (today's fix always sets one), so that's how we find them. Removing them
// makes computeBalanceBreakdown() fall back to the next most recent real
// settlement and recompute the live balance from the (correctly tracked)
// transaction ledger since then — no manual number-guessing involved.
// Guarded by an env var; safe to leave on since a second run finds nothing
// left to remove.
if (process.env.CLEANUP_BAD_CAFE_SETTLEMENTS_20261116 === '1') {
  const __settlements3 = db.getSettlements();
  const __bad = __settlements3.filter((s) => s.source === 'cafe-pos' && typeof s.cashIncrease !== 'number');
  if (__bad.length) {
    const __kept = __settlements3.filter((s) => !(s.source === 'cafe-pos' && typeof s.cashIncrease !== 'number'));
    db.saveSettlements(__kept);
    console.log('[youth-budget] removed', __bad.length, 'pre-fix cafe-pos settlement(s) that had overwritten cash/bank:', __bad.map((s) => s.id));
  }
}

// 집사님이 정리해서 넘겨준 이월 잔액(source==='carryover')은 사실 카페 쪽
// 누적 자금이지 청년부 자체 돈이 아니므로, 카페/청년부 계좌 분리 이후엔
// 카페 계좌로 분류해야 한다. 이미 account:'cafe'로 지정된 건 다시
// 건드리지 않으므로(멱등) 매 배포/재시작마다 돌아도 안전 — 별도 env 게이트
// 없이 둔다.
(function reclassifyDeaconCarryoverAsCafe() {
  const txns = db.getTransactions();
  let changed = 0;
  txns.forEach((t) => {
    if (t.source === 'carryover' && t.account !== 'cafe') {
      t.account = 'cafe';
      changed += 1;
    }
  });
  if (changed) {
    db.saveTransactions(txns);
    console.log('[youth-budget] 집사님 이월 잔액(carryover) 거래', changed, '건을 카페 계좌로 재분류했습니다.');
  }
})();

// "소스가 카페면 카페 제정이야" — cafe-pos에서 동기화된(source==='cafe')
// 거래는 예외 없이 카페 계좌(account='cafe') · 카페 운영(tag='cafe-ops')
// 이어야 한다. 실제 카페→청년부 이체는 별도 source('transfer')로 기록되므로
// (/api/admin/transfer 참고), source='cafe'인데 account가 'youth'로 덮어써진
// 건은 전부 과거의 수동 오분류다. 이미 올바른 건은 다시 건드리지 않으므로
// (멱등) 별도 env 게이트 없이 매 배포/재시작마다 돌아도 안전.
(function forceCafeSourceToCafeAccount() {
  const txns = db.getTransactions();
  let accountChanged = 0;
  let tagChanged = 0;
  txns.forEach((t) => {
    if (t.source === 'cafe') {
      if (t.account !== 'cafe') { t.account = 'cafe'; accountChanged += 1; }
      if (t.tag !== 'cafe-ops') { t.tag = 'cafe-ops'; tagChanged += 1; }
    }
  });
  if (accountChanged || tagChanged) {
    db.saveTransactions(txns);
    console.log(`[youth-budget] source=cafe 거래 중 계좌 재지정 ${accountChanged}건, 태그 재지정(→카페 운영) ${tagChanged}건 완료.`);
  }
})();

// 읽기 전용 진단: "청년부 실제 잔액 구성"이 비정상적으로 커 보이는 원인과,
// 청년부 계정으로 잘못 들어가 있을 수 있는 카페 관련 수동입력 거래를
// 찾기 위해 현재 정산/거래 데이터를 로그로만 출력한다 (아무것도 바꾸지
// 않음). DIAG_ACCOUNTS_20261120=1일 때만, 매 배포/재시작마다 실행.
if (process.env.DIAG_ACCOUNTS_20261120 === '1') {
  const allSettlements = db.getSettlements().slice().sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  console.log('[diag] ===== 전체 정산(settlements) 목록 (' + allSettlements.length + '건) =====');
  allSettlements.forEach((s) => {
    console.log(`[diag] ${s.createdAt} | date=${s.date} | account=${settlementAccount(s)} (raw=${s.account || ''}) | source=${s.source} | cash=${s.cash} tng=${s.tng} bank=${s.bank} total=${s.total} | id=${s.id} | memo=${(s.memo || '').slice(0, 60)}`);
  });

  console.log('[diag] ===== youth / cafe breakdown =====');
  console.log('[diag] youth:', JSON.stringify(computeAccountBreakdown('youth')));
  console.log('[diag] cafe: ', JSON.stringify(computeAccountBreakdown('cafe')));

  // account='youth'로 잡히는 수동입력(manual/membership) 거래를 작성자별로
  // 묶어서 건수/합계/카테고리를 보여준다 — 집사님이 입력한 카페 관련 건을
  // 찾아내기 위함. cafe-pos 자동동기화(source='cafe') 건은 이미 account
  // 분리가 되어 있으므로 제외.
  const allTxns = db.getTransactions();
  const candidates = allTxns.filter((t) => txnAccount(t) === 'youth' && t.source !== 'cafe' && t.source !== 'channel-transfer' && t.source !== 'transfer');
  const byAuthor = {};
  candidates.forEach((t) => {
    const key = t.author || '(미기재)';
    if (!byAuthor[key]) byAuthor[key] = { count: 0, income: 0, expense: 0, categories: new Set() };
    byAuthor[key].count += 1;
    byAuthor[key][t.type] += t.amount;
    byAuthor[key].categories.add(t.category || '');
  });
  console.log('[diag] ===== account=youth 수동/회비 거래, 작성자별 집계 =====');
  Object.entries(byAuthor).forEach(([author, v]) => {
    console.log(`[diag] 작성자="${author}" | 건수=${v.count} | 수입합=${v.income} | 지출합=${v.expense} | 카테고리=[${[...v.categories].join(', ')}]`);
  });

  // channel이 비어 있는(=현금/TnG/은행 어디에도 못 넣는) account=youth 거래
  // 목록 — "실제 잔액 구성"엔 안 잡히고 unattributedAmount로만 잡히는 건들.
  const noChannel = candidates.filter((t) => !t.channel);
  console.log('[diag] ===== account=youth, channel 없는 거래 ' + noChannel.length + '건 (최근 15개만) =====');
  noChannel.slice(-15).forEach((t) => {
    console.log(`[diag]   ${t.date} | ${t.type} | ${t.category} | ${t.item} | ${t.amount} | author=${t.author || ''} | memo=${(t.memo || '').slice(0, 40)}`);
  });

  // cafe-pos에서 동기화된(source='cafe') 거래 — 태그(카페운영/청년부지출/
  // 확인필요)별, 그리고 그 중 account가 명시적으로 'youth'로 덮어써진 건이
  // 있는지. "집사님이 입력한 카페 지출/입금이 청년부 일반지출로 돼있다"는
  // 말이 이 데이터(tag 분류) 얘기인지 확인하기 위함.
  const cafeSynced = allTxns.filter((t) => t.source === 'cafe');
  const byTag2 = {};
  cafeSynced.forEach((t) => {
    const key = t.tag || '(태그없음)';
    if (!byTag2[key]) byTag2[key] = { count: 0, expense: 0, income: 0, overriddenToYouth: 0 };
    byTag2[key].count += 1;
    byTag2[key][t.type] += t.amount;
    if (t.account === 'youth') byTag2[key].overriddenToYouth += 1;
  });
  console.log('[diag] ===== source=cafe(cafe-pos 동기화) 거래 ' + cafeSynced.length + '건, 태그별 집계 =====');
  Object.entries(byTag2).forEach(([tag, v]) => {
    console.log(`[diag] 태그="${tag}" | 건수=${v.count} | 수입합=${v.income} | 지출합=${v.expense} | account=youth로 덮어쓴 건=${v.overriddenToYouth}`);
  });

  const youthGeneralCafe = cafeSynced.filter((t) => t.tag === 'youth-general');
  const byAuthor2 = {};
  youthGeneralCafe.forEach((t) => {
    const key = t.author || t.vendor || '(미기재)';
    if (!byAuthor2[key]) byAuthor2[key] = { count: 0, expense: 0, income: 0 };
    byAuthor2[key].count += 1;
    byAuthor2[key][t.type] += t.amount;
  });
  console.log('[diag] ===== source=cafe & tag=youth-general(청년부지출로 분류된 카페 거래) ' + youthGeneralCafe.length + '건, 작성자/거래처별 =====');
  Object.entries(byAuthor2).forEach(([author, v]) => {
    console.log(`[diag] 작성자/거래처="${author}" | 건수=${v.count} | 수입합=${v.income} | 지출합=${v.expense}`);
  });
  console.log('[diag] ===== 위 목록 샘플 10건 =====');
  youthGeneralCafe.slice(0, 10).forEach((t) => {
    console.log(`[diag]   ${t.date} | ${t.type} | ${t.item} | ${t.amount} | author=${t.author || ''} | vendor=${t.vendor || ''} | memo=${(t.memo || '').slice(0, 40)}`);
  });
}

// 읽기 전용 진단: 임창현님 개인 "회비 납부 내역" 엑셀(회비 수납 74건 +
// 일반운영 지출 26건)을 건별로 앱에 반영하기 전에, 이미 수동으로 입력돼
// 있어 중복될 만한 건이 있는지 확인하기 위해 account=youth인 수동/회비
// 거래를 전부 날짜순으로 덤프한다 (cafe/채널전환/수련회시드는 제외).
// DIAG_YOUTH_LEDGER_20261120=1일 때만, 아무것도 바꾸지 않음.
if (process.env.DIAG_YOUTH_LEDGER_20261120 === '1') {
  const allTxns2 = db.getTransactions();
  const ledger = allTxns2
    .filter((t) => txnAccount(t) === 'youth' && t.source !== 'cafe' && t.source !== 'channel-transfer' && !(t.externalId || '').startsWith('retreat1_'))
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date));
  console.log('[diag] ===== account=youth, 수련회시드 제외 거래 ' + ledger.length + '건 (날짜순) =====');
  ledger.forEach((t) => {
    console.log(`[diag]   ${t.date} | ${t.type} | ${t.category} | ${t.item || ''} | ${t.payer || ''} | ${t.amount} | source=${t.source} | author=${t.author || ''} | id=${t.id}`);
  });
}

// 청년부 결산용 — account=youth 전체 거래를 날짜순으로 덤프(채널전환
// 내부조정만 제외, 수련회시드 포함). 아무것도 바꾸지 않음.
// DIAG_YOUTH_SETTLEMENT_LIST_20261121=1일 때만.
if (process.env.DIAG_YOUTH_SETTLEMENT_LIST_20261121 === '1') {
  const allTxns3 = db.getTransactions();
  const ledger3 = allTxns3
    .filter((t) => txnAccount(t) === 'youth' && t.source !== 'channel-transfer')
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date));
  let inc = 0, exp = 0;
  console.log('[diag2] ===== account=youth 결산용 전체 거래 ' + ledger3.length + '건 (날짜순) =====');
  ledger3.forEach((t) => {
    if (t.type === 'income') inc += t.amount; else exp += t.amount;
    console.log(`[diag2]   ${t.date} | ${t.type} | ${t.category} | ${t.item || ''} | ${t.payer || ''} | ${t.amount} | source=${t.source} | tag=${t.tag || ''}`);
  });
  console.log(`[diag2] ===== 합계: 수입 ${inc} / 지출 ${exp} / 순액 ${inc - exp} =====`);
}

// 1차 수련회 잔액(2,557.18) 중 부장집사님이 현금으로 보관 중이던
// RM600을, 카페가 대신 청년부 계좌로 채워넣는 것으로 기록한다
// (/api/admin/transfer와 동일한 구조: 카페=지출, 청년부=수입, 날짜는
// 수련회 날짜 2026-04-30). 카페는 -600, 청년부는 +600, 전체(카페+청년부)
// 합계는 변동 없음(계좌 간 이동일 뿐). transferId 고정 — 멱등.
if (process.env.ADD_DEACON_600_CAFE_TRANSFER_20261121 === '1') {
  const txns7 = db.getTransactions();
  const transferId7 = 'xfer_deacon600_20261121';
  if (!txns7.some((t) => t.transferId === transferId7)) {
    const now7 = new Date().toISOString();
    const memo7 = '1차 수련회 잔액 중 부장집사님이 현금 보관 중이던 RM600을 카페가 대신 청년부 계좌로 채워넣음(수련회 날짜 기준 소급)';
    const outTxn7 = {
      id: 'tx_' + uid(), date: '2026-04-30', type: 'expense', source: 'transfer',
      account: 'cafe', tag: null, category: '청년부 후원', item: '청년부 후원/이체',
      amount: 600, channel: channelToBucket('계좌이체'), memo: memo7, author: '주섬주섬청년부',
      transferId: transferId7, externalId: null, createdAt: now7
    };
    const inTxn7 = {
      id: 'tx_' + uid(), date: '2026-04-30', type: 'income', source: 'transfer',
      account: 'youth', tag: 'youth-general', category: '카페 후원금', item: '카페 후원금',
      amount: 600, channel: channelToBucket('계좌이체'), memo: memo7, author: '주섬주섬청년부',
      transferId: transferId7, externalId: null, createdAt: now7
    };
    txns7.push(outTxn7, inTxn7);
    db.saveTransactions(txns7);
    recordAutoSettlement(outTxn7);
    recordAutoSettlement(inTxn7);
    console.log('[youth-budget] applied ADD_DEACON_600_CAFE_TRANSFER_20261121: 카페 -600 / 청년부 +600 (2026-04-30)');
  }
}

// 테스트로 만든 RM100 현금→계좌이체 전환(계좌 전환 탭 테스트) 삭제.
// /api/admin/transactions/:id DELETE와 동일한 로직 — 짝 거래 + 관련
// auto 정산 스냅샷도 같이 지운다. 멱등 — 해당 transferId가 없으면 통과.
if (process.env.DELETE_TEST_CHANNEL_TRANSFER_20261121 === '1') {
  const txns8 = db.getTransactions();
  const idsToRemove8 = new Set(
    txns8.filter((t) => t.transferId === 'chxfer_mwp739k789ihdk').map((t) => t.id)
  );
  if (idsToRemove8.size) {
    db.saveTransactions(txns8.filter((t) => !idsToRemove8.has(t.id)));
    const settlements8 = db.getSettlements();
    const kept8 = settlements8.filter((s) => !idsToRemove8.has(s.txnId));
    if (kept8.length !== settlements8.length) db.saveSettlements(kept8);
    console.log('[youth-budget] applied DELETE_TEST_CHANNEL_TRANSFER_20261121: ' + idsToRemove8.size + '건 삭제(테스트 RM100 계좌 전환)');
  }
}

// transfer/channel-transfer 쌍이 날짜·계좌·금액까지 똑같이 겹치는 게
// 있는지(실수로 두 번 기록됐는지) 확인한다. 아무것도 바꾸지 않음.
// DIAG_CHECK_TRANSFER_DUPES_20261121=1일 때만.
if (process.env.DIAG_CHECK_TRANSFER_DUPES_20261121 === '1') {
  const all7 = db.getTransactions();
  const pairs7 = all7.filter((t) => (t.source === 'transfer' || t.source === 'channel-transfer') && t.type === 'expense');
  console.log('[diag5] ===== transfer/channel-transfer 쌍 ' + pairs7.length + '건 =====');
  pairs7.sort((a, b) => a.date.localeCompare(b.date)).forEach((t) => {
    const sibling = all7.find((x) => x.transferId === t.transferId && x.id !== t.id);
    console.log(`[diag5]   ${t.date} | source=${t.source} | account=${txnAccount(t)} | item="${t.item}" | ${t.amount} | channel=${t.channel}->${sibling ? sibling.channel : '?'} | transferId=${t.transferId} | txId=${t.id} | createdAt=${t.createdAt}`);
  });
  const key = (t) => `${t.date}|${txnAccount(t)}|${t.source}|${t.amount}|${t.channel}`;
  const byKey = {};
  pairs7.forEach((t) => { (byKey[key(t)] = byKey[key(t)] || []).push(t); });
  const dupeKeys = Object.entries(byKey).filter(([, v]) => v.length > 1);
  console.log('[diag5] 날짜/계좌/소스/금액/채널까지 겹치는 그룹: ' + dupeKeys.length + '건');
  dupeKeys.forEach(([k, v]) => console.log('[diag5]   ' + k + ' -> ' + v.map((t) => t.id + '(transferId=' + t.transferId + ', createdAt=' + t.createdAt + ')').join(', ')));
}

// 윤수빈/임혜란 집사님 관련 거래를 전부 찾는다. 아무것도 바꾸지 않음.
// DIAG_FIND_YOONSUBIN_20261121=1일 때만.
if (process.env.DIAG_FIND_YOONSUBIN_20261121 === '1') {
  const all6 = db.getTransactions();
  const hits6 = all6.filter((t) =>
    String(t.payer || '').includes('윤수빈') ||
    String(t.item || '').includes('윤수빈') ||
    String(t.memo || '').includes('윤수빈') ||
    String(t.payer || '').includes('임혜란') ||
    String(t.item || '').includes('임혜란') ||
    String(t.memo || '').includes('임혜란')
  ).sort((a, b) => a.date.localeCompare(b.date));
  console.log('[diag4] ===== 윤수빈/임혜란 관련 거래 ' + hits6.length + '건 =====');
  hits6.forEach((t) => {
    console.log(`[diag4]   ${t.date} | ${t.type} | account=${txnAccount(t)} | ${t.category} | item="${t.item || ''}" | payer="${t.payer || ''}" | ${t.amount} | source=${t.source} | memo="${t.memo || ''}" | id=${t.id}`);
  });
}

// "집사님" 관련 거래(전체 소스/계좌 불문) + 금액 600 거래를 전부 찾는다.
// 아무것도 바꾸지 않음. DIAG_FIND_DEACON_600_20261121=1일 때만.
if (process.env.DIAG_FIND_DEACON_600_20261121 === '1') {
  const all5 = db.getTransactions();
  const hits = all5.filter((t) =>
    String(t.payer || '').includes('집사님') ||
    String(t.author || '').includes('집사님') ||
    String(t.item || '').includes('집사님') ||
    t.amount === 600
  ).sort((a, b) => a.date.localeCompare(b.date));
  console.log('[diag3] ===== "집사님" 관련 또는 금액=600 거래 ' + hits.length + '건 =====');
  hits.forEach((t) => {
    console.log(`[diag3]   ${t.date} | ${t.type} | account=${txnAccount(t)} | ${t.category} | ${t.item || ''} | ${t.payer || ''} | ${t.amount} | source=${t.source} | tag=${t.tag || ''} | author=${t.author || ''} | id=${t.id}`);
  });
  const settlementsAll = db.getSettlements();
  const sHits = settlementsAll.filter((s) => String(s.memo || '').includes('집사님') || s.total === 600);
  console.log('[diag3] 관련 정산: ' + sHits.length + '건');
  sHits.forEach((s) => console.log('[diag3]   ' + JSON.stringify(s)));
}

// 집사님이 정리해서 넘겨준 과거 카페 지출입금("집사님(정리본)") 53건이
// tag=youth-general("청년부 일반지출")로 들어가 있던 걸 cafe-ops("카페
// 운영")로 바꾼다. 내용은 실제로 청년부를 위해 쓴 카페 돈(식사 지원 등)이
// 맞지만, 이미 카페 쪽 정산으로 마감된 금액이라 청년부 자체 장부와는
// 분리해야 한다는 판단 — account는 원래도 'cafe'라 안 바뀜, tag만 바뀜.
// 멱등 — 조건(source=cafe, tag=youth-general, 해당 작성자)에 맞는 게
// 없으면 더 이상 할 일이 없으므로 재실행해도 안전.
if (process.env.RETAG_DEACON_CAFE_EXPENSES_20261120 === '1') {
  const txns = db.getTransactions();
  let changed = 0;
  txns.forEach((t) => {
    if (t.source === 'cafe' && t.tag === 'youth-general' && t.author === '집사님(정리본)') {
      t.tag = 'cafe-ops';
      changed += 1;
    }
  });
  if (changed) {
    db.saveTransactions(txns);
    console.log('[youth-budget] applied RETAG_DEACON_CAFE_EXPENSES_20261120 migration:', changed, '건을 cafe-ops로 재태깅');
  }
}

// One-time seed: 2026년 4/30~5/2 "1차 수련회" 재정 내역 (별도로 관리되던
// 엑셀 파일을 Finance로 이전). 회비/후원은 사람별로 각 건씩, 구매 내역은
// 88건을 하나로 뭉쳐서 "행사비 - 1차 수련회"로 반영한다 ("ⓘ 리더단 회의비"
// 7건/RM348은 엑셀에도 "별도항목/후원 포함 x"로 표시돼 있어 수련회 예산에서
// 제외). 멱등 — externalId가 'retreat1_'로 시작하는 거래가 이미 있으면 다시
// 넣지 않으므로 재배포/재시작에도 안전.
if (process.env.SEED_RETREAT1_20261120 === '1') {
  const __txns = db.getTransactions();
  if (!__txns.some((t) => t.externalId && t.externalId.startsWith('retreat1_'))) {
    const now = new Date().toISOString();
    const mkIncome = (externalId, date, category, payer, amount, receiveChannel, memo) => ({
      id: 'tx_' + uid(),
      date,
      type: 'income',
      source: category === '회비' ? 'membership' : 'manual',
      account: 'youth',
      tag: 'youth-general',
      category,
      item: `${category} - ${payer}`,
      payer,
      receiveChannel,
      channel: channelToBucket(receiveChannel),
      amount,
      memo: memo || '',
      author: '임창현',
      externalId,
      createdAt: now
    });
    // 회비 (20명 중 박해령 1명은 100 납부 후 전액 환불되어 순액 0 — 반영 생략)
    const FEE_ROWS = [
      ['2026-04-08', '황나엘', 150, 'QR'],
      ['2026-04-11', '오찬영', 100, 'QR', '1박2일'],
      ['2026-04-13', '강은수', 125, 'QR'],
      ['2026-04-16', '한윤서', 125, 'QR'],
      ['2026-04-18', '윤수빈', 125, 'QR'],
      ['2026-04-18', '전용준 집사님', 125, 'QR'],
      ['2026-04-18', '홍서희', 125, 'QR'],
      ['2026-04-18', '전소윤', 125, 'Cash'],
      ['2026-04-18', '심민석', 125, 'Cash'],
      ['2026-04-18', '원가온', 125, 'Cash'],
      ['2026-04-19', '임창현', 150, 'QR'],
      ['2026-04-19', '김정음', 150, 'QR'],
      ['2026-04-19', '김우제', 150, 'Cash'],
      ['2026-04-20', '김애스더', 100, 'QR', '1박2일'],
      ['2026-04-22', '김보민', 100, 'QR', '1박2일'],
      ['2026-04-22', '김효찬', 150, 'QR'],
      ['2026-04-26', '최원준', 150, 'QR'],
      ['2026-04-27', '서연우', 150, 'QR'],
      ['2026-04-28', '박규리', 150, 'QR']
    ];
    const feeTxns = FEE_ROWS.map((row, i) => {
      const [date, payer, amount, method, memo] = row;
      return mkIncome(`retreat1_fee_${i + 1}`, date, '회비', payer, amount, method === 'QR' ? '계좌이체' : '현금', memo);
    });

    // 후원 (날짜가 엑셀에 없어 수련회 기간 첫날 날짜로, 결제수단도 명시가
    // 없어 계좌이체로 가정 — 실제와 다르면 나중에 수정 가능)
    const SPONSOR_ROWS = [
      ['김태준', 500],
      ['바자회', 650, '여전도회'],
      ['지경섭 집사님', 150],
      ['고수영 목사님', 1000],
      ['서연우 자매 어머니', 800, '금액 확인 필요(엑셀 원본 메모)'],
      ['임창현 친구', 150],
      ['정진석 집사님', 200],
      ['엄철규 장로님', 500],
      ['이현택 장로님', 300],
      ['최민혁 장로님', 500],
      ['류은영 권사님', 300]
    ];
    const sponsorTxns = SPONSOR_ROWS.map((row, i) => {
      const [payer, amount, memo] = row;
      return mkIncome(`retreat1_sponsor_${i + 1}`, '2026-04-30', '후원금', payer, amount, '계좌이체', memo);
    });
    const churchTxn = mkIncome('retreat1_church', '2026-04-30', '후원금', '교회(지원금)', 2000, '계좌이체', '교회 지원금');

    // 구매 내역 88건을 하나로 뭉침 (리더단 회의비 7건/RM348은 제외)
    const expenseTxn = {
      id: 'tx_' + uid(),
      date: '2026-04-30',
      type: 'expense',
      source: 'manual',
      account: 'youth',
      tag: 'youth-general',
      category: '행사비',
      item: '1차 수련회',
      vendor: '',
      payment: '카드',
      channel: channelToBucket('계좌이체'),
      amount: 6992.82,
      memo: '1차 수련회(4/30~5/2) 구매 내역 88건 합산 (숙박/식비/액티비티/조모임 지원금 등). "리더단 회의비" 7건(RM348)은 엑셀 원본에도 별도항목/후원 포함 제외로 표시되어 제외함.',
      author: '임창현',
      receiptUrl: null,
      externalId: 'retreat1_expense',
      createdAt: now
    };

    const allNew = [...feeTxns, ...sponsorTxns, churchTxn, expenseTxn];
    db.saveTransactions(__txns.concat(allNew));

    const b = computeAccountBreakdown('youth');
    const settlements = db.getSettlements();
    settlements.push({
      id: 'settle_retreat1_20261120',
      date: '2026-04-30',
      account: 'youth',
      cash: b.current.cash,
      tng: b.current.tng,
      bank: b.current.bank,
      total: b.current.total,
      source: 'manual',
      memo: '1차 수련회 재정 반영 (회비 19건/후원 11건+교회지원금/구매내역 합산 1건)',
      createdAt: now
    });
    db.saveSettlements(settlements);
    console.log('[youth-budget] applied SEED_RETREAT1_20261120 migration:', allNew.length, 'transactions added');
  }
}

// One-time seed: 임창현님 개인 "회비 납부 내역" 엑셀(수련회와 별개로, 4~11월
// 매달 걷은 회비 + 그 돈으로 치른 토요모임/아웃팅 등 일반 운영비) — "내
// 개인 운영장부가 즉 청년부 재정 장부"라는 확인에 따라 전부 account='youth'
// 수동 건별 입력으로 반영한다. 날짜는 실제 거래일 그대로, 회비 성격이
// 아닌 2건(카페 아웃팅지원 2000, 윤수빈 편 임혜란 집사님 후원 300)과
// 회비·후원이 섞여 있던 1건(여사무엘 150 = 회비 50 + 수련회후원 100)은
// 후원금으로 성격을 바로잡아 분리했다. 현금 유무 안내("현금을 은행으로
// 다 옮길 걸로 쳐줘")에 따라 QR/Cash 구분 없이 전부 계좌이체(은행)로
// 반영. 이미 수동으로 들어가 있던 2건(11/14 돈까스랩 348, 11/18 꽃구매
// 98)은 중복 방지를 위해 제외. 멱등 — externalId가 'youthledger_'로
// 시작하는 거래가 이미 있으면 다시 넣지 않음.
if (process.env.SEED_YOUTH_LEDGER_20261120 === '1') {
  const __txns2 = db.getTransactions();
  if (!__txns2.some((t) => t.externalId && t.externalId.startsWith('youthledger_'))) {
    const now2 = new Date().toISOString();
    const mkI = (externalId, date, category, payer, amount, memo) => ({
      id: 'tx_' + uid(),
      date,
      type: 'income',
      source: category === '회비' ? 'membership' : 'manual',
      account: 'youth',
      tag: 'youth-general',
      category,
      item: `${category} - ${payer}`,
      payer,
      receiveChannel: '계좌이체',
      channel: 'bank',
      amount,
      memo: memo || '',
      author: '임창현',
      externalId,
      createdAt: now2
    });
    const mkE = (externalId, date, category, item, amount, memo) => ({
      id: 'tx_' + uid(),
      date,
      type: 'expense',
      source: 'manual',
      account: 'youth',
      tag: 'youth-general',
      category,
      item,
      vendor: '',
      payment: '카드',
      channel: 'bank',
      amount,
      memo: memo || '',
      author: '임창현',
      receiptUrl: null,
      externalId,
      createdAt: now2
    });

    // 회비 (72건) — 엑셀 원본의 QR/Cash 구분은 전부 은행으로 통일
    const FEE_ROWS2 = [
      ['2026-04-03', '김보민', 50], ['2026-04-04', '김정국', 50], ['2026-04-04', '핼리', 50],
      ['2026-04-04', '여사무엘', 50], ['2026-04-04', '오찬영', 50], ['2026-04-08', '황나엘', 50],
      ['2026-04-18', '윤수빈', 50], ['2026-04-18', '전용준 집사님', 50], ['2026-04-18', '배현아', 50],
      ['2026-04-18', '강은수', 50], ['2026-04-18', '전소윤', 50], ['2026-04-18', '원가온', 50],
      ['2026-04-19', '김정음', 50], ['2026-04-19', '임창현', 50], ['2026-04-19', '김우제', 50],
      ['2026-04-26', '최원준', 50],
      ['2026-05-15', '오찬영', 50], ['2026-05-15', '박세희', 50], ['2026-05-16', '윤수빈', 50],
      ['2026-05-16', '홍서희', 50], ['2026-05-16', '여사무엘', 50], ['2026-05-16', '임창현', 50],
      ['2026-05-23', '한윤서', 50], ['2026-05-23', '김영규', 50], ['2026-05-23', '최원준', 50],
      ['2026-05-23', '강은수', 50], ['2026-05-30', '원가온', 50], ['2026-06-28', '배현아', 50],
      ['2026-06-25', '원가온', 50], ['2026-06-26', '김보민', 50], ['2026-06-26', '황나엘', 50],
      ['2026-06-27', '홍서희', 50], ['2026-06-28', '김영규', 50],
      ['2026-07-04', '배현아', 50], ['2026-07-04', '박규리', 50], ['2026-07-04', '윤수빈', 50],
      ['2026-07-04', '박세희', 50], ['2026-07-06', '김우제', 100], ['2026-07-06', '김정음', 150],
      ['2026-07-06', '최원준', 50], ['2026-07-06', '임창현', 50],
      ['2026-07-23', '이재아', 100], ['2026-07-23', '오찬영', 50], ['2026-07-23', '박세희', 50],
      ['2026-07-23', '유재신', 50], ['2026-07-23', '서제민', 50], ['2026-07-26', '조하영', 100],
      ['2026-07-26', '김우제', 50], ['2026-07-26', '원가온', 50], ['2026-07-27', '한윤서', 50],
      ['2026-07-27', '임창현', 50], ['2026-08-01', '윤수빈', 50], ['2026-08-01', '황나엘', 50],
      ['2026-09-02', '한윤서', 50], ['2026-09-02', '임창현', 50],
      ['2026-09-20', '임창현', 50], ['2026-09-29', '김보민', 50], ['2026-10-01', '윤수빈', 50],
      ['2026-10-03', '한윤서', 50],
      ['2026-10-10', '조하영', 50], ['2026-10-10', '황나엘', 100], ['2026-11-01', '황나엘', 50],
      ['2026-10-29', '윤수빈', 120], ['2026-11-01', '문혜성', 70], ['2026-11-01', '김보민', 70],
      ['2026-11-01', '정햇살', 70], ['2026-11-04', '원가온', 70], ['2026-11-07', '최원준', 70],
      ['2026-11-08', '조한나', 50], ['2026-11-09', '전소윤', 70], ['2026-11-09', '김근태', 70],
      ['2026-11-09', '전용준', 70]
    ];
    const feeTxns2 = FEE_ROWS2.map((row, i) => {
      const [date, payer, amount] = row;
      return mkI(`youthledger_fee_${i + 1}`, date, '회비', payer, amount);
    });

    // 성격 수정: 회비가 아니라 후원금인 3건
    const sponsorTxns2 = [
      mkI('youthledger_sponsor_1', '2026-05-16', '후원금', '여사무엘', 100, '수련회 후원 (원본 150 중 회비 50 분리, 나머지 100)'),
      mkI('youthledger_sponsor_2', '2026-10-25', '후원금', '카페', 2000, '아웃팅 지원'),
      mkI('youthledger_sponsor_3', '2026-11-09', '후원금', '임혜란 집사님', 300, '윤수빈 편으로 전달')
    ];

    // 일반 운영 지출 23건 (엑셀 25건 중 이미 수동입력돼 있던 2건
    // (11/14 돈까스랩 348, 11/18 꽃구매·전용준집사 98)은 중복이라 제외)
    const EXPENSE_ROWS2 = [
      ['2026-05-16', '식비·다과', '토요모임 - 5/16 저녁 쿠시야', 477.0],
      ['2026-05-30', '행사비', '액티비티 - 배드민턴 대여', 66.78],
      ['2026-05-30', '행사비', '액티비티 - 배드민턴 채 대여', 40.0],
      ['2026-05-30', '식비·다과', '액티비티 - 음료수 구매', 51.9],
      ['2026-05-30', '식비·다과', '액티비티 - 밥값 / 스끼시', 558.9],
      ['2026-06-20', '행사비', '외부강사 - 여경진 목사님 방문', 300.0],
      ['2026-07-04', '식비·다과', '토요모임 - 7/4 저녁 만나도시락', 442.0],
      ['2026-07-11', '식비·다과', '토요모임 - 솥밥', 714.55],
      ['2026-07-12', '행사비', '집들이 - 찬영&우제', 200.0],
      ['2026-09-06', '식비·다과', '일요모임 - 청년집회 / 왕족족발', 452.4],
      ['2026-09-19', '식비·다과', '토요모임 - 저녁 만나도시락', 322.0],
      ['2026-10-30', '행사비', '아웃팅 - 에어비앤비(예약번호 HMZSMZJCRT)', 1030.0],
      ['2026-11-07', '식비·다과', '토요모임 - 만나도시락', 297.0],
      ['2026-11-07', '행사비', '아웃팅 - 빌리지그로서', 307.7],
      ['2026-11-07', '행사비', '아웃팅 - 이온 슈퍼마켓(타만말루리점)', 74.3],
      ['2026-11-07', '행사비', '아웃팅 - KK수퍼마트(겐팅점)', 20.9],
      ['2026-11-07', '행사비', '아웃팅 - 맥도날드(페트론 주유소점) / 겐팅하이랜드', 87.15],
      ['2026-11-09', '행사비', '아웃팅 - 최원준/2명', 28.0],
      ['2026-11-09', '행사비', '아웃팅 - 한윤서/2명', 28.0],
      ['2026-11-09', '행사비', '아웃팅 - 임창현/3명', 42.0],
      ['2026-11-09', '행사비', '아웃팅 - 전용준/4명', 56.0],
      ['2026-11-09', '식비·다과', '아웃팅_점심 - 수끼야', 558.9],
      ['2026-11-15', '기타', '꽃 구매 - 정햇살/입교', 108.0]
    ];
    const expenseTxns2 = EXPENSE_ROWS2.map((row, i) => {
      const [date, category, item, amount] = row;
      return mkE(`youthledger_expense_${i + 1}`, date, category, item, amount);
    });

    const allNew2 = [...feeTxns2, ...sponsorTxns2, ...expenseTxns2];
    db.saveTransactions(__txns2.concat(allNew2));
    console.log('[youth-budget] applied SEED_YOUTH_LEDGER_20261120 migration:', allNew2.length, 'transactions added',
      `(회비 ${feeTxns2.length} / 후원금 ${sponsorTxns2.length} / 지출 ${expenseTxns2.length})`);
  }
}

// "청년부 재정은 현금이 없고 다 은행이다" / "현금을 은행으로 다 옮길 걸로
// 쳐줘" / "다 현금 말고 은행으로 잡아줘" — 청년부 계좌(account='youth')
// 거래 중 channel이 'cash'(현금)·'tng'(터치앤고)로 잡혀 있는 건(주로
// 수련회 회비 중 Cash로 낸 건들)을 전부 bank(은행)로 바꾼다. 표시용
// 라벨(receiveChannel/payment)도 같이 맞춘다. 카페 계좌는 실제로 카페
// 서랍에 현금이 있으므로 건드리지 않는다. 위의 수련회/회비장부 시드보다
// 반드시 뒤에 와야 한다(새로 시드되는 Cash 건까지 다 잡아야 하므로).
// 이미 bank인 건 다시 안 건드리므로(멱등) env 게이트 없이 상시 적용.
(function forceYouthChannelToBank() {
  const txns = db.getTransactions();
  let changed = 0;
  txns.forEach((t) => {
    if (txnAccount(t) === 'youth' && (t.channel === 'cash' || t.channel === 'tng')) {
      t.channel = 'bank';
      if (t.receiveChannel) t.receiveChannel = '계좌이체';
      if (t.payment) t.payment = '카드';
      changed += 1;
    }
  });
  if (changed) {
    db.saveTransactions(txns);
    console.log(`[youth-budget] 청년부 계좌 거래 중 현금/TnG로 잡혀있던 ${changed}건을 은행(bank)으로 재지정했습니다.`);
  }
})();

// "청년부는 전체가 계좌에 있어 / 내가 그때그때 계좌로 넣어" — "실제 잔액
// 구성"에 남아있던 현금/터치앤고 분량은, 과거(계좌분리 이전) 기준점
// 정산이 물려준 숫자일 뿐 실제로는 임창현님이 그때그때 다 은행에 입금해둔
// 돈이라는 확인. 총액(RM)은 그대로 두고 현금/TnG를 전부 계좌로 합쳐
// 새 기준점을 오늘 날짜로 다시 찍는다 — 총액을 추측/변경하지 않고,
// 지금 이 순간의 실제 잔액 구성 계산값을 그대로 가져와 cash/tng만
// bank로 옮겨 담는다. 멱등 — 고정 id(settle_youth_allbank_20261120)가
// 이미 있으면 다시 안 만듦.
if (process.env.RESET_YOUTH_BASELINE_TO_BANK_20261120 === '1') {
  const settlements3 = db.getSettlements();
  if (!settlements3.some((s) => s.id === 'settle_youth_allbank_20261120')) {
    const b = computeAccountBreakdown('youth');
    const now3 = new Date().toISOString();
    settlements3.push({
      id: 'settle_youth_allbank_20261120',
      date: now3.slice(0, 10),
      account: 'youth',
      cash: 0,
      tng: 0,
      bank: b.current.total,
      total: b.current.total,
      source: 'manual',
      memo: `청년부는 현금 없이 전부 은행 — 실제 잔액 구성의 현금(${b.current.cash})/터치앤고(${b.current.tng})를 계좌로 합산해 새 기준점으로 기록 (총액 ${b.current.total}은 그대로 유지)`,
      createdAt: now3
    });
    db.saveSettlements(settlements3);
    console.log('[youth-budget] applied RESET_YOUTH_BASELINE_TO_BANK_20261120:', JSON.stringify(b.current), '-> cash:0, tng:0, bank:' + b.current.total);
  }
}

// "대시보드 잔액과 실제 잔액 구성이 같아야지" — "청년부 잔액(장부 기준)"
// RM2,378과 "실제 잔액 구성" RM19,820이 안 맞는 근본 원인은, 계좌분리
// 이전(2026-11-14)에 찍힌 정산 2건(settle_seed0001, settle_cardfix20261114)
// 이 account 필드가 없어 기본값으로 'youth'에 귀속되면서, 사실은
// 카페+청년부가 합쳐져 있던 당시의 큰 금액을 전부 청년부 몫인 것처럼
// "실제 잔액 구성"의 기준점으로 물려받았기 때문 — 위의
// RESET_YOUTH_BASELINE_TO_BANK 마이그레이션도 그 잘못된 숫자를 그대로
// 계좌로만 옮겼을 뿐 근본 원인은 못 고쳤다. 지금은 청년부 거래가 전부
// 건별로 제대로 들어가 있으므로(수련회+회비장부 이관 완료), 장부 기준
// 합계(= "청년부 잔액")를 유일한 진실로 삼아 실제 잔액 구성의 기준점을
// 오늘 날짜로 다시 찍는다 — cash:0, tng:0, bank:장부잔액. 카페 쪽
// 정산은 전혀 건드리지 않는다. 멱등 — 고정 id가 이미 있으면 다시 안 만듦.
if (process.env.RESET_YOUTH_BASELINE_TO_LEDGER_20261120 === '1') {
  const settlements4 = db.getSettlements();
  if (!settlements4.some((s) => s.id === 'settle_youth_ledgerfix_20261120')) {
    const txns4 = db.getTransactions();
    const ledgerBalance = txns4
      .filter((t) => txnAccount(t) === 'youth')
      .reduce((sum, t) => sum + (t.type === 'income' ? t.amount : -t.amount), 0);
    const now4 = new Date().toISOString();
    settlements4.push({
      id: 'settle_youth_ledgerfix_20261120',
      date: now4.slice(0, 10),
      account: 'youth',
      cash: 0,
      tng: 0,
      bank: ledgerBalance,
      total: ledgerBalance,
      source: 'manual',
      memo: `"청년부 잔액(장부 기준)"과 "실제 잔액 구성"이 어긋나 있던 근본 원인(2026-11-14 계좌분리 이전 정산 2건이 account 미지정으로 청년부에 잘못 귀속)을 바로잡음 — 장부 합계(${ledgerBalance})를 그대로 새 기준점으로 기록, 전부 계좌.`,
      createdAt: now4
    });
    db.saveSettlements(settlements4);
    console.log('[youth-budget] applied RESET_YOUTH_BASELINE_TO_LEDGER_20261120: ledgerBalance=' + ledgerBalance + ' -> cash:0, tng:0, bank:' + ledgerBalance);
  }
}

// "정크 파일 지워줘" — 이제 더 이상 기준점으로 쓰이지 않는, 과거 오류의
// 흔적인 청년부 정산 4건을 정리한다: (1) settle_retreat1_20261120 —
// 수련회 시드 당시 잘못된 기준점을 그대로 물려받아 찍혔던 RM19,553짜리
// 정산, (2) settle_youth_allbank_20261120 — RESET_YOUTH_BASELINE_TO_BANK
// 단계에서 찍었던 RM19,820짜리 정산(장부기준 수정으로 완전히 대체됨),
// (3)(4) 2026-11-14 계좌분리 이전, account 미지정으로 청년부에 잘못
// 귀속됐던 원조 정산 2건(카페+청년부가 합쳐져 있던 당시 금액). 지금의
// 유일한 기준점인 settle_youth_ledgerfix_20261120(오늘자, 장부기준)은
// 그대로 둔다. 카페 쪽 정산은 전혀 건드리지 않음. 멱등 — 이미 지워졌으면
// 아무것도 안 함.
if (process.env.CLEANUP_JUNK_YOUTH_SETTLEMENTS_20261120 === '1') {
  const settlements5 = db.getSettlements();
  const toDelete = settlements5.filter((s) =>
    s.id === 'settle_retreat1_20261120' ||
    s.id === 'settle_youth_allbank_20261120' ||
    (settlementAccount(s) === 'youth' && s.date === '2026-11-14')
  );
  if (toDelete.length) {
    const keepIds = new Set(settlements5.map((s) => s.id).filter((id) => !toDelete.some((d) => d.id === id)));
    const remaining = settlements5.filter((s) => keepIds.has(s.id));
    db.saveSettlements(remaining);
    console.log('[youth-budget] applied CLEANUP_JUNK_YOUTH_SETTLEMENTS_20261120: 삭제된 정크 정산',
      toDelete.length + '건:', toDelete.map((s) => `${s.id}(${s.date}/RM${s.total})`).join(', '));
  }
}

// 읽기 전용 진단: "2026-11-18 청년부 자동반영" 정산이 사실 카페 숫자처럼
// 보인다는 제보 확인용 — 현재 남아있는 정산 전체를 날짜순으로 덤프.
// 아무것도 바꾸지 않음. DIAG_SETTLEMENTS_20261120=1일 때만.
if (process.env.DIAG_SETTLEMENTS_20261120 === '1') {
  const allSettle = db.getSettlements().slice().sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));
  console.log('[diag] ===== 정산 전체 ' + allSettle.length + '건 (날짜순) =====');
  allSettle.forEach((s) => {
    console.log(`[diag]   ${s.date} | account=${settlementAccount(s)} | source=${s.source} | cash=${s.cash} tng=${s.tng} bank=${s.bank} total=${s.total} | id=${s.id} | createdAt=${s.createdAt}`);
  });
}

// "꽃 구매가 청년부였는데... 이거 카페지..." — 2026-11-18에 꽃 구매(RM98,
// 청년부 지출)가 저장되면서 자동으로 찍힌 "청년부 자동반영" 정산
// (settle_mwlldaxdccg0l1, cash=1513.85/tng=1414.8/bank=14992.6)은, 그
// 시점에 아직 source=cafe 거래들이 account='youth'로 잘못 덮어써져 있던
// (오늘 고친) 버그 때문에, 실제로는 카페 전체 숫자가 거의 그대로 "청년부"
// 라벨을 달고 찍힌 것 — 카페의 11/16 정산(cash=1513.85/tng=1414.8/
// bank=15090.6)과 거의 동일하고 차이(bank 98)가 정확히 그 꽃 구매 금액과
// 일치하는 게 그 증거. 지금은 더 이상 기준점으로도 안 쓰이는 죽은
// 데이터라 혼란만 주므로 삭제. 멱등 — 고정 id 매칭.
if (process.env.DELETE_STALE_YOUTH_AUTO_SETTLEMENT_20261120 === '1') {
  const settlements6 = db.getSettlements();
  const before = settlements6.length;
  const remaining6 = settlements6.filter((s) => s.id !== 'settle_mwlldaxdccg0l1');
  if (remaining6.length !== before) {
    db.saveSettlements(remaining6);
    console.log('[youth-budget] applied DELETE_STALE_YOUTH_AUTO_SETTLEMENT_20261120: settle_mwlldaxdccg0l1 삭제');
  }
}

// 변경이력(히스토리) 시드: 2026-11-20 세션에 코드 배포로 일괄 반영된
// 이관/수정 작업들을 관리자 "변경이력" 화면에서 볼 수 있도록 기록한다.
// 멱등 — id가 이미 있으면 다시 안 넣음.
if (process.env.SEED_CHANGELOG_20261120 === '1') {
  const log = db.getChangelog();
  const existing = new Set(log.map((e) => e.id));
  const CHANGELOG_ENTRIES = [
    ['cl_20261120_1', '1차 수련회(4/30~5/2) 재정 이관', '엑셀 "세부내역" 탭 기준으로 회비 19건·후원 11건+교회지원금·구매내역(88건 중 리더단회의비 7건/RM348 제외) 1건으로 합산 반영. 총 32건.'],
    ['cl_20261120_2', '집사님 카페 지출 53건 재태깅', '과거 집사님이 정리해 넘긴 카페 지출입금 53건(청년부 일반지출로 잡혀있던 것)을 카페 운영으로 재분류. 계좌는 원래도 카페라 안 바뀌고 태그만 바뀜.'],
    ['cl_20261120_3', 'source=cafe 거래 계좌 강제 규칙 적용', '"소스가 카페면 카페 제정" 원칙에 따라 cafe-pos 동기화 거래는 예외 없이 계좌=카페·태그=카페운영으로 고정하는 상시 규칙 추가. 최초 적용 시 계좌 오분류 457건 수정.'],
    ['cl_20261120_4', '회비 납부 내역(4~11월) 개인장부 이관', '임창현님 개인 "회비 납부 내역" 엑셀을 청년부 재정에 건별 반영 — 회비 72건(RM4,130), 후원금 3건(RM2,400, 성격 수정분 포함: 여사무엘 150→회비50+후원100 분리, 카페 아웃팅지원 2000, 임혜란 집사님 300), 일반운영 지출 23건(RM6,263.48, 이미 입력된 2건 제외).'],
    ['cl_20261120_5', '청년부 계좌 현금/TnG → 은행 통일', '"청년부 재정은 현금 없이 다 은행" 확인에 따라 청년부 계좌 거래의 channel이 현금/터치앤고로 남아있던 건(수련회 회비 중 Cash 4건)을 전부 은행으로 재지정하는 상시 규칙 추가.'],
    ['cl_20261120_6', '청년부 실제잔액 기준점 재설정(현금/TnG→계좌)', '"실제 잔액 구성"에 남아있던 현금(1,529)/터치앤고(1,000)를 계좌로 합산한 새 기준점 정산을 오늘 날짜로 기록. 총액(RM19,820)은 변경 없이 유지, 구성만 전부 계좌로 재분류. 카페 쪽 정산은 그대로 둠.'],
    ['cl_20261120_7', '전체 오류 검증 통과', '거래 id/externalId 중복, 금액·타입·날짜·카테고리·태그 이상, 계좌별 채널 강제 적용, 정산 중복 등 12개 항목 전체 점검 — 문제 없음 확인.'],
    ['cl_20261120_8', '청년부 실제잔액 ↔ 장부잔액 불일치 근본 수정', '"청년부 잔액(장부 기준)"과 "실제 잔액 구성"이 어긋나던 근본 원인(2026-11-14 계좌분리 이전 정산 2건이 account 미지정으로 청년부에 잘못 귀속)을 찾아, 장부 합계를 그대로 새 기준점(전부 계좌)으로 재설정 — 이제 두 숫자가 항상 일치함. 카페 쪽 정산은 그대로 둠.'],
    ['cl_20261120_9', '정크 정산 4건 삭제 + 정산 폼에 계좌 선택 추가', '더 이상 기준점으로 쓰이지 않는 과거 오류 정산 4건(RM19,553/RM19,820/2026-11-14 2건)을 정리. "실제 잔액 기준점 다시 맞추기" 폼에 계좌(청년부/카페) 선택 드롭다운 추가 — 지금까지는 청년부 고정이었음.'],
    ['cl_20261120_10', '정크 정산 1건 추가 삭제 (11/18 청년부 자동반영)', '꽃 구매(RM98) 저장 시 자동 찍혔던 2026-11-18 "청년부" 정산이, 당시 아직 안 고쳐져 있던 계좌 버그 때문에 사실상 카페 전체 숫자를 그대로 담고 있던 것으로 확인돼 삭제. 이제 정산은 청년부 1건(오늘자, 장부기준)·카페 1건(11/16, 카페 자동)만 남음.']
  ];
  let added = 0;
  CHANGELOG_ENTRIES.forEach(([id, title, detail]) => {
    if (existing.has(id)) return;
    log.push({ id, date: '2026-11-20', title, detail, createdAt: new Date().toISOString() });
    added += 1;
  });
  if (added) {
    db.saveChangelog(log);
    console.log('[youth-budget] applied SEED_CHANGELOG_20261120:', added, 'entries added');
  }
}

// 읽기 전용 전체 오류 검증: 이번 세션에 적용한 여러 마이그레이션(카페 재태깅,
// source=cafe 계좌 강제, 회비 장부 시드, 은행 채널 통일) 이후 데이터 전체가
// 일관적인지 한 번에 점검한다. 아무것도 바꾸지 않음.
// FULL_VALIDATION_20261120=1일 때만 실행.
if (process.env.FULL_VALIDATION_20261120 === '1') {
  const txns = db.getTransactions();
  const settlements = db.getSettlements();
  const problems = [];

  // 1) id 중복
  const idCounts = {};
  txns.forEach((t) => { idCounts[t.id] = (idCounts[t.id] || 0) + 1; });
  const dupIds = Object.entries(idCounts).filter(([, c]) => c > 1);
  if (dupIds.length) problems.push(`거래 id 중복 ${dupIds.length}건: ${dupIds.slice(0, 5).map(([id]) => id).join(', ')}`);
  else console.log('[check] 거래 id 중복: 없음 (전체 ' + txns.length + '건)');

  // 2) externalId 중복 (null/빈값 제외)
  const extCounts = {};
  txns.forEach((t) => { if (t.externalId) extCounts[t.externalId] = (extCounts[t.externalId] || 0) + 1; });
  const dupExt = Object.entries(extCounts).filter(([, c]) => c > 1);
  if (dupExt.length) problems.push(`externalId 중복 ${dupExt.length}건: ${dupExt.slice(0, 5).map(([id]) => id).join(', ')}`);
  else console.log('[check] externalId 중복: 없음');

  // 3) 금액이 0 이하이거나 숫자가 아닌 거래
  const badAmount = txns.filter((t) => typeof t.amount !== 'number' || !(t.amount > 0) || Number.isNaN(t.amount));
  if (badAmount.length) problems.push(`amount가 0 이하/비정상인 거래 ${badAmount.length}건: ${badAmount.slice(0, 5).map((t) => t.id).join(', ')}`);
  else console.log('[check] 금액(amount) 0 이하/비정상: 없음');

  // 4) type이 income/expense가 아닌 거래
  const badType = txns.filter((t) => t.type !== 'income' && t.type !== 'expense');
  if (badType.length) problems.push(`type이 income/expense가 아닌 거래 ${badType.length}건: ${badType.slice(0, 5).map((t) => t.id + '(' + t.type + ')').join(', ')}`);
  else console.log('[check] type(income/expense) 이상: 없음');

  // 5) 날짜 형식(YYYY-MM-DD) 이상 또는 오늘(2026-11-20) 이후 미래 날짜
  const TODAY = '2026-11-20';
  const badDate = txns.filter((t) => !/^\d{4}-\d{2}-\d{2}/.test(t.date || ''));
  const futureDate = txns.filter((t) => /^\d{4}-\d{2}-\d{2}/.test(t.date || '') && t.date > TODAY);
  if (badDate.length) problems.push(`날짜 형식 이상 ${badDate.length}건: ${badDate.slice(0, 5).map((t) => t.id).join(', ')}`);
  else console.log('[check] 날짜 형식 이상: 없음');
  if (futureDate.length) problems.push(`오늘(${TODAY}) 이후 미래 날짜 거래 ${futureDate.length}건: ${futureDate.slice(0, 5).map((t) => `${t.id}(${t.date})`).join(', ')}`);
  else console.log('[check] 미래 날짜 거래: 없음');

  // 6) 카테고리가 정해진 목록(PURCHASE_CATEGORIES/INCOME_CATEGORIES) 밖인 거래
  // — source=cafe(cafe-pos 동기화)는 메뉴명을 그대로 카테고리로 쓰고(Ice
  // Coffee, Cookie 등), source=carryover(집사님 이월 잔액 1건)도 별도
  // "이월금" 성격이라 둘 다 이 체크 대상에서 제외한다. source=transfer(카페
  // ↔청년부 후원/이체)와 source=channel-transfer(계좌 간 현금/은행 전환)도
  // 각각 고정된 카테고리('청년부 후원'/'카페 후원금', '계좌 전환')를 쓰는
  // 전용 기능이라 이 목록 체크 대상이 아니다.
  const badCategory = txns.filter((t) => {
    if (t.source === 'cafe' || t.source === 'carryover' || t.source === 'transfer' || t.source === 'channel-transfer') return false;
    const list = t.type === 'income' ? INCOME_CATEGORIES : PURCHASE_CATEGORIES;
    return t.category && !list.includes(t.category);
  });
  if (badCategory.length) problems.push(`정의되지 않은 카테고리 ${badCategory.length}건: ${badCategory.slice(0, 5).map((t) => `${t.id}(${t.category})`).join(', ')}`);
  else console.log('[check] 카테고리(목록 밖): 없음');

  // 7) tag가 정의된 값 밖인 거래
  const VALID_TAGS = ['', 'cafe-ops', 'youth-general', 'needs-review', null, undefined];
  const badTag = txns.filter((t) => !VALID_TAGS.includes(t.tag));
  if (badTag.length) problems.push(`정의되지 않은 tag ${badTag.length}건: ${badTag.slice(0, 5).map((t) => `${t.id}(${t.tag})`).join(', ')}`);
  else console.log('[check] tag(목록 밖): 없음');

  // 8) needs-review(확인 필요) 태그가 남아있는 거래
  const needsReview = txns.filter((t) => t.tag === 'needs-review');
  if (needsReview.length) console.log(`[check] ⚠️ tag="확인필요"로 남아있는 거래 ${needsReview.length}건 (오류는 아니지만 수동 확인 필요)`);
  else console.log('[check] needs-review 거래: 없음');

  // 9) channel이 없어 실제 잔액 구성(현금/TnG/은행)에서 빠지는 거래
  const noChannel = txns.filter((t) => t.source !== 'channel-transfer' && !t.channel);
  console.log(`[check] channel 없는 거래(잔액 구성에서 제외됨): ${noChannel.length}건` + (noChannel.length ? ' — ' + noChannel.slice(0, 5).map((t) => `${t.id}(${t.date}/${t.category})`).join(', ') : ''));

  // 10) 청년부 계좌인데 channel이 은행(bank)이 아닌 거래 (이번 마이그레이션이 제대로 먹혔는지)
  const youthNotBank = txns.filter((t) => txnAccount(t) === 'youth' && t.channel && t.channel !== 'bank');
  if (youthNotBank.length) problems.push(`청년부 계좌인데 bank가 아닌 거래 ${youthNotBank.length}건: ${youthNotBank.slice(0, 5).map((t) => `${t.id}(${t.channel})`).join(', ')}`);
  else console.log('[check] 청년부 계좌 중 bank 아닌 거래: 없음 (현금/TnG 통일 마이그레이션 정상 적용)');

  // 11) 카페 소스인데 계좌/태그가 여전히 안 맞는 거래 (강제 마이그레이션 검증)
  const cafeSrcBad = txns.filter((t) => t.source === 'cafe' && (t.account !== 'cafe' || t.tag !== 'cafe-ops'));
  if (cafeSrcBad.length) problems.push(`source=cafe인데 account/tag 불일치 ${cafeSrcBad.length}건: ${cafeSrcBad.slice(0, 5).map((t) => `${t.id}(${t.account}/${t.tag})`).join(', ')}`);
  else console.log('[check] source=cafe 계좌/태그 강제 적용: 정상');

  // 12) 정산(settlement) id 중복 및 account 미해석
  const settleIdCounts = {};
  settlements.forEach((s) => { settleIdCounts[s.id] = (settleIdCounts[s.id] || 0) + 1; });
  const dupSettle = Object.entries(settleIdCounts).filter(([, c]) => c > 1);
  if (dupSettle.length) problems.push(`정산 id 중복 ${dupSettle.length}건: ${dupSettle.map(([id]) => id).join(', ')}`);
  else console.log('[check] 정산(settlement) id 중복: 없음 (전체 ' + settlements.length + '건)');

  // 13) 전체 합계 및 계좌별 breakdown 눈으로 확인용 출력
  const youthB = computeAccountBreakdown('youth');
  const cafeB = computeAccountBreakdown('cafe');
  console.log('[check] 청년부 실제 잔액 구성:', JSON.stringify(youthB.current));
  console.log('[check] 카페 실제 잔액 구성:', JSON.stringify(cafeB.current));
  const grandBalance = txns.reduce((sum, t) => sum + (t.type === 'income' ? t.amount : -t.amount), 0);
  console.log('[check] 전체 거래 장부상 순합계(모든 소스, 이체 포함 net 0):', grandBalance);

  console.log('[check] ===== 요약 =====');
  if (problems.length === 0) {
    console.log('[check] ✅ 문제 없음 — 위 12개 항목 전부 통과');
  } else {
    console.log(`[check] ❌ 문제 ${problems.length}건 발견:`);
    problems.forEach((p) => console.log('[check]   - ' + p));
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
function computeAccountBreakdown(account) {
  // Only a real physical-count baseline (manual entry or the cafe-pos
  // settlement webhook) is used as the anchor — never an "auto" row, which
  // is just a snapshot of this same calculation at an earlier moment. If an
  // auto row were used as the anchor, deleting or editing a transaction
  // that happened *before* that snapshot would have no way to flow through
  // (it would already be "baked into" the frozen baseline forever). Always
  // anchoring on the last real count and re-summing every transaction that
  // still exists since then means a delete/edit anywhere is reflected
  // immediately, with no stale numbers to clean up. Both the settlement
  // history and the transaction ledger are filtered to this one account
  // (카페 vs 청년부) first, so each account's baseline/delta math never
  // crosses into the other's money.
  const settlements = db.getSettlements()
    .filter((s) => s.source !== 'auto' && settlementAccount(s) === account)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  const baseline = settlements[0] || null;
  const baselineCreatedAt = baseline ? baseline.createdAt : null;
  const txns = db.getTransactions().filter((t) => txnAccount(t) === account);

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
    account,
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

// account: 'youth' | 'cafe' | undefined/'all'. 'all'은 두 계좌를 각자 계산한
// 뒤 숫자만 더한 것 — 기준점(baseline)은 계좌마다 다른 시점일 수 있어 하나로
// 합쳐 보여줄 의미가 없으므로 null로 둔다(각 계좌의 baseline은 byAccount에
// 들어있음).
function computeBalanceBreakdown(account) {
  if (account === 'cafe' || account === 'youth') return computeAccountBreakdown(account);
  const youth = computeAccountBreakdown('youth');
  const cafe = computeAccountBreakdown('cafe');
  return {
    account: 'all',
    baseline: null,
    deltas: {
      cash: youth.deltas.cash + cafe.deltas.cash,
      tng: youth.deltas.tng + cafe.deltas.tng,
      bank: youth.deltas.bank + cafe.deltas.bank
    },
    current: {
      cash: youth.current.cash + cafe.current.cash,
      tng: youth.current.tng + cafe.current.tng,
      bank: youth.current.bank + cafe.current.bank,
      total: youth.current.total + cafe.current.total
    },
    unattributedCount: youth.unattributedCount + cafe.unattributedCount,
    unattributedAmount: youth.unattributedAmount + cafe.unattributedAmount,
    byAccount: { youth, cafe }
  };
}

// Called right after a channel-tagged income/expense is saved, so the 정산
// (settlement) history gets a new row reflecting the up-to-date cash/TnG/
// bank split automatically — instead of requiring a fresh manual count each
// time. A transaction with no channel can't be placed in a bucket, so it's
// left out (same as in computeBalanceBreakdown) and no row is added for it.
function recordAutoSettlement(txn) {
  if (!txn || !txn.channel) return;
  const account = txnAccount(txn);
  const b = computeAccountBreakdown(account);
  const entry = {
    id: 'settle_' + uid(),
    date: txn.date,
    account,
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
  res.json(computeBalanceBreakdown(req.query.account));
});

// Shared by the manual "카페에서 새 내역 가져오기" button (POST /api/admin/sync-cafe)
// and the settlement webhook below — pulling new orders/expenses from
// cafe-pos and recording them as transactions. Factored out so pressing
// "정산 전송" on the cafe side can trigger both the balance update AND this
// transaction sync in one go, instead of requiring a separate manual click.
// skipAutoSettlement: the settlement webhook calls this right before it
// saves the REAL "cafe-pos" settlement row a moment later — an "auto"
// breadcrumb row here would just be an almost-identical row sitting next
// to the real one in 정산 history a few milliseconds apart, which reads as
// a duplicate rather than useful information. The standalone manual
// "새 내역 가져오기" button (the other caller) has no such real settlement
// following it, so it keeps the breadcrumb (default false = keep it).
//
// fullRescan: the normal flow only asks cafe-pos for orders/expenses newer
// than the saved cursor (lastOrderAt/lastExpenseCreatedAt), so once that
// cursor has moved past something, a transaction that gets deleted here
// (by mistake, or during cleanup) can never come back on a normal press —
// it's simply older than the cursor. fullRescan re-asks cafe-pos for its
// ENTIRE order/expense history instead of just what's new. This is safe to
// press any number of times, including back-to-back, because every row is
// still deduped by externalId below: anything still present is skipped,
// and only rows that are genuinely missing (i.e. were deleted, or never
// made it in) get re-added. The saved cursor itself is never moved
// backwards by this — see below — so it doesn't change how normal syncs
// behave afterwards.
// performCafeSync itself is queued below (search "cafeSyncQueue") so that
// pressing the button several times fast, or a webhook press landing while
// a manual sync is still mid-flight, can never run two syncs at once — each
// one reads the cursor/transaction list only after the previous one has
// fully written its results, so a repeat press that found nothing new
// correctly reports added: 0 instead of re-adding the same rows.
async function doPerformCafeSync({ skipAutoSettlement = false, fullRescan = false } = {}) {
  const baseUrl = process.env.CAFE_BASE_URL || 'https://jusumcafe.up.railway.app';
  const user = process.env.CAFE_ADMIN_USER;
  const pass = process.env.CAFE_ADMIN_PASS;
  if (!user || !pass) {
    throw new Error('CAFE_ADMIN_USER / CAFE_ADMIN_PASS 환경변수가 설정되어 있지 않습니다 (cafe-pos의 ADMIN_USER/ADMIN_PASS와 동일한 값).');
  }
  const settings = db.getSettings();
  const fetchSince = fullRescan ? { lastOrderAt: null, lastExpenseCreatedAt: null } : settings.cafeSync;
  const { newTxns, lastOrderAt, lastExpenseCreatedAt } = await syncCafe({ baseUrl, user, pass, since: fetchSince });

  const txns = db.getTransactions();
  const existingIds = new Set(txns.map((t) => t.externalId).filter(Boolean));
  const toAdd = newTxns.filter((t) => !existingIds.has(t.externalId));
  toAdd.forEach((t) => txns.push(t));
  db.saveTransactions(txns);

  // Only ever move the cursor forward. On a normal sync this is a no-op
  // (lastOrderAt/lastExpenseCreatedAt already come out >= the saved
  // cursor). On a fullRescan it stops the cursor from being reset
  // backwards to the start of history, which would otherwise make the
  // *next* normal press re-scan everything again.
  settings.cafeSync = {
    lastOrderAt: (lastOrderAt || '') > (settings.cafeSync.lastOrderAt || '') ? lastOrderAt : settings.cafeSync.lastOrderAt,
    lastExpenseCreatedAt: (lastExpenseCreatedAt || '') > (settings.cafeSync.lastExpenseCreatedAt || '') ? lastExpenseCreatedAt : settings.cafeSync.lastExpenseCreatedAt
  };
  db.saveSettings(settings);

  // One settlement row for the whole sync, not one per order, so the
  // 정산 history doesn't get flooded by a busy day at the cafe.
  if (!skipAutoSettlement && toAdd.some((t) => t.channel)) {
    const b = computeAccountBreakdown('cafe'); // 동기화되는 거래는 전부 카페 쪽 계좌
    const list = db.getSettlements();
    list.push({
      id: 'settle_' + uid(),
      date: todayStr(),
      account: 'cafe',
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

  return { added: toAdd.length, needsReview: toAdd.filter((t) => t.tag === 'needs-review').length };
}

// Queue wrapper: every operation that reads-then-writes the transaction or
// settlement ledgers (the manual sync button, a recovery fullRescan, and
// the settlement webhook's sync-then-settle sequence below) is chained onto
// this one promise, so they always run strictly one after another even if
// two requests arrive milliseconds apart — never two at once reading the
// same "before" state and both deciding the same row (or the same
// increase) is new.
let cafeOpsQueue = Promise.resolve();
function queueCafeOp(fn) {
  const run = cafeOpsQueue.then(fn);
  cafeOpsQueue = run.then(
    () => {},
    () => {}
  );
  return run;
}

function performCafeSync(opts) {
  return queueCafeOp(() => doPerformCafeSync(opts));
}

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
  const { from, to, source, type, tag, category, account } = req.query;
  if (from) txns = txns.filter((t) => t.date >= from);
  if (to) txns = txns.filter((t) => t.date <= to);
  if (source) txns = txns.filter((t) => t.source === source);
  if (type) txns = txns.filter((t) => t.type === type);
  if (tag) txns = txns.filter((t) => t.tag === tag);
  if (category) txns = txns.filter((t) => t.category === category);
  if (account === 'youth' || account === 'cafe') txns = txns.filter((t) => txnAccount(t) === account);
  txns.sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt));
  res.json(txns);
});

app.get('/api/summary', (req, res) => {
  let txns = db.getTransactions();
  const { account } = req.query;
  if (account === 'youth' || account === 'cafe') txns = txns.filter((t) => txnAccount(t) === account);
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

    // 현금↔은행 같은 "채널 전환"은 실제 수입/지출이 아니라 같은 돈이 형태만
    // 바꾼 것 — 태그/카테고리/월별 집계에 끼면 수입·지출이 실제보다 부풀려
    // 보이므로 제외한다. (balance와 bySource는 그대로 — balance는 +금액/-금액
    // 짝이라 net 0이고, bySource엔 "계좌 전환"으로 따로 보여서 투명함.)
    if (t.source === 'channel-transfer') return;

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
app.post('/api/webhook/cafe-settlement', async (req, res) => {
  const secret = process.env.SETTLEMENT_WEBHOOK_SECRET;
  if (!secret) return res.status(503).json({ error: 'SETTLEMENT_WEBHOOK_SECRET not configured on this server' });
  if (req.headers['x-webhook-secret'] !== secret) {
    return res.status(401).json({ error: 'invalid webhook secret' });
  }
  // 카페(jusumcafe)의 현금/터치앤고/은행 숫자는 카페 자신의 서랍/지갑/계좌
  // 기준일 뿐, 파이낸스 전체(회비·수동입력 포함)의 전부가 아니다. 그래서
  // 셋 다 "카페 쪽에서 이번에 얼마나 늘었는지(증가분)"만 받아서 파이낸스의
  // 현재 잔액에 더한다 — 카페가 보낸 절대금액으로 그대로 덮어쓰면, 그 사이
  // 회비 입금이나 수동 입력으로 쌓인 파이낸스 자체의 현금/은행 잔액이 통째로
  // 사라진다(실제로 이 버그 때문에 은행 잔액이 한 번에 0 근처로 사라졌었음).
  const { date, cashIncrease, tngIncrease, bankIncrease, memo, externalId } = req.body || {};
  if (typeof cashIncrease !== 'number' || typeof tngIncrease !== 'number' || typeof bankIncrease !== 'number') {
    return res.status(400).json({ error: 'cashIncrease, tngIncrease, and bankIncrease must all be numbers' });
  }

  // The whole read-sync-write sequence below runs as one queued operation
  // (same queue performCafeSync uses) so that two overlapping webhook
  // deliveries — a retry, or a second "정산 전송" press landing before the
  // first finishes — can never both read the same "before" balance and
  // both apply their increase on top of it. doPerformCafeSync (not the
  // queued performCafeSync wrapper) is called directly here since we're
  // already running inside the queue.
  try {
    const entry = await queueCafeOp(async () => {
      // externalId identifies which cafe-pos settlement record this push
      // is for. cafe-pos resends the same settlement (same externalId)
      // when an admin edits a mistyped entry there — in that case this
      // isn't a NEW increase on top of the old one, it's a CORRECTION to
      // it. We only ever apply the difference between the new increase
      // and whatever increase we already applied for that same
      // externalId, so editing a cafe-pos settlement twice (or resending
      // after a webhook hiccup) can't double (or triple) count the same
      // sales.
      const list = db.getSettlements();
      const existing = externalId ? list.find((s) => s.externalId === externalId) : null;
      const appliedCashIncrease = cashIncrease - (existing ? existing.cashIncrease || 0 : 0);
      const appliedTngIncrease = tngIncrease - (existing ? existing.tngIncrease || 0 : 0);
      const appliedBankIncrease = bankIncrease - (existing ? existing.bankIncrease || 0 : 0);

      // Snapshot the current balance BEFORE the sync below touches the
      // ledger. The increases above (from cafe-pos) already account for
      // every order since cafe-pos's last settlement, including the ones
      // about to be synced in as transactions — so this has to be the
      // pre-sync figure, or those orders would get counted twice (once via
      // the increase, once via their own channel delta once synced in).
      // This is the CAFE account's own balance, not the combined total —
      // cafe-pos's drawer/wallet counts have nothing to do with 청년부's own
      // cash/membership money, now that the two are fully separate.
      const before = computeAccountBreakdown('cafe').current;
      const cash = before.cash + appliedCashIncrease;
      const tng = before.tng + appliedTngIncrease;
      const bank = before.bank + appliedBankIncrease;

      // "정산 전송" 버튼 한 번으로 잔액뿐 아니라 그동안 쌓인 주문/지출
      // 내역도 같이 가져오도록, 새 정산 기준점을 기록하기 전에 먼저 카페
      // 동기화를 돌린다. 동기화된 거래들의 createdAt이 아래에서 저장할
      // 정산 기준점의 createdAt보다 앞서게 되므로, 그 기준점 이후의
      // 증감만 다시 더하는 computeBalanceBreakdown 로직에서 중복으로
      // 잡히지 않는다 — 방금 센 실제 증가분에 이미 반영돼 있는 거래이기
      // 때문. 동기화가 실패해도(환경변수 누락, 네트워크 오류 등) 정산
      // 자체는 실패시키지 않는다 — 정산이 더 중요하므로.
      let cafeSync = null;
      try {
        cafeSync = await doPerformCafeSync({ skipAutoSettlement: true });
      } catch (e) {
        console.error('[youth-budget] cafe-settlement 웹훅 이전 카페 동기화 실패:', e.message);
      }

      const memoText = memo || `카페 정산 자동 반영 (현금 ${fmtSigned(cashIncrease)}, TnG ${fmtSigned(tngIncrease)}, 은행 ${fmtSigned(bankIncrease)})`;
      let entry;
      if (existing) {
        existing.date = date || existing.date;
        existing.account = 'cafe';
        existing.cash = cash;
        existing.tng = tng;
        existing.bank = bank;
        existing.total = cash + tng + bank;
        existing.cashIncrease = cashIncrease;
        existing.tngIncrease = tngIncrease;
        existing.bankIncrease = bankIncrease;
        existing.memo = memoText;
        existing.updatedAt = new Date().toISOString();
        entry = existing;
      } else {
        entry = {
          id: 'settle_' + uid(),
          date: date || todayStr(),
          account: 'cafe',
          cash,
          tng,
          bank,
          total: cash + tng + bank,
          source: 'cafe-pos',
          externalId: externalId || null,
          cashIncrease,
          tngIncrease,
          bankIncrease,
          memo: memoText,
          createdAt: new Date().toISOString()
        };
        list.push(entry);
      }
      db.saveSettlements(list);

      return cafeSync ? Object.assign({}, entry, { cafeSync }) : entry;
    });

    res.status(201).json(entry);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ---------------- Admin-only API ----------------
app.use('/api/admin', maybeRequireAdmin);

// 변경이력(변경이력/이관 작업 감사 로그) — 최신순.
app.get('/api/admin/changelog', (req, res) => {
  const log = db.getChangelog().slice().sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt));
  res.json(log);
});

app.post('/api/admin/transactions', (req, res) => {
  const { date, type, source, tag, category, item, amount, memo, author, channel, account } = req.body || {};
  if (!date || (type !== 'income' && type !== 'expense') || !category || typeof amount !== 'number' || amount <= 0) {
    return res.status(400).json({ error: 'date, type (income|expense), category and a positive amount are required' });
  }
  const entry = {
    id: 'tx_' + uid(),
    date,
    type,
    source: source || 'manual',
    account: account === 'cafe' || account === 'youth' ? account : ((source || 'manual') === 'cafe' ? 'cafe' : 'youth'),
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
  const { date, category, vendor, item, amount, memo, author, payment, account } = req.body || {};
  const amountNum = parseFloat(amount);
  if (!date || !PURCHASE_CATEGORIES.includes(category) || !item || Number.isNaN(amountNum) || amountNum <= 0 || !PAYMENT_METHODS.includes(payment)) {
    return res.status(400).json({ error: `date, category (${PURCHASE_CATEGORIES.join('|')}), item, a positive amount, and payment (${PAYMENT_METHODS.join('|')}) are required` });
  }
  const entry = {
    id: 'tx_' + uid(),
    date,
    type: 'expense',
    source: 'manual',
    // 밥값 등 청년부 지출이 카페 돈으로 결제되는 경우가 있어서, 어느
    // 계좌에서 실제로 돈이 나갔는지 입력하는 사람이 고른다 — 기본은
    // 청년부. tag(청년부 일반지출)는 account와 별개로 "용도"를 나타내는
    // 값이라 account가 cafe여도 그대로 둔다.
    account: account === 'cafe' ? 'cafe' : 'youth',
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
// membership-fee history; everything else is "manual". Factored out like
// createPurchaseEntry so both the admin form and the public no-login
// quick-entry page (아래) share one implementation.
function createIncomeEntry(req, res) {
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
    account: 'youth',
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
}

app.post('/api/admin/income', createIncomeEntry);

// Public quick-entry — no admin login needed to add 회비/후원금 등 입금,
// same as the purchase/expense quick-entry page above.
app.post('/api/entries/income', createIncomeEntry);

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
      account: 'youth',
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
  const fields = ['date', 'type', 'source', 'account', 'tag', 'category', 'item', 'vendor', 'payer', 'payment', 'channel', 'amount', 'memo', 'author'];
  fields.forEach((f) => { if (req.body && req.body[f] !== undefined) txns[idx][f] = req.body[f]; });
  db.saveTransactions(txns);
  res.json(txns[idx]);
});

app.delete('/api/admin/transactions/:id', (req, res) => {
  const all = db.getTransactions();
  const target = all.find((t) => t.id === req.params.id);
  // A 이체(transfer)의 두 짝(카페쪽 지출 / 청년부쪽 수입)은 transferId로
  // 묶여 있다 — 한쪽만 지우면 돈이 한쪽 계좌에서만 사라진 것처럼 보이는
  // 반쪽짜리 기록이 남으므로, 삭제할 땐 짝도 같이 지운다.
  const idsToRemove = new Set([req.params.id]);
  if (target && target.transferId) {
    all.forEach((t) => { if (t.transferId === target.transferId) idsToRemove.add(t.id); });
  }
  const txns = all.filter((t) => !idsToRemove.has(t.id));
  db.saveTransactions(txns);
  // Remove the 정산 row that was auto-created for this transaction (and its
  // transfer sibling, if any) too, so the settlement history doesn't keep a
  // stale entry for something that no longer exists. (The live balance
  // total is unaffected either way — see computeBalanceBreakdown — this is
  // just keeping the history honest.)
  const settlements = db.getSettlements();
  const kept = settlements.filter((s) => !idsToRemove.has(s.txnId));
  if (kept.length !== settlements.length) db.saveSettlements(kept);
  res.json({ ok: true });
});

// Pull new orders/expenses from cafe-pos.
// Manual settlement entry (e.g. seeding today's actual balances by hand).
// account defaults to 'youth' since this form is for counting 청년부's own
// cash/TnG/bank by hand — cafe's own count always comes in automatically via
// the /api/webhook/cafe-settlement route instead.
app.post('/api/admin/settlements', (req, res) => {
  const { date, cash, tng, bank, memo, account } = req.body || {};
  if (typeof cash !== 'number' || typeof tng !== 'number' || typeof bank !== 'number') {
    return res.status(400).json({ error: 'cash, tng, bank must all be numbers' });
  }
  const entry = {
    id: 'settle_' + uid(),
    date: date || todayStr(),
    account: account === 'cafe' ? 'cafe' : 'youth',
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
  const { date, cash, tng, bank, memo, account } = req.body || {};
  if (date !== undefined) list[idx].date = date;
  if (cash !== undefined) list[idx].cash = parseFloat(cash);
  if (tng !== undefined) list[idx].tng = parseFloat(tng);
  if (bank !== undefined) list[idx].bank = bank === '' || bank === null ? null : parseFloat(bank);
  if (memo !== undefined) list[idx].memo = memo;
  if (account === 'cafe' || account === 'youth') list[idx].account = account;
  list[idx].total = list[idx].cash + list[idx].tng + (list[idx].bank || 0);
  db.saveSettlements(list);
  res.json(list[idx]);
});

// 현금 ↔ 터치앤고 ↔ 은행 전환: 같은 계좌 안에서 돈의 "형태"만 바뀌는 경우
// (예: 현금이 너무 많이 쌓여서 은행에 입금). /api/admin/transfer(카페→청년부)
// 와 달리 계좌(account)는 그대로고 채널(channel)만 한쪽에서 다른 쪽으로
// 옮겨진다 — 지출 한 건 + 수입 한 건을 짝지어 기록해서 전체 합계는 그대로
// 유지하면서 현금/은행 버킷 구성만 정확히 반영되게 한다. 보고서(카테고리별
// ·월별 집계)에는 실제 수입/지출이 아니므로 집계에서 제외된다 — /api/summary
// 참고.
app.post('/api/admin/channel-transfer', (req, res) => {
  const { date, amount, from, to, memo, author, account } = req.body || {};
  const amountNum = parseFloat(amount);
  if (!date || Number.isNaN(amountNum) || amountNum <= 0 || !RECEIVE_CHANNELS.includes(from) || !RECEIVE_CHANNELS.includes(to) || from === to) {
    return res.status(400).json({ error: `date, 0보다 큰 amount, 서로 다른 from/to (${RECEIVE_CHANNELS.join('|')})가 필요합니다` });
  }
  const acc = account === 'cafe' ? 'cafe' : 'youth';
  const transferId = 'chxfer_' + uid();
  const now = new Date().toISOString();
  const label = `${from} → ${to} 전환`;
  const outTxn = {
    id: 'tx_' + uid(),
    date,
    type: 'expense',
    source: 'channel-transfer',
    account: acc,
    tag: null,
    category: '계좌 전환',
    item: label,
    amount: amountNum,
    channel: channelToBucket(from),
    memo: memo || '',
    author: author || '',
    transferId,
    externalId: null,
    createdAt: now
  };
  const inTxn = {
    id: 'tx_' + uid(),
    date,
    type: 'income',
    source: 'channel-transfer',
    account: acc,
    tag: null,
    category: '계좌 전환',
    item: label,
    amount: amountNum,
    channel: channelToBucket(to),
    memo: memo || '',
    author: author || '',
    transferId,
    externalId: null,
    createdAt: now
  };
  const txns = db.getTransactions();
  txns.push(outTxn, inTxn);
  db.saveTransactions(txns);
  recordAutoSettlement(outTxn);
  recordAutoSettlement(inTxn);
  res.status(201).json({ out: outTxn, in: inTxn });
});

// 카페 → 청년부 후원/이체: 실제로 카페 쪽 현금/계좌가 청년부 쪽으로 넘어갈
// 때 관리자가 기록한다. 카페 계좌엔 지출, 청년부 계좌엔 수입으로 각각 한
// 건씩 남아서 두 계좌 잔액에 모두 정확히 반영되고, "전체" 합계는 서로
// 상쇄되어 바뀌지 않는다 — 실제로 바깥에서 돈이 들어오거나 나간 게 아니라
// 두 계좌 사이에서만 옮겨졌기 때문. (카페 돈으로 청년모임 식사비 등을
// "직접" 결제한 경우는 이 기능이 아니라 기존의 "청년부 일반지출" 태그로
// 충분히 표현된다 — 그건 카페 계좌에서 바로 나간 지출일 뿐, 현금이 청년부
// 손으로 넘어온 적이 없기 때문.)
app.post('/api/admin/transfer', (req, res) => {
  const { date, amount, memo, channel, author } = req.body || {};
  const amountNum = parseFloat(amount);
  if (!date || Number.isNaN(amountNum) || amountNum <= 0 || !RECEIVE_CHANNELS.includes(channel)) {
    return res.status(400).json({ error: `date, 0보다 큰 amount, channel (${RECEIVE_CHANNELS.join('|')})이 필요합니다` });
  }
  const transferId = 'xfer_' + uid();
  const ch = channelToBucket(channel);
  const now = new Date().toISOString();
  const outTxn = {
    id: 'tx_' + uid(),
    date,
    type: 'expense',
    source: 'transfer',
    account: 'cafe',
    tag: null,
    category: '청년부 후원',
    item: '청년부 후원/이체',
    amount: amountNum,
    channel: ch,
    memo: memo || '',
    author: author || '',
    transferId,
    externalId: null,
    createdAt: now
  };
  const inTxn = {
    id: 'tx_' + uid(),
    date,
    type: 'income',
    source: 'transfer',
    account: 'youth',
    tag: 'youth-general',
    category: '카페 후원금',
    item: '카페 후원금',
    amount: amountNum,
    channel: ch,
    memo: memo || '',
    author: author || '',
    transferId,
    externalId: null,
    createdAt: now
  };
  const txns = db.getTransactions();
  txns.push(outTxn, inTxn);
  db.saveTransactions(txns);
  recordAutoSettlement(outTxn);
  recordAutoSettlement(inTxn);
  res.status(201).json({ out: outTxn, in: inTxn });
});

app.delete('/api/admin/settlements/:id', (req, res) => {
  const list = db.getSettlements().filter((s) => s.id !== req.params.id);
  db.saveSettlements(list);
  res.json({ ok: true });
});

app.post('/api/admin/sync-cafe', async (req, res) => {
  try {
    const fullRescan = !!(req.body && req.body.full);
    const result = await performCafeSync({ fullRescan });
    res.json(result);
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
  // 채널 전환(현금↔은행 등)은 실제 수입이 아니므로 "수입" 시트에선 뺀다 —
  // 그래도 "전체내역" 시트에는 그대로 남아 있어 추적 가능하다.
  const income = txns.filter((t) => t.type === 'income' && t.source !== 'channel-transfer');
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
