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
// bucket(cash/tng/bank) → 표시용 라벨. paymentToChannel/channelToBucket의
// 역방향. channel이 어떤 경로로든 바뀌었을 때 receiveChannel/payment 같은
// 표시용 필드도 같이 맞춰주는 데 쓴다.
const BUCKET_TO_RECEIVE_CHANNEL = { cash: '현금', tng: '터치앤고', bank: '계좌이체' };
const BUCKET_TO_PAYMENT = { cash: '현금', tng: '터치앤고', bank: '카드' };

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

// "청년부 재정은 현금이 없고 다 은행이다" 불변식을 저장 시점마다 바로
// 적용한다. 예전엔 이 규칙을 서버 시작 시 한 번만 훑어서 되돌리는
// forceYouthChannelToBank() IIFE로만 지켰는데, 그 사이(다음 배포/재시작
// 전까지)는 생성·수정 직후 잠깐이 아니라 계속 틀린 채로 남아있을 수
// 있었다(수입 입력 시 수금채널을 현금/터치앤고로 고르거나, 카페→청년부
// 이체에서 현금으로 전달했다고 기록하는 경우 등 실제로 거쳐가는 경로가
// 여러 곳이라서). 이제 거래가 만들어지거나 수정되는 모든 지점에서 이
// 함수를 거쳐 즉시 맞춘다 — 표시용 라벨(receiveChannel/payment)도 같이
// 맞춘다. 카페 계좌는 실제로 서랍에 현금이 있으므로 건드리지 않는다.
// 이미 bank인 건 다시 안 건드리므로 몇 번을 호출해도 안전하다(멱등).
function normalizeChannelForAccount(t) {
  if (t && txnAccount(t) === 'youth' && (t.channel === 'cash' || t.channel === 'tng')) {
    t.channel = 'bank';
    if (t.receiveChannel) t.receiveChannel = '계좌이체';
    if (t.payment) t.payment = '카드';
  }
  return t;
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
  // 신뢰할 수 없는 맨 왼쪽(공격자가 보낸 헤더 그대로)이 아니라, Railway
  // 엣지가 실제 연결에서 붙인 맨 오른쪽 항목을 써야 한다 — 왼쪽을 쓰면
  // 공격자가 X-Forwarded-For를 요청마다 바꿔서 로그인 잠금을 그냥 우회할
  // 수 있었음(cafe-pos 쪽과 동일한 버그, 동일하게 수정).
  const header = req.headers['x-forwarded-for'];
  if (header) {
    const parts = header.split(',').map((p) => p.trim()).filter(Boolean);
    if (parts.length) return parts[parts.length - 1];
  }
  return req.socket.remoteAddress || 'unknown';
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
// 거래는 예외 없이 카페 계좌(account='cafe')여야 한다. 실제 카페→청년부
// 이체는 별도 source('transfer')로 기록되므로(/api/admin/transfer 참고),
// source='cafe'인데 account가 'youth'로 덮어써진 건은 전부 과거의 수동
// 오분류다. account는 이미 올바른 건 다시 건드리지 않으므로(멱등) 별도
// env 게이트 없이 매 배포/재시작마다 돌아도 안전.
//
// tag는 전과 달리 "태그가 아예 없을 때만" cafe-ops로 채워준다 — 예전엔
// tag가 cafe-ops가 아니면 매번 강제로 되돌렸는데, 그러면 관리자가 "확인
// 필요"(needs-review)로 들어온 건을 손으로 "카페 운영"/"청년부 일반지출"
// 중 하나로 재분류해도 다음 배포/재시작 때 전부 cafe-ops로 도로 덮어써져
// 버렸다(그 재분류 기능이 사실상 무용지물이었던 버그). cafeSync.js가 동기화
// 시점에 이미 tag를 채워 넣으므로(원재료/운영비→cafe-ops, 기타→needs-review)
// 여기서는 과거(그 로직이 없던 시절) 데이터처럼 tag가 비어있는 경우에만
// 보정한다.
(function forceCafeSourceToCafeAccount() {
  const txns = db.getTransactions();
  let accountChanged = 0;
  let tagChanged = 0;
  txns.forEach((t) => {
    if (t.source === 'cafe') {
      if (t.account !== 'cafe') { t.account = 'cafe'; accountChanged += 1; }
      if (!t.tag) { t.tag = 'cafe-ops'; tagChanged += 1; }
    }
  });
  if (accountChanged || tagChanged) {
    db.saveTransactions(txns);
    console.log(`[youth-budget] source=cafe 거래 중 계좌 재지정 ${accountChanged}건, 태그 없는 건 카페운영으로 채움 ${tagChanged}건 완료.`);
  }
})();

// "청년부 재정은 현금이 없고 다 은행이다" / "현금을 은행으로 다 옮길 걸로
// 쳐줘" / "다 현금 말고 은행으로 잡아줘" — 청년부 계좌(account='youth')
// 거래 중 channel이 'cash'(현금)·'tng'(터치앤고)로 잡혀 있는 건을 전부
// bank(은행)로 바꾼다. 표시용 라벨(receiveChannel/payment)도 같이
// 맞춘다. 카페 계좌는 실제로 카페 서랍에 현금이 있으므로 건드리지 않는다.
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

// 읽기 전용 전체 오류 검증 — 거래/정산 데이터 전체가 일관적인지(중복,
// 이상값, 계좌/채널 규칙 위반 등) 언제든 다시 점검할 수 있는 상시 진단.
// 아무것도 바꾸지 않음. FULL_VALIDATION_20261120=1일 때만 실행.
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

  // 5) 날짜 형식(YYYY-MM-DD) 이상 또는 오늘 이후 미래 날짜 — 하드코딩된
  // 날짜 문자열이었던 걸 todayStr()로 바꿈(예전 값 그대로면 하루만 지나도
  // 모든 "오늘" 거래가 "미래 날짜"로 잘못 걸림).
  const TODAY = todayStr();
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
  // tag는 더 이상 반드시 'cafe-ops'일 필요 없음 — 관리자가 손으로
  // needs-review/youth-general로 재분류했을 수 있으므로, 여기서는 account만
  // 엄격히 체크하고 tag는 "비어있지 않은지"만 확인한다.
  const cafeSrcBad = txns.filter((t) => t.source === 'cafe' && (t.account !== 'cafe' || !t.tag));
  if (cafeSrcBad.length) problems.push(`source=cafe인데 account 불일치 또는 tag 없음 ${cafeSrcBad.length}건: ${cafeSrcBad.slice(0, 5).map((t) => `${t.id}(${t.account}/${t.tag})`).join(', ')}`);
  else console.log('[check] source=cafe 계좌 강제 적용 및 tag 존재: 정상');

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
function computeAccountBreakdown(account, { excludeCafeSynced = false } = {}) {
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
    // excludeCafeSynced: used only by the cafe-settlement webhook, to avoid
    // double-counting cafe-pos sales/expenses that were ALSO already
    // individually imported as transactions (source='cafe') by the manual
    // "카페에서 새 내역 가져오기" button — see the comment at the webhook
    // handler below for the full double-count scenario this prevents.
    if (excludeCafeSynced && t.source === 'cafe') return;
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

  let txns = db.getTransactions();
  let purgedIds = new Set();
  if (fullRescan) {
    // 수입 동기화 행은 "그 날짜/채널/품목으로 묶인 주문 id 목록"을 해시해서
    // externalId를 만드는데(lib/cafeSync.js), fullRescan은 커서 없이 카페의
    // 주문 전체를 다시 긁어오므로 그 묶음(어떤 주문들이 한 그룹에 들어가는지)
    // 자체가 지난번(부분 동기화 때)과 달라질 수 있다 — 그 사이 같은
    // 날짜/채널/품목으로 주문이 더 들어왔다면 그룹 구성이 달라져서 새
    // externalId가 기존 것과 안 맞고, 기존 행은 안 지워진 채 겹치는
    // 금액이 또 하나 추가되는 이중계산이 날 수 있었다("빠진 내역 전체
    // 복구"가 매번 눌러도 안전하다던 전제가 사실은 깨져 있던 지점). 이를
    // 막기 위해 fullRescan에서는 기존 cafe-pos 수입 동기화 행을 전부 지우고
    // 지금 막 다시 계산한 완전한 묶음으로 통째로 교체한다 — "지금 카페
    // 주문 전체 기준으로 다시 그리기"이므로 항상 멱등하다. 지출(expense)
    // 동기화 행은 주문 묶음이 아니라 cafe-pos 지출 1건=1행으로 고정
    // externalId라 이 문제가 없어 그대로 둔다.
    txns.forEach((t) => {
      if (t.source === 'cafe' && t.type === 'income' && (t.externalId || '').startsWith('cafepos_income_')) {
        purgedIds.add(t.id);
      }
    });
    if (purgedIds.size) txns = txns.filter((t) => !purgedIds.has(t.id));
  }

  const existingIds = new Set(txns.map((t) => t.externalId).filter(Boolean));
  const toAdd = newTxns.filter((t) => !existingIds.has(t.externalId));
  toAdd.forEach((t) => txns.push(t));
  db.saveTransactions(txns);

  if (purgedIds.size) {
    // 지운 행들이 남겨놓은 auto 정산 스냅샷도 같이 정리 — 숫자엔 영향 없지만
    // (auto는 기준점으로 안 쓰임) 고아 이력이 쌓이지 않도록.
    const settlementsAfterPurge = db.getSettlements();
    const keptAfterPurge = settlementsAfterPurge.filter((s) => !purgedIds.has(s.txnId));
    if (keptAfterPurge.length !== settlementsAfterPurge.length) db.saveSettlements(keptAfterPurge);
  }

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
// 관리자 "정산" 탭은 "기준점을 손으로 다시 맞추는" 화면이라, 손으로 세거나
// 카페 쪽 자동 동기화로 들어온 실제 기준점(source !== 'auto')만 보여준다.
// 후원/이체·계좌 전환을 기록할 때마다 recordAutoSettlement()가 남기는
// source='auto' 스냅샷은 computeAccountBreakdown()의 기준점으로도 쓰이지
// 않는 단순 이력이라, 여기 섞여 보이면 "또 정크가 쌓였나" 하고 혼동만 준다
// — 그래서 기본적으로 숨긴다(지우진 않음, 데이터는 남아있음).
app.get('/api/settlements', (req, res) => {
  const settlements = db.getSettlements()
    .filter((s) => s.source !== 'auto')
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
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
      // ledger, and with excludeCafeSynced:true — NOT just because of the
      // sync about to run inside this same call, but because the manual
      // "카페에서 새 내역 가져오기" button might ALREADY have individually
      // imported some of these same orders/expenses (as source='cafe'
      // transactions with their own channel) at any earlier point between
      // settlements. cashIncrease/tngIncrease/bankIncrease always represent
      // the FULL change over the whole settled window regardless of
      // whether some of it was already synced in as transactions — so if
      // `before` included those transactions' own channel deltas too, this
      // settlement's increase would double-count them (once as live
      // transaction deltas, once via the increase). Excluding source='cafe'
      // here means cafe-pos sales/expenses are only ever counted once, via
      // this authoritative increase — never twice via both paths. This is
      // the CAFE account's own balance, not the combined total — cafe-pos's
      // drawer/wallet counts have nothing to do with 청년부's own
      // cash/membership money, now that the two are fully separate.
      const before = computeAccountBreakdown('cafe', { excludeCafeSynced: true }).current;
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
  normalizeChannelForAccount(entry);
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
  normalizeChannelForAccount(entry);
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
  normalizeChannelForAccount(entry);
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
  // 구매(지출) 수정 폼에는 channel 선택란이 따로 없고 payment(결제수단)만
  // 고치게 되어 있다 — payment를 현금→카드로 바꿔도 channel이 그대로
  // 'cash'로 남으면 실제 잔액 구성(현금/TnG/계좌, computeAccountBreakdown)이
  // 예전 결제수단 기준으로 계속 집계돼서 정산 탭 숫자가 안 맞게 된다.
  // payment만 바뀌고 channel은 호출자가 따로 안 보낸 경우에 한해 channel을
  // payment 기준으로 다시 계산해준다 (channel을 명시적으로 같이 보낸
  // 경우엔 그 값을 그대로 존중한다).
  if (req.body && req.body.payment !== undefined && req.body.channel === undefined) {
    const derived = paymentToChannel(req.body.payment);
    if (derived) txns[idx].channel = derived;
  }
  // account를 cafe→youth로 바꾸는 경우처럼, 수정 결과로 "청년부 계좌인데
  // channel이 현금/TnG" 상태가 될 수도 있다 — 생성 때와 동일하게 여기서도
  // 바로 맞춰준다.
  normalizeChannelForAccount(txns[idx]);
  // channel이 이번 요청에서 바뀌었다면(수입 수정 화면의 "수금 채널" 직접
  // 수정이든, 위의 payment→channel 유도든) 표시용 라벨(receiveChannel/
  // payment)도 최종 channel에 맞춰 같이 정리한다. 안 그러면 잔액 집계는
  // 맞게 바뀌는데 목록 화면의 "수금 채널"/"결제수단" 칸은 예전 값 그대로
  // 남아서 또 다른 형태의 안 맞는 느낌을 준다. 그 필드가 원래 없던
  // 레코드엔 새로 만들지 않는다(undefined는 그대로 둔다).
  if (req.body && (req.body.channel !== undefined || req.body.payment !== undefined)) {
    const finalCh = txns[idx].channel;
    if (finalCh) {
      if (txns[idx].receiveChannel !== undefined) txns[idx].receiveChannel = BUCKET_TO_RECEIVE_CHANNEL[finalCh] || txns[idx].receiveChannel;
      if (txns[idx].payment !== undefined) txns[idx].payment = BUCKET_TO_PAYMENT[finalCh] || txns[idx].payment;
    }
  }
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
  // 청년부 계좌는 현금/터치앤고를 실제로 보유하지 않는다 — "현금 ↔ 계좌
  // 전환" 폼의 계좌 선택란 기본값이 "청년부"라서, 실수로 청년부를 고른
  // 채 현금 관련 전환을 기록하면(예: 현금→은행) 두 다리 모두 결국
  // normalizeChannelForAccount에 의해 bank로 맞춰져 서로 상쇄되는
  // 의미없는 거래 쌍만 남는다. 조용히 무해한 쓰레기 거래를 남기는 대신
  // 여기서 바로 막아서 왜 안 되는지 알려준다.
  if (acc === 'youth' && (from !== '계좌이체' || to !== '계좌이체')) {
    return res.status(400).json({ error: '청년부 계좌는 현금/터치앤고를 보유하지 않아요(전부 은행 처리). 현금↔계좌 전환은 카페 계좌에서만 기록하세요.' });
  }
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
  // outTxn은 카페 계좌라 실제 서랍 현금/계좌 그대로 두지만, inTxn은 항상
  // 청년부 계좌라서 전달 수단이 현금/터치앤고였어도(카페에서 현금으로
  // 직접 건네준 경우) 청년부 쪽에는 "은행"으로 맞춰준다 — 청년부는 현금을
  // 들고 있지 않는다는 불변식을 여기서도 지켜야 한다.
  normalizeChannelForAccount(inTxn);
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
