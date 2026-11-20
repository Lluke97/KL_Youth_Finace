// Tiny JSON-file "database" — same approach as cafe-pos. Good enough for a
// youth group's transaction volume on a single Railway volume.
const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

function filePath(name) {
  return path.join(DATA_DIR, name + '.json');
}

// Baked into the deployed image (seed-data/), not into the persistent
// volume. On a brand-new/empty volume (first deploy, or a wiped volume),
// transactions.json doesn't exist yet — in that one case we seed it from
// the historical import instead of starting from an empty ledger. Once the
// real file exists on the volume, it's the source of truth and this is
// never consulted again.
const SEED_DIR = path.join(__dirname, 'seed-data');
function loadSeed(name, fallback) {
  const fp = path.join(SEED_DIR, name + '.seed.json');
  if (!fs.existsSync(fp)) return fallback;
  try {
    return JSON.parse(fs.readFileSync(fp, 'utf8'));
  } catch (e) {
    console.error('Failed to parse seed', fp, e);
    return fallback;
  }
}

function readJSON(name, fallback) {
  const fp = filePath(name);
  if (!fs.existsSync(fp)) {
    fs.writeFileSync(fp, JSON.stringify(fallback, null, 2));
    return JSON.parse(JSON.stringify(fallback));
  }
  try {
    return JSON.parse(fs.readFileSync(fp, 'utf8'));
  } catch (e) {
    console.error('Failed to parse', fp, e);
    return JSON.parse(JSON.stringify(fallback));
  }
}

function writeJSON(name, data) {
  const fp = filePath(name);
  const tmp = fp + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2));
  fs.renameSync(tmp, fp);
}

// A transaction:
// {
//   id, date: 'YYYY-MM-DD', type: 'income'|'expense',
//   source: 'cafe'|'membership'|'manual'|'carryover',
//   tag: 'cafe-ops'|'youth-general'|null   // only meaningful for source='cafe' expenses —
//                                            // lets us split "cafe operating cost" from
//                                            // "youth-group spending that happened to be
//                                            // logged through the cafe ledger"
//   category, item, amount, memo, author,
//   externalId: string|null,  // e.g. cafe-pos order/expense id, used to dedupe sync
//   createdAt: ISO timestamp
// }
const DEFAULT_SETTINGS = {
  cafeSync: { lastOrderAt: null, lastExpenseCreatedAt: null },
  carryoverSet: false
};

// A settlement: the actual cash / Touch 'n Go / bank balance at a point in
// time (실사 금액) — distinct from the transaction ledger above, which is
// built from recorded income/expense rows. The two should track each other,
// but a settlement is the "ground truth" check: what's actually on hand.
// { id, date, cash, tng, bank, total, source: 'cafe-pos'|'manual', memo, createdAt }
module.exports = {
  getTransactions: () => readJSON('transactions', loadSeed('transactions', [])),
  saveTransactions: (txns) => writeJSON('transactions', txns),

  getSettlements: () => readJSON('settlements', loadSeed('settlements', [])),
  saveSettlements: (settlements) => writeJSON('settlements', settlements),

  getSettings: () => Object.assign({}, DEFAULT_SETTINGS, readJSON('settings', DEFAULT_SETTINGS)),
  saveSettings: (settings) => writeJSON('settings', settings),

  // 변경이력(changelog): 코드 배포로 일괄 반영된 마이그레이션/데이터 이관
  // 작업의 기록. 일반 거래(transactions)와 달리 "무엇을, 왜, 몇 건" 바꿨는지
  // 설명하는 관리자 전용 감사 로그. { id, date, title, detail, createdAt }
  getChangelog: () => readJSON('changelog', []),
  saveChangelog: (entries) => writeJSON('changelog', entries)
};
