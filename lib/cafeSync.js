// Pulls new orders/expenses from the cafe-pos admin API (jusumcafe.up.railway.app)
// and turns them into youth-budget transactions.
//
// Orders -> cafe income. Individual orders aren't pulled in one-by-one (a
// busy day is 50+ orders and nobody needs each coffee listed separately in
// Finance) — they're grouped into one income row per day per payment
// channel instead, e.g. "11/15 매출 (14건)" for that day's cash orders and
// a separate row for that day's TnG orders. Keeping cash/TnG as separate
// rows (rather than one lump per day) matters because the balance
// breakdown attributes each transaction to a cash/TnG/bank bucket by its
// `channel` field. "Free"/comped orders are skipped since no money changed
// hands.
// Expenses -> cafe expense, pulled in one row per cafe-pos expense entry
// (unlike income, each purchase is still listed individually — so the
// 구매지출 list shows exactly what was bought). cafe-pos's own categories
// are 원재료/운영비/기타; 원재료 and 운영비 are assumed to be real cafe
// operating cost. 기타 is ambiguous (it's sometimes used for youth-group
// spend logged through the cafe entry screen for convenience), so it's
// tagged 'needs-review' instead of being silently counted as cafe-ops — an
// admin re-tags it from the UI.
const https = require('https');
const http = require('http');
const crypto = require('crypto');

function fetchJSON(baseUrl, pathname, auth) {
  return new Promise((resolve, reject) => {
    const url = new URL(pathname, baseUrl);
    const lib = url.protocol === 'https:' ? https : http;
    const headers = {};
    if (auth) headers.Authorization = 'Basic ' + Buffer.from(auth.user + ':' + auth.pass).toString('base64');
    const req = lib.get(url, { headers }, (res) => {
      let body = '';
      res.on('data', (c) => (body += c));
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300) {
          return reject(new Error(`${pathname} -> HTTP ${res.statusCode}: ${body.slice(0, 200)}`));
        }
        try { resolve(JSON.parse(body)); } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
  });
}

function myDateStr(iso) {
  // cafe-pos runs on Asia/Kuala_Lumpur; match that when deriving a date.
  const fmt = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kuala_Lumpur', year: 'numeric', month: '2-digit', day: '2-digit' });
  return fmt.format(new Date(iso));
}

function expenseCategoryTag(category) {
  if (category === '기타') return 'needs-review';
  return 'cafe-ops'; // 원재료 | 운영비
}

async function syncCafe({ baseUrl, user, pass, since }) {
  const auth = { user, pass };
  const orders = await fetchJSON(baseUrl, '/api/admin/orders', auth);
  const expenses = await fetchJSON(baseUrl, '/api/admin/expenses', auth);

  const newTxns = [];

  // Group same-day, same-channel orders into one income row instead of one
  // row per order (see file header). Key = 'YYYY-MM-DD|cash'|'YYYY-MM-DD|tng'.
  const incomeGroups = new Map();
  orders
    .filter((o) => (!since.lastOrderAt || o.createdAt > since.lastOrderAt) && (o.pay === '현금' || o.pay === '터치앤고'))
    .forEach((o) => {
      const date = myDateStr(o.createdAt);
      const channel = o.pay === '현금' ? 'cash' : 'tng';
      const key = date + '|' + channel;
      if (!incomeGroups.has(key)) incomeGroups.set(key, { date, channel, pay: o.pay, amount: 0, count: 0, orderIds: [] });
      const g = incomeGroups.get(key);
      g.amount += o.total;
      g.count += 1;
      g.orderIds.push(o.id);
    });

  incomeGroups.forEach((g) => {
    const label = g.date.slice(5).replace('-', '/'); // '2026-11-15' -> '11/15'
    // Hash the grouped order ids (rather than reusing a single order's id)
    // so this group's externalId stays unique and stable even if the same
    // day/channel gets synced again later with a different set of orders.
    const idHash = crypto.createHash('md5').update(g.orderIds.join(',')).digest('hex').slice(0, 12);
    newTxns.push({
      id: 'tx_cafepos_income_' + idHash,
      date: g.date,
      type: 'income',
      source: 'cafe',
      tag: 'cafe-ops',
      category: '카페 매출',
      item: `${label} 매출 (${g.count}건)`,
      amount: g.amount,
      memo: `결제: ${g.pay}`,
      channel: g.channel,
      author: '카페 POS',
      externalId: 'cafepos_income_' + idHash,
      createdAt: new Date().toISOString()
    });
  });

  expenses
    .filter((e) => !since.lastExpenseCreatedAt || e.createdAt > since.lastExpenseCreatedAt)
    .forEach((e) => {
      newTxns.push({
        id: 'tx_cafepos_expense_' + e.id,
        date: e.date,
        type: 'expense',
        source: 'cafe',
        tag: expenseCategoryTag(e.category),
        category: e.category,
        item: e.memo || e.category,
        amount: e.amount,
        payment: e.payment || '',
        channel: e.payment === '현금' ? 'cash' : e.payment === '카드' ? 'bank' : e.payment === '터치앤고' ? 'tng' : null,
        memo: e.memo || '',
        author: e.author || '카페 POS',
        externalId: 'cafepos_expense_' + e.id,
        createdAt: new Date().toISOString()
      });
    });

  const maxOrderAt = orders.reduce((m, o) => (o.createdAt > m ? o.createdAt : m), since.lastOrderAt || '');
  const maxExpenseAt = expenses.reduce((m, e) => (e.createdAt > m ? e.createdAt : m), since.lastExpenseCreatedAt || '');

  return { newTxns, lastOrderAt: maxOrderAt || since.lastOrderAt, lastExpenseCreatedAt: maxExpenseAt || since.lastExpenseCreatedAt };
}

module.exports = { syncCafe };
