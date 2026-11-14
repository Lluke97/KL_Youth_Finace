// Pulls new orders/expenses from the cafe-pos admin API (jusumcafe.up.railway.app)
// and turns them into youth-budget transactions.
//
// Orders -> cafe income (revenue actually collected; "free"/comped orders are
// skipped since no money changed hands).
// Expenses -> cafe expense. cafe-pos's own categories are 원재료/운영비/기타;
// 원재료 and 운영비 are assumed to be real cafe operating cost. 기타 is
// ambiguous (it's sometimes used for youth-group spend logged through the
// cafe entry screen for convenience), so it's tagged 'needs-review' instead
// of being silently counted as cafe-ops — an admin re-tags it from the UI.
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

  orders
    .filter((o) => (!since.lastOrderAt || o.createdAt > since.lastOrderAt) && (o.pay === '현금' || o.pay === '터치앤고'))
    .forEach((o) => {
      const itemsSummary = (o.items || []).map((it) => it.name || it.id).join(', ');
      newTxns.push({
        id: 'tx_cafepos_order_' + o.id,
        date: myDateStr(o.createdAt),
        type: 'income',
        source: 'cafe',
        tag: 'cafe-ops',
        category: '카페 매출',
        item: itemsSummary || '주문',
        amount: o.total,
        memo: `결제: ${o.pay}`,
        author: o.cashier || '카페 POS',
        externalId: 'cafepos_order_' + o.id,
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
