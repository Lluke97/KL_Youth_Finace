// Pulls new orders/expenses from the cafe-pos admin API (jusumcafe.up.railway.app)
// and turns them into youth-budget transactions.
//
// Orders -> cafe income, one row per day + payment channel + product name
// (matching how the historical cafe data already looked in Finance before
// cafe-pos existed — e.g. "Ice coffee +287", "Cookie +280" for a given
// day), rather than one row per individual order (a busy day is 50+
// orders of the same few drinks) or one single lump sum per day (loses
// which product the revenue came from). Cash/TnG stay separate rows even
// within the same product, because the balance breakdown attributes each
// transaction to a cash/TnG/bank bucket by its `channel` field — collapsing
// them would make that split unrecoverable. "Free"/comped orders are
// skipped since no money changed hands.
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

  // Group same-day, same-channel orders by product name instead of one row
  // per order (see file header). Item-level key = 'YYYY-MM-DD|cash|Ice coffee'.
  // Also track the day+channel's true order total separately (channelTotals)
  // as a reconciliation check below — pos.html computes an order's `total`
  // as the exact sum of item.price*qty, so in the normal case the two will
  // already match to the cent; this just guards against an edited order
  // whose total and items somehow drifted apart.
  const incomeGroups = new Map();
  const channelTotals = new Map();
  orders
    .filter((o) => (!since.lastOrderAt || o.createdAt > since.lastOrderAt) && (o.pay === '현금' || o.pay === '터치앤고'))
    .forEach((o) => {
      const date = myDateStr(o.createdAt);
      const channel = o.pay === '현금' ? 'cash' : 'tng';
      const ctKey = date + '|' + channel;
      if (!channelTotals.has(ctKey)) channelTotals.set(ctKey, { date, channel, pay: o.pay, orderTotal: 0, orderIds: [] });
      const ct = channelTotals.get(ctKey);
      ct.orderTotal += o.total;
      ct.orderIds.push(o.id);

      (o.items && o.items.length ? o.items : [{ name: '기타', qty: 1, price: o.total }]).forEach((it) => {
        const name = (it.name || '기타').trim() || '기타';
        const qty = typeof it.qty === 'number' ? it.qty : 1;
        const price = typeof it.price === 'number' ? it.price : 0;
        const key = date + '|' + channel + '|' + name;
        if (!incomeGroups.has(key)) incomeGroups.set(key, { date, channel, pay: o.pay, itemName: name, amount: 0, qty: 0, orderIds: [] });
        const g = incomeGroups.get(key);
        g.amount += qty * price;
        g.qty += qty;
        g.orderIds.push(o.id);
      });
    });

  // If a day+channel's items don't sum to exactly its orders' total (should
  // only happen on a hand-edited order), add one adjustment row for the
  // difference instead of silently losing it — the cash/TnG balance this
  // feeds into must always match the orders' actual total, never the
  // item breakdown.
  channelTotals.forEach((ct, ctKey) => {
    let itemSum = 0;
    incomeGroups.forEach((g) => {
      if (g.date === ct.date && g.channel === ct.channel) itemSum += g.amount;
    });
    const diff = Math.round((ct.orderTotal - itemSum) * 100) / 100;
    if (Math.abs(diff) >= 0.01) {
      incomeGroups.set(ctKey + '|__조정__', {
        date: ct.date, channel: ct.channel, pay: ct.pay, itemName: '조정', amount: diff, qty: 0, orderIds: ct.orderIds
      });
    }
  });

  incomeGroups.forEach((g) => {
    // Hash the grouped order ids (rather than reusing a single order's id)
    // so this group's externalId stays unique and stable even if the same
    // day/channel/product gets synced again later with a different set of
    // orders.
    const idHash = crypto.createHash('md5').update(g.itemName + '|' + g.orderIds.join(',')).digest('hex').slice(0, 12);
    newTxns.push({
      id: 'tx_cafepos_income_' + idHash,
      date: g.date,
      type: 'income',
      source: 'cafe',
      tag: 'cafe-ops',
      category: g.itemName,
      item: g.itemName,
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
