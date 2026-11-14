// CLI wrapper around lib/cafeSync — lets you run a sync by hand, e.g. via a
// Railway cron job, without going through the HTTP API.
// Env vars: CAFE_BASE_URL, CAFE_ADMIN_USER, CAFE_ADMIN_PASS
const db = require('../db');
const { syncCafe } = require('../lib/cafeSync');

async function main() {
  const baseUrl = process.env.CAFE_BASE_URL || 'https://jusumcafe.up.railway.app';
  const user = process.env.CAFE_ADMIN_USER;
  const pass = process.env.CAFE_ADMIN_PASS;
  if (!user || !pass) {
    console.error('Set CAFE_ADMIN_USER and CAFE_ADMIN_PASS (cafe-pos\'s ADMIN_USER/ADMIN_PASS) before syncing.');
    process.exit(1);
  }

  const settings = db.getSettings();
  const { newTxns, lastOrderAt, lastExpenseCreatedAt } = await syncCafe({
    baseUrl,
    user,
    pass,
    since: settings.cafeSync
  });

  const txns = db.getTransactions();
  const existingIds = new Set(txns.map((t) => t.externalId).filter(Boolean));
  const toAdd = newTxns.filter((t) => !existingIds.has(t.externalId));
  toAdd.forEach((t) => txns.push(t));
  db.saveTransactions(txns);

  settings.cafeSync = { lastOrderAt, lastExpenseCreatedAt };
  db.saveSettings(settings);

  console.log(`Synced ${toAdd.length} new transaction(s) from ${baseUrl}.`);
}

main().catch((e) => { console.error('Sync failed:', e.message); process.exit(1); });
