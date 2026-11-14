// One-off import of 집사님's historical cafe ledger (Open_Cafe-*.xlsx) into
// our transactions store. Safe to re-run: every imported row gets a stable
// externalId derived from its sheet+row content, and rows that already exist
// (by externalId) are skipped rather than duplicated.
//
// Usage: node scripts/import-excel.js /path/to/Open_Cafe-2026.xlsx
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const db = require('../db');

const file = process.argv[2];
if (!file) {
  console.error('Usage: node scripts/import-excel.js /path/to/file.xlsx');
  process.exit(1);
}

// Historical expense categories that are NOT cafe operating cost — they were
// logged in the same sheet for convenience but are general youth-group spend.
const YOUTH_GENERAL_EXPENSE_CATEGORIES = new Set(['Youth support']);
// Historical income rows that are not cafe sales (donations received through
// the cafe, not tied to any menu item).
const DONATION_INCOME_ITEMS = new Set(['Support', '후원금']);

function toDateStr(raw) {
  let d = raw;
  if (typeof raw === 'number') {
    const parsed = XLSX.SSF.parse_date_code(raw);
    d = new Date(parsed.y, parsed.m - 1, parsed.d);
  } else if (!(raw instanceof Date)) {
    d = new Date(raw);
  }
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function externalId(parts) {
  return 'xlsx-' + crypto.createHash('sha1').update(parts.join('|')).digest('hex').slice(0, 16);
}

function importIncome(wb, txns, seen) {
  const ws = wb.Sheets['Income'];
  if (!ws) { console.warn('No "Income" sheet found — skipping.'); return 0; }
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  let added = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const [dateRaw, menu, unit, qty, price] = row;
    if (!dateRaw || !menu) continue;
    const date = toDateStr(dateRaw);
    const amount = Number(price) || 0;
    if (amount === 0) continue;

    let source, category, memo, type = 'income';
    if (menu === '이월금') {
      source = 'carryover';
      category = '이월금';
      memo = '2025-11-30 기준 이월 잔액 (집사님 정리본)';
    } else if (DONATION_INCOME_ITEMS.has(menu)) {
      source = 'cafe';
      category = '후원금';
      memo = '카페를 통해 들어온 후원금';
    } else {
      source = 'cafe';
      category = menu;
      memo = `${menu} ${qty || ''}개 판매`.trim();
    }

    const id = externalId(['income', date, String(menu), String(unit), String(qty), String(price), String(i)]);
    if (seen.has(id)) continue;
    seen.add(id);
    txns.push({
      id: 'tx_' + id,
      date,
      type,
      source,
      tag: source === 'cafe' ? 'cafe-ops' : null,
      category,
      item: String(menu),
      amount,
      memo,
      author: '집사님(정리본)',
      externalId: id,
      createdAt: new Date().toISOString()
    });
    added++;
  }
  return added;
}

function importExpenses(wb, txns, seen) {
  const ws = wb.Sheets['Expenses'];
  if (!ws) { console.warn('No "Expenses" sheet found — skipping.'); return 0; }
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  let added = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const [dateRaw, cat, items, amt] = row;
    if (!dateRaw || !cat) continue;
    const date = toDateStr(dateRaw);
    const amount = Number(amt) || 0;
    if (amount === 0) continue;

    const tag = YOUTH_GENERAL_EXPENSE_CATEGORIES.has(cat) ? 'youth-general' : 'cafe-ops';
    const id = externalId(['expense', date, String(cat), String(items), String(amt), String(i)]);
    if (seen.has(id)) continue;
    seen.add(id);
    txns.push({
      id: 'tx_' + id,
      date,
      type: 'expense',
      source: 'cafe',
      tag,
      category: cat,
      item: String(items || cat),
      amount,
      memo: String(items || ''),
      author: '집사님(정리본)',
      externalId: id,
      createdAt: new Date().toISOString()
    });
    added++;
  }
  return added;
}

function main() {
  const wb = XLSX.readFile(path.resolve(file));
  const txns = db.getTransactions();
  const seen = new Set(txns.filter((t) => t.externalId).map((t) => t.externalId));

  const incomeAdded = importIncome(wb, txns, seen);
  const expenseAdded = importExpenses(wb, txns, seen);

  db.saveTransactions(txns);
  console.log(`Imported ${incomeAdded} income rows and ${expenseAdded} expense rows from ${file}.`);
  console.log(`Total transactions now: ${txns.length}`);
}

main();
