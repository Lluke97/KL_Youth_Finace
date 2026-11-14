// Exports all transactions to a Google Sheet. Needs:
//   GOOGLE_SHEET_ID              - the spreadsheet to write into
//   GOOGLE_SERVICE_ACCOUNT_JSON  - the *contents* of a Google Cloud service
//                                  account key (JSON), as a single env var.
// Setup (one-time, by Luke):
//   1. Google Cloud Console -> create a project -> enable "Google Sheets API"
//   2. IAM -> Service Accounts -> create one -> "Keys" tab -> Add key (JSON)
//   3. Open the target Google Sheet -> Share -> add the service account's
//      email (looks like xxx@xxx.iam.gserviceaccount.com) as an Editor
//   4. Put the whole JSON key file's content into GOOGLE_SERVICE_ACCOUNT_JSON
//      (Railway variable), and the sheet's ID (from its URL) into
//      GOOGLE_SHEET_ID.
const { google } = require('googleapis');

function isConfigured() {
  return !!(process.env.GOOGLE_SHEET_ID && process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
}

async function getSheetsClient() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

const HEADER = ['ID', 'Date', 'Type', 'Source', 'Tag', 'Category', 'Item', 'Amount', 'Memo', 'Author'];

async function exportTransactions(transactions) {
  if (!isConfigured()) throw new Error('Google Sheets not configured (set GOOGLE_SHEET_ID and GOOGLE_SERVICE_ACCOUNT_JSON)');
  const sheets = await getSheetsClient();
  const spreadsheetId = process.env.GOOGLE_SHEET_ID;
  const sheetName = process.env.GOOGLE_SHEET_TAB || 'Transactions';

  const rows = transactions
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((t) => [t.id, t.date, t.type, t.source, t.tag || '', t.category, t.item, t.amount, t.memo || '', t.author || '']);

  await sheets.spreadsheets.values.clear({ spreadsheetId, range: `${sheetName}!A:Z` });
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [HEADER, ...rows] }
  });
  return rows.length;
}

module.exports = { isConfigured, exportTransactions };
