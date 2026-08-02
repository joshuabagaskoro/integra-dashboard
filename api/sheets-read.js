// api/sheets-read.js
// Serverless function to read data from Google Sheets

const { google } = require('googleapis');

const SHEETS_ID = process.env.GOOGLE_SHEETS_ID; // You'll set this in Vercel env vars
const CREDENTIALS = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

async function authenticateSheets() {
  const auth = new google.auth.GoogleAuth({
    credentials: CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const sheets = await authenticateSheets();
    const { sheetName } = req.query; // Which tab to read: "Domes", "Barges", etc.

    if (!sheetName) {
      return res.status(400).json({ error: 'sheetName query param required' });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEETS_ID,
      range: `${sheetName}!A:Z`, // Read all columns up to Z
    });

    const rows = response.data.values || [];
    if (rows.length === 0) {
      return res.status(200).json({ data: [], headers: [] });
    }

    const headers = rows[0];
    const data = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx] || '';
      });
      return obj;
    });

    res.status(200).json({ headers, data });
  } catch (error) {
    console.error('Sheets read error:', error);
    res.status(500).json({ error: 'Failed to read from Sheets', details: error.message });
  }
}
