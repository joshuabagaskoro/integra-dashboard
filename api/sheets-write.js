// api/sheets-write.js
// Serverless function to write data to Google Sheets

const { google } = require('googleapis');

const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const CREDENTIALS = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

async function authenticateSheets() {
  const auth = new google.auth.GoogleAuth({
    credentials: CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { sheetName, data } = req.body; // data is array of arrays

    if (!sheetName || !data) {
      return res.status(400).json({ error: 'sheetName and data required' });
    }

    const sheets = await authenticateSheets();

    // Clear existing data (optional, or replace only)
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SHEETS_ID,
      range: `${sheetName}!A:Z`,
    });

    // Write new data
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEETS_ID,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: data },
    });

    res.status(200).json({ success: true, message: `${sheetName} updated` });
  } catch (error) {
    console.error('Sheets write error:', error);
    res.status(500).json({ error: 'Failed to write to Sheets', details: error.message });
  }
}
