// api/sheets-read.js
// Reads a tab's data from Google Sheets and returns it as {headers, data} where data is
// an array of {ColumnHeader: value} objects. Requires GOOGLE_SERVICE_ACCOUNT_JSON and
// GOOGLE_SHEETS_ID env vars to be set in Vercel (Settings -> Environment Variables).

const { google } = require("googleapis");

async function authenticateSheets() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const sheetsId = process.env.GOOGLE_SHEETS_ID;
    if (!sheetsId) return res.status(500).json({ error: "GOOGLE_SHEETS_ID env var not set" });

    const { sheetName } = req.query;
    if (!sheetName) return res.status(400).json({ error: "sheetName query param required" });

    const sheets = await authenticateSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetsId,
      range: `${sheetName}!A:Z`,
      // Without this, the API returns FORMATTED display strings by default (e.g.
      // "10,489" with the thousands-separator comma baked in as text) instead of the
      // raw number 10489. The dashboard's parseFloat() stops at the first comma, so
      // every value with a thousands separator was silently truncating — a barge
      // showing "10,489" in the sheet was being read back as just "10". This forces
      // raw numeric values instead.
      valueRenderOption: "UNFORMATTED_VALUE",
    });

    const rows = response.data.values || [];
    if (rows.length === 0) {
      return res.status(200).json({ data: [], headers: [] });
    }

    const headers = rows[0];
    const data = rows.slice(1)
      .filter((row) => row.some((cell) => cell !== undefined && cell !== ""))
      .map((row) => {
        const obj = {};
        headers.forEach((header, idx) => { obj[header] = row[idx] !== undefined ? row[idx] : ""; });
        return obj;
      });

    res.status(200).json({ headers, data });
  } catch (error) {
    console.error("Sheets read error:", error);
    res.status(500).json({ error: "Failed to read from Sheets", details: error.message });
  }
}
