// api/cron-update-exchange-rate.js
// Triggered automatically by Vercel Cron (see vercel.json) — runs once a day, entirely
// server-side, independent of whether anyone has the dashboard open in a browser.
//
// This exists because the dashboard's own auto-fetch (see the useEffect in
// IntegraDashboard.jsx calling /api/fetch-exchange-rate) only runs while a logged-in
// admin has a browser tab open — if nobody with admin access opens the app on a given
// day, the rate simply doesn't update that day. This cron job is the actual guarantee
// of "always daily," independent of anyone using the app at all.
//
// Deploy: this file alone isn't enough — vercel.json must also declare the schedule
// (included alongside this file). Requires the same GOOGLE_SERVICE_ACCOUNT_JSON and
// GOOGLE_SHEETS_ID env vars as the other Sheets functions.

const { google } = require("googleapis");

async function authenticateSheets() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// Same primary/fallback logic as api/fetch-exchange-rate.js. Duplicated rather than
// imported — Vercel serverless functions are deployed as independent bundles, and
// duplicating ~40 lines here is simpler and more reliable than setting up a shared
// module path across functions.
async function fetchFromBankIndonesia() {
  const today = new Date().toISOString().split("T")[0];
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 5);
  const start = startDate.toISOString().split("T")[0];

  const url = `https://www.bi.go.id/biwebservice/wskursbi.asmx/getSubKursLokal3?mts=USD&startdate=${start}&enddate=${today}`;
  const response = await fetch(url, { headers: { "User-Agent": "Integra-Dashboard" } });
  if (!response.ok) throw new Error(`BI web service returned HTTP ${response.status}`);

  const xml = await response.text();
  const jualMatches = [...xml.matchAll(/<jual[^>]*>([\d.]+)<\/jual[^>]*>/gi)];
  const beliMatches = [...xml.matchAll(/<beli[^>]*>([\d.]+)<\/beli[^>]*>/gi)];
  const dateMatches = [...xml.matchAll(/<tanggal[^>]*>([^<]+)<\/tanggal[^>]*>/gi)];
  if (!jualMatches.length || !beliMatches.length) throw new Error("No jual/beli rate found in BI response");

  const jual = parseFloat(jualMatches[jualMatches.length - 1][1]);
  const beli = parseFloat(beliMatches[beliMatches.length - 1][1]);
  if (!jual || !beli) throw new Error("BI rate values did not parse as numbers");

  return { rate: (jual + beli) / 2, date: dateMatches.length ? dateMatches[dateMatches.length - 1][1].trim() : today, source: "Bank Indonesia (JISDOR)" };
}

async function fetchFromFallback() {
  const response = await fetch("https://open.er-api.com/v6/latest/USD", { headers: { "User-Agent": "Integra-Dashboard" } });
  if (!response.ok) throw new Error(`Fallback returned HTTP ${response.status}`);
  const data = await response.json();
  const idRate = data?.rates?.IDR;
  if (!idRate) throw new Error("IDR rate missing from fallback response");
  return { rate: idRate, date: new Date().toISOString().split("T")[0], source: "open.er-api.com (fallback)" };
}

export default async function handler(req, res) {
  // Vercel signs cron requests with this header — reject anything else so this can't be
  // triggered by a random public request hitting the URL directly.
  const authHeader = req.headers["authorization"];
  if (process.env.CRON_SECRET && authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  let result;
  let biError = null;
  try {
    result = await fetchFromBankIndonesia();
  } catch (error) {
    biError = error.message;
    try {
      result = await fetchFromFallback();
    } catch (fallbackError) {
      console.error("Cron: both exchange rate sources failed", biError, fallbackError.message);
      return res.status(500).json({ error: "Both sources failed", details: `BI: ${biError} | Fallback: ${fallbackError.message}` });
    }
  }

  try {
    const sheetsId = process.env.GOOGLE_SHEETS_ID;
    if (!sheetsId) return res.status(500).json({ error: "GOOGLE_SHEETS_ID env var not set" });
    const sheets = await authenticateSheets();

    // Read current history, prepend today's rate (skip if today's date is already
    // present — avoids duplicate rows if the cron somehow fires twice in a day), then
    // write the full list back. Matches exactly what the dashboard's own
    // writeExchangeRateToSheets does client-side, for consistency.
    const existing = await sheets.spreadsheets.values.get({ spreadsheetId: sheetsId, range: "ExchangeRates!A:C" });
    const rows = existing.data.values || [["Date", "Rate (IDR/USD)", "Source"]];
    const headers = rows[0];
    const dataRows = rows.slice(1).filter((r) => r[0] !== result.date);
    const newRows = [headers, [result.date, String(Math.round(result.rate)), result.source], ...dataRows];

    await sheets.spreadsheets.values.clear({ spreadsheetId: sheetsId, range: "ExchangeRates!A:Z" });
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetsId, range: "ExchangeRates!A1", valueInputOption: "RAW", requestBody: { values: newRows },
    });

    res.status(200).json({
      success: true, date: result.date, rate: Math.round(result.rate), source: result.source,
      ...(biError ? { note: `BI unavailable (${biError}), used fallback` } : {}),
    });
  } catch (error) {
    console.error("Cron: Sheets write failed", error);
    res.status(500).json({ error: "Fetched rate successfully but failed to write to Sheets", details: error.message });
  }
}
