// api/fetch-exchange-rate.js
// Vercel serverless function — fetches the daily USD -> IDR exchange rate server-side,
// avoiding the CORS issues a direct browser fetch to a third-party API would hit.
// Called by the dashboard once per day (see the useEffect in IntegraDashboard.jsx that
// checks localStorage's "lastExRateFetch" before calling this), and on every successful
// fetch it pushes the result straight to the ExchangeRates Sheet tab.
//
// Deploy: place this file at api/fetch-exchange-rate.js in your repo root (same level as
// your package.json). Vercel auto-detects anything under /api as a serverless function —
// no extra configuration needed.

// PRIMARY: Bank Indonesia's own JISDOR (Jakarta Interbank Spot Dollar Rate) — the
// official Indonesian central bank benchmark for USD/IDR, free, no API key, published
// daily on business days at 08:00 WIB. This is the authoritative source for an
// Indonesian mining operation's own financial reporting, not just a generic global
// aggregator.
//
// Their web service is an older ASMX/SOAP-style endpoint, but supports plain GET with
// query params. It returns XML rather than JSON — parsed below with regex rather than a
// full XML library (avoids adding another npm dependency, and is more tolerant of exact
// tag-name variations than a strict DOM query would be). If BI's schema or availability
// ever changes in a way this can't parse, it automatically falls through to the
// FALLBACK source below rather than failing outright.
async function fetchFromBankIndonesia() {
  const today = new Date().toISOString().split("T")[0];
  // JISDOR isn't published on weekends/holidays — look back up to 5 days for the most
  // recent business-day rate if today's isn't out yet (e.g. early morning before 08:00
  // WIB, or today is a weekend).
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 5);
  const start = startDate.toISOString().split("T")[0];

  const url = `https://www.bi.go.id/biwebservice/wskursbi.asmx/getSubKursLokal3?mts=USD&startdate=${start}&enddate=${today}`;
  const response = await fetch(url, { headers: { "User-Agent": "Integra-Dashboard" } });
  if (!response.ok) throw new Error(`BI web service returned HTTP ${response.status}`);

  const xml = await response.text();
  // Extract every <jual>/<beli> (sell/buy) pair present, tolerant of tag-name suffixes
  // (e.g. <jual_c>) some versions of this service have used historically. Takes the
  // LAST pair in the response, since results are chronological and we want the most
  // recent business day within the lookback window.
  const jualMatches = [...xml.matchAll(/<jual[^>]*>([\d.]+)<\/jual[^>]*>/gi)];
  const beliMatches = [...xml.matchAll(/<beli[^>]*>([\d.]+)<\/beli[^>]*>/gi)];
  const dateMatches = [...xml.matchAll(/<tanggal[^>]*>([^<]+)<\/tanggal[^>]*>/gi)];
  if (!jualMatches.length || !beliMatches.length) throw new Error("No jual/beli rate found in BI response — schema may have changed or no data in range");

  const jual = parseFloat(jualMatches[jualMatches.length - 1][1]);
  const beli = parseFloat(beliMatches[beliMatches.length - 1][1]);
  if (!jual || !beli) throw new Error("BI rate values did not parse as numbers");

  const midRate = (jual + beli) / 2;
  const rateDate = dateMatches.length ? dateMatches[dateMatches.length - 1][1].trim() : today;

  return { rate: midRate, date: rateDate, source: "Bank Indonesia (JISDOR)" };
}

// FALLBACK: used only if Bank Indonesia's service is unreachable or its response can't
// be parsed. exchangerate-api.com's old free v4 endpoint was deprecated in favor of a
// paid v6 API — their genuinely free, no-key, no-signup endpoint now lives at this
// different domain, open.er-api.com.
async function fetchFromFallback() {
  const response = await fetch("https://open.er-api.com/v6/latest/USD", {
    headers: { "User-Agent": "Integra-Dashboard" },
  });
  if (!response.ok) {
    const bodyText = await response.text().catch(() => "");
    throw new Error(`Fallback (open.er-api.com) returned HTTP ${response.status}${bodyText ? " — " + bodyText.slice(0, 150) : ""}`);
  }
  const data = await response.json();
  const idRate = data?.rates?.IDR;
  if (!idRate) throw new Error("IDR rate missing from fallback response");
  return { rate: idRate, date: new Date().toISOString().split("T")[0], source: "open.er-api.com (fallback)" };
}

export default async function handler(req, res) {
  let result;
  let biError = null;

  try {
    result = await fetchFromBankIndonesia();
  } catch (error) {
    biError = error.message;
    console.error("Bank Indonesia fetch failed, falling back:", error.message);
    try {
      result = await fetchFromFallback();
    } catch (fallbackError) {
      console.error("Fallback fetch also failed:", fallbackError.message);
      return res.status(500).json({
        error: "Failed to fetch exchange rate from both sources",
        details: `Bank Indonesia: ${biError} | Fallback: ${fallbackError.message}`,
      });
    }
  }

  res.status(200).json({
    date: result.date,
    rate: Math.round(result.rate),
    source: result.source,
    timestamp: new Date().toISOString(),
    ...(biError ? { note: `Bank Indonesia unavailable (${biError}), used fallback source` } : {}),
  });
}
