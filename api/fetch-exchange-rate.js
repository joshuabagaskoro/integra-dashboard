// api/fetch-exchange-rate.js
// Vercel serverless function — fetches the daily USD -> IDR exchange rate server-side,
// avoiding the CORS issues a direct browser fetch to a third-party API would hit.
// Called by the dashboard once per day (see the useEffect in IntegraDashboard.jsx that
// checks localStorage's "lastExRateFetch" before calling this).
//
// Deploy: place this file at api/fetch-exchange-rate.js in your repo root (same level as
// your package.json). Vercel auto-detects anything under /api as a serverless function —
// no extra configuration needed.

export default async function handler(req, res) {
  try {
    // exchangerate-api.com's open endpoint — free, no API key required for this route.
    // If this ever stops working (rate limits, API changes), swap the URL for another
    // free provider; the response just needs to end up shaped as { rates: { IDR: number } }
    // or be adapted below.
    const response = await fetch("https://api.exchangerate-api.com/v4/latest/USD", {
      headers: { "User-Agent": "Integra-Dashboard" },
    });

    if (!response.ok) throw new Error(`Upstream fetch failed: ${response.status}`);

    const data = await response.json();
    const idRate = data?.rates?.IDR;
    if (!idRate) throw new Error("IDR rate missing from response");

    const today = new Date().toISOString().split("T")[0];

    res.status(200).json({
      date: today,
      rate: Math.round(idRate),
      source: "exchangerate-api.com",
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    console.error("Exchange rate fetch error:", error);
    res.status(500).json({ error: "Failed to fetch exchange rate" });
  }
}
