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
    // The old free endpoint (api.exchangerate-api.com/v4/latest) has been deprecated —
    // exchangerate-api.com restructured to a paid v6 API requiring a key. Their current
    // free, no-key, no-signup open-access endpoint lives at a different domain:
    // open.er-api.com. Verified working as of this update. If it ever stops working
    // again, the response just needs to be shaped as { rates: { IDR: number } }, so
    // swapping providers only requires changing this URL.
    const response = await fetch("https://open.er-api.com/v6/latest/USD", {
      headers: { "User-Agent": "Integra-Dashboard" },
    });

    if (!response.ok) {
      const bodyText = await response.text().catch(() => "");
      throw new Error(`Upstream fetch failed: ${response.status}${bodyText ? " — " + bodyText.slice(0, 200) : ""}`);
    }

    const data = await response.json();
    const idRate = data?.rates?.IDR;
    if (!idRate) throw new Error("IDR rate missing from response — upstream API shape may have changed");

    const today = new Date().toISOString().split("T")[0];

    res.status(200).json({
      date: today,
      rate: Math.round(idRate),
      source: "open.er-api.com",
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    console.error("Exchange rate fetch error:", error);
    res.status(500).json({ error: "Failed to fetch exchange rate", details: error.message });
  }
}
