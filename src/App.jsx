import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  Package, Ship, Calendar, TrendingUp, AlertTriangle, CheckCircle2,
  Layers, ChevronDown, Gauge, Upload, X,
  Search, Plus, Trash2, Lock, Unlock, LayoutGrid, FileUp, MapPin, FileText, Printer, FileDown, DollarSign, History, LogOut, Settings, Menu, RefreshCw, MessageSquare
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

/* ----------------------------- constants ----------------------------- */

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const TODAY = new Date();

/* Bump this ONLY when DEFAULT_DOMES or DEFAULT_BARGES is edited with new Excel-sourced
 * data (stock updates, new barge plans). Do NOT change it for feature/UI/logic edits that
 * don't touch the underlying data — that's the whole point of this timestamp. */
const DATA_LAST_UPDATED = "2026-08-01";

const DEFAULT_SETTINGS = {
  bargeSize: 10500,
  targetGrade: 1.35,
  tolerance: 0.03,
  totalQuota: 618000,
  planTarget: 54,
};

const PRODUCTION_TARGETS_2026 = {
  "IMN-1": { targetWMT: 150000, targetNi: 1.65 },
  "IMN-2": { targetWMT: 210000, targetNi: 1.45 },
  "IMN-3": { targetWMT: 120000, targetNi: 1.25 },
  "IMN-4": { targetWMT: 80000, targetNi: 1.35 },
};

// Lab-assay fields checked for missing data (0%) before finalizing a barge or on the
// Check Status modal. Al2O3 deliberately excluded — it's never actually collected by
// the lab, so checking it would falsely flag every single dome.
const LAB_FIELDS = [["ni", "Ni"], ["fe", "Fe"], ["co", "Co"], ["sio2", "SiO2"], ["mgo", "MgO"], ["simg", "Si:Mg"]];

// Client-side credentials — fine for an internal team tool, not for external/production
// use. Session token lives in sessionStorage (clears on browser close) rather than
// localStorage, per the "logout when browser closes" requirement.
// Full sub-feature structure for the granular (dev-only) feature configuration UI.
// IMPORTANT: not every key here actually gates something in the real UI yet — some map
// onto UI elements that don't exist as separate toggleable pieces (e.g. Stock's search
// bar isn't a removable component, it's baked into the table). WIRED_SUBFEATURES below
// lists exactly which keys are genuinely connected to conditional rendering. Toggling an
// unwired key still saves correctly to the Sheet — it just won't visibly change anything
// yet. This is intentional and documented rather than shipping fake toggles.
const FEATURE_STRUCTURE = {
  Stock: {
    label: "Stock Management",
    features: [
      { key: "Stock_ViewInventory", label: "View Inventory" },
      { key: "Stock_SearchDomes", label: "Search & Filter Domes" },
      { key: "Stock_ContractorFilter", label: "Filter by Contractor" },
      { key: "Stock_ProductionTargets", label: "View Production Targets" },
      { key: "Stock_ImportExcel", label: "Import Excel Updates" },
      { key: "Stock_ExportData", label: "Export Data" },
    ],
  },
  Barging: {
    label: "Barging Plan",
    features: [
      { key: "Barging_ViewBarges", label: "View Barges" },
      { key: "Barging_CreateManual", label: "Create Barge (Manual)" },
      { key: "Barging_CreateExcel", label: "Create Barge (Excel)" },
      { key: "Barging_EditBarge", label: "Edit Barge Details" },
      { key: "Barging_FinalizeBarges", label: "Finalize Barges" },
      { key: "Barging_LoadingProgress", label: "Loading Progress Tracker" },
      { key: "Barging_BatchQueue", label: "Batch Queue Auto-Sync" },
    ],
  },
  Timeline: {
    label: "Timeline",
    features: [
      { key: "Timeline_BargeTimeline", label: "Barge Timeline View" },
      { key: "Timeline_LoadingProgress", label: "Loading Progress" },
      { key: "Timeline_FilterBy", label: "Filter by Date/Status" },
    ],
  },
  Financials: {
    label: "Financials",
    features: [
      { key: "Financials_HMAPrice", label: "HPM Price Display" },
      { key: "Financials_ExchangeRate", label: "Exchange Rate Display" },
      { key: "Financials_RoyaltyCalculator", label: "Royalty Calculator" },
      { key: "Financials_BargeBreakdown", label: "Barge Breakdown Table" },
      { key: "Financials_ExportCSV", label: "Export CSV" },
    ],
  },
  LoginLog: {
    label: "Login Log",
    features: [
      { key: "LoginLog_ViewHistory", label: "View Login History" },
      { key: "LoginLog_FilterByUser", label: "Filter by User" },
      { key: "LoginLog_FilterByStatus", label: "Filter by Status" },
      { key: "LoginLog_ExportCSV", label: "Export CSV" },
    ],
  },
  Settings: {
    label: "Settings",
    features: [
      { key: "Settings_ExportToSheets", label: "Export to Google Sheets" },
      { key: "Settings_AccountManagement", label: "Account Management" },
      { key: "Settings_FeatureConfiguration", label: "Feature Configuration" },
      { key: "Settings_AccountInfo", label: "Account Info Display" },
      { key: "Settings_Logout", label: "Logout Button" },
    ],
  },
  ChatAssistant: {
    label: "Chat Assistant",
    features: [{ key: "ChatAssistant_Widget", label: "Loading Report Widget" }],
  },
};

// Sub-features that are ACTUALLY wired to real conditional rendering right now. Anything
// not in this set is genuinely saved/loaded correctly but doesn't hide/show anything yet
// — see the comment on FEATURE_STRUCTURE above.
//
// Settings_AccountManagement and Settings_FeatureConfiguration are deliberately NOT
// wired: only the single "dev" account can ever see that section at all (it's gated by
// username, not a flag), so toggling either off would only ever affect dev's own access
// to the tools that let them fix a bad toggle — the same self-lockout risk fixed
// elsewhere in Settings. Not applicable until there's more than one dev-tier account.
const WIRED_SUBFEATURES = new Set([
  "Settings_ExportToSheets",
  "Settings_Logout",
  "ChatAssistant_Widget",
  "Stock_ImportExcel",
  "Barging_CreateManual",
  "Barging_CreateExcel",
  "Barging_FinalizeBarges",
]);

const CREDENTIALS = {
  operation: { password: "opsintegra2026", role: "operation" },
  integra: { password: "admin-imn", role: "admin" },
  dev: { password: "test123", role: "admin" }, // Test account — for trying new features before wider rollout
};
const SESSION_MAX_AGE_MS = 60 * 60 * 1000; // 1 hour

// HPM (Harga Patokan Mineral) is tracked manually here, updated directly from whatever
// figure Integra is actually given each period — no formula calculation. Deliberately
// simpler than deriving HPM from HMA + Corrective Factors: the person confirmed this is
// the workflow they want.
const DEFAULT_HPM_HISTORY = [
  { date: "2026-07-15", price: 53.60, unit: "USD/WMT" },
  { date: "2026-07-01", price: 56.58, unit: "USD/WMT" },
  { date: "2026-06-15", price: 59.52, unit: "USD/WMT" },
];
const DEFAULT_EXCHANGE_RATE_HISTORY = [
  { date: "2026-07-29", rate: 17990, source: "Market mid-rate (Yahoo Finance / Pluang), verified via search" },
  { date: "2026-07-28", rate: 18010, source: "Market mid-rate (Yahoo Finance / Pluang), verified via search" },
];
const ROYALTY_TARIFF = 0.15; // flat 15% — see build note about bracket-based tariffs

const PALETTE = [
  "#22D3B8", "#F5B841", "#9B8CFF", "#FF7A7A", "#60A5FA", "#4ADE80",
  "#F472B6", "#FB923C", "#38BDF8", "#A3E635", "#E879F9", "#FBBF24",
  "#2DD4BF", "#F87171", "#818CF8", "#34D399",
];

function colorFor(id) {
  let h = 0;
  const s = String(id);
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  return PALETTE[h % PALETTE.length];
}

function fmtShortDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return iso;
  return `${d} ${MONTHS[m - 1]} ${y}`;
}

// Legend points from the site KMZ (734.kmz), matched to the "Location" column values.
const SITE_POINTS = [
  { name: "IMN-1-SELATAN", lat: -4.4439197, lng: 122.4930222 },
  { name: "IMN-2-UTARA", lat: -4.4352934, lng: 122.4914483 },
  { name: "IMN-2-SELATAN", lat: -4.4430854, lng: 122.4931939 },
  { name: "IMN-3-UTARA", lat: -4.4359619, lng: 122.4898015 },
  { name: "IMN-4-SELATAN", lat: -4.4507929, lng: 122.4930962 },
  { name: "JETTY", lat: -4.4544005, lng: 122.4921219, isJetty: true },
];

/* ----------------------------- embedded site data -----------------------------
 * Imported from INTEGRA_Google_Sheets_Template (Dome Inventory Existing +
 * Dome Production IMN-1..4), as of 27 July 2026. Replace via the Import button
 * once fresher data is exported from Google Sheets.
 */
// Sheets is now the single source of truth for domes and barges — no hardcoded
// snapshot baked into the app anymore. Every previous update required editing this
// file directly, which meant the embedded data went stale the moment ANYTHING
// changed in Sheets, and created exactly the "which one is real" confusion this
// was meant to prevent. The app now shows an explicit empty/syncing state on first
// load until the initial Sheets sync completes, rather than silently falling back
// to old numbers that look plausible but aren't current.
const DEFAULT_DOMES = [
];

/* ----------------------------- actual barge data (finalized) -----------------------------
 * Real barges shipped, imported from BARGE_PLAN_01..08.xlsx (Date / Barge Name / Tugboat
 * Name / Dome ID / WMT). Sorted by date, numbered 1-8. Marked finalized so their tonnage is
 * treated as already accounted for. Ni grades pulled from DEFAULT_DOMES at import time.
 * "DOME 391" mapped to "DOME 391  (D.30/AMR-S.2)" (only close match in dome database).
 */
const DEFAULT_BARGES = [
];

/* ----------------------------- helpers ----------------------------- */

const fmt = (n, d = 0) => Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });

// Ensures every dome carries an immutable initialStock baseline (set once, never decremented).
// Current .stock decreases as barges finalize; initialStock - stock = stock already barged out.
function withInitialStock(list) {
  return list.map((d) => (d.initialStock !== undefined ? d : { ...d, initialStock: d.stock }));
}

// The single source of truth for "how much of this dome is actually left." Given a
// dome's INITIAL/gross stock (the one number that should never be second-guessed — it's
// what was physically received) and the current barges, ALWAYS derives current remaining
// stock fresh from what's actually been consumed by FINALIZED barges, rather than
// trusting a separately-stored "current stock" figure that could be stale or wrong.
//
// This matters specifically for Google Sheets sync: if a dome's stock is ever updated by
// editing the Sheet directly (or by any path other than this app's Import button), the
// Sheet's "Stock (WMT)" column could easily end up holding the raw/gross figure instead
// of a properly-netted current one — nobody ran the netting logic on it. Blindly trusting
// that column on every sync pull would re-introduce exactly that stale/wrong number into
// the dashboard. So the sync pull below ignores "Stock (WMT)" for computation purposes
// entirely and only ever trusts "Initial Stock (WMT)", deriving current stock fresh
// every time from real barge consumption. This makes stock tracking self-correcting
// regardless of how or where the underlying Sheet data got updated.
function reconcileStock(domeId, grossInitialStock, barges) {
  const used = barges
    .filter((b) => b.finalized)
    .reduce((sum, b) => sum + (b.sources || []).filter((s) => s.id === domeId).reduce((s, x) => s + x.amt, 0), 0);
  if (used <= 0) return { initialStock: grossInitialStock, stock: grossInitialStock };
  if (used > grossInitialStock) return { initialStock: grossInitialStock + (used - grossInitialStock), stock: 0 };
  return { initialStock: grossInitialStock, stock: grossInitialStock - used };
}

// Client-side regex parser for operation team loading reports — genuinely just pattern
// matching, no API call of any kind. (The build spec's own overview mentioned "Claude
// API Parser," but its implementation section and troubleshooting notes both confirm
// this is regex-only — free, instant, no keys, no serverless function. Labeling it as
// an AI/Claude parser anywhere in the UI would be inaccurate, so all such references
// were corrected to just "Parse Report" throughout.)
function parseLoadingReportText(text) {
  try {
    if (!text || text.length === 0) return { error: "Text is empty" };

    const result = {
      shipmentNumber: null, qtyOnBoard: null, progressPercent: null, balanceDue: null, reportDate: null,
    };

    const shipmentMatch = text.match(/(?:Shipment|Barge|BG\.?)\s*(?:#)?\s*(\d+)/i);
    if (shipmentMatch) result.shipmentNumber = parseInt(shipmentMatch[1]);

    // Indonesian-style numbers assumed (dot = thousands separator, e.g. "1.800" = 1800),
    // matching the real report format this was built against.
    const qtyMatch = text.match(/(?:Qty\s+on\s+board|Qty|QTY)[:\s]+([0-9.,]+)\s*(?:wmt|WMT)?/i);
    if (qtyMatch) result.qtyOnBoard = parseFloat(qtyMatch[1].replace(/\./g, "").replace(/,/g, "."));

    const progressMatch = text.match(/(?:Progress|progress)[:\s]+(\d+)\s*%?/i);
    if (progressMatch) result.progressPercent = parseInt(progressMatch[1]);

    const balanceMatch = text.match(/(?:Blc|Balance\s+(?:due)?)[:\s]+([0-9.,]+)\s*(?:wmt|WMT)?/i);
    if (balanceMatch) result.balanceDue = parseFloat(balanceMatch[1].replace(/\./g, "").replace(/,/g, "."));

    const dateMatch = text.match(/(?:Tgl\.?|Date)?[:\s]*(\d{1,2})[-.\/](\d{1,2})[-.\/](\d{4})/);
    if (dateMatch) {
      const [, d, m, y] = dateMatch;
      result.reportDate = `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
    }

    const missingFields = Object.entries(result).filter(([, v]) => v === null).map(([k]) => k);
    return {
      ...result,
      missingFields: missingFields.length > 0 ? missingFields : null,
      confidence: Object.values(result).filter((v) => v !== null).length / 5,
    };
  } catch (error) {
    console.error("Parse error:", error);
    return { error: "Parse error: " + error.message };
  }
}

function defaultShipDate(no, planTarget) {
  const dayOfYear = Math.floor(((no - 0.5) * 365) / planTarget);
  const d = new Date(2026, 0, 1 + Math.max(0, dayOfYear));
  return d.toISOString().slice(0, 10);
}

function aggregateDomes(list) {
  const stock = list.reduce((s, d) => s + d.stock, 0);
  // Exclude domes with 0% for a given metric (unassayed placeholder, not a real reading)
  // from that metric's weighted average, so they don't drag it down.
  const w = (key) => {
    const valid = list.filter((d) => d[key] > 0);
    const validWeight = valid.reduce((s, d) => s + d.stock, 0);
    return validWeight > 0 ? valid.reduce((s, d) => s + d.stock * d[key], 0) / validWeight : 0;
  };
  return {
    domeCount: list.length, stock,
    avgNi: w("ni"), avgFe: w("fe"), avgCo: w("co"),
    avgSio2: w("sio2"), avgMgo: w("mgo"), avgAl2o3: w("al2o3"),
  };
}

/* ----------------------------- blend engine -----------------------------
 * Business rules:
 *  1) Prioritize consuming the lowest-Ni domes first.
 *  2) A single barge can draw from any number of existing-inventory contractors,
 *     but at most 2 production contractors (IMN-1..4).
 */

// Picks at most 2 production-source contractors to pair for a barge (the historical
// "2 contractors max" cap now applies only to production ore — see generateFromPool,
// which layers this on top of an unrestricted set of inventory contractors).
function pickProductionPair(pool, target) {
  const EPS = 1e-6;
  const cands = pool.filter((p) => p.remaining > EPS && p.grade > 0 && p.source === "production");
  if (!cands.length) return [];

  const anchor = cands.reduce((min, c) => (c.grade < min.grade ? c : min), cands[0]);
  const A = anchor.contractor;

  const byContractor = {};
  cands.forEach((c) => {
    if (!byContractor[c.contractor]) byContractor[c.contractor] = { stock: 0, ni: 0 };
    byContractor[c.contractor].stock += c.remaining;
    byContractor[c.contractor].ni += c.remaining * c.grade;
  });
  const others = Object.entries(byContractor)
    .filter(([cid]) => cid !== A)
    .map(([cid, v]) => ({ cid, avg: v.ni / v.stock, stock: v.stock }));

  const above = others.filter((o) => o.avg > target).sort((a, b) => (a.avg - target) - (b.avg - target));
  let B = null;
  if (above.length) B = above[0].cid;
  else if (others.length) B = others.sort((a, b) => b.avg - a.avg)[0].cid;

  return B && B !== A ? [A, B] : [A];
}

// Fills one barge using only domes from the given contractor set. Ni content is bounded
// on both sides: never below `target` (the floor) and never above `target + 0.2` (the
// ceiling). Highs and lows are blended in the exact ratio that lands on `target` for
// each chunk (lowest-Ni material used as the primary filler, per the existing priority
// rule), which keeps the cumulative average safely inside the band without depending on
// pre-banked capacity. Any leftover single-sided fill (only highs or only lows left) is
// bounded by the ceiling headroom / floor surplus banked from earlier balanced chunks.
// Extreme-grade domes (Ni < 1% or Ni > 1.5% — hard to source, so rationed across barges)
// are capped at EXTREME_GRADE_CAP WMT per barge, regardless of how much stock remains.
// If there isn't enough on-spec material left to reach bargeSize within both bounds, the
// barge is intentionally left incomplete rather than filled out of spec.
const EXTREME_GRADE_CAP = 500;
function fillOneBarge(pool, bargeSize, target, allowedContractors, simgTarget) {
  const EPS = 1e-6;
  const CEILING_PAD = 0.2; // max allowed excess over target, in Ni percentage points
  let filled = 0, ni = 0; // ni = running Ni mass, i.e. sum(amt * grade/100)
  const sources = [];
  let guard = 0;
  const t = target / 100;
  const maxT = (target + CEILING_PAD) / 100;

  // Si/Mg steering: lower score = better fit toward the stated target. Used only to
  // decide WHICH candidate dome to reach for first — it never excludes a dome from the
  // pool. Safe to freely reorder highs/lows by this instead of by Ni-closeness, because
  // the h/l ratio-blend math below self-corrects to hit exactly `target` Ni regardless
  // of which specific high/low pair is chosen, as long as one is above and one is below.
  const simgScore = simgTarget
    ? (p) => (simgTarget.op === "lte" ? (p.simg || 0) : -(p.simg || 0))
    : null;

  // Tracks how much of each dome has been used within THIS barge only — separate from
  // pool[i].remaining, which persists across barges. Extreme-grade domes are capped by
  // this per-barge tally, not by their overall remaining stock.
  const usedThisBarge = {};
  const isExtreme = (p) => p.grade < 1 || p.grade > 1.5;
  const availCap = (p) => {
    if (!isExtreme(p)) return p.remaining;
    const already = usedThisBarge[p.id] || 0;
    return Math.max(0, Math.min(p.remaining, EXTREME_GRADE_CAP - already));
  };
  const take = (p, amt) => {
    p.remaining -= amt;
    usedThisBarge[p.id] = (usedThisBarge[p.id] || 0) + amt;
  };

  while (filled < bargeSize - EPS && guard < 80) {
    guard++;
    const need = bargeSize - filled;
    const cands = pool.filter((p) => p.remaining > EPS && p.grade > 0 && allowedContractors.has(p.contractor) && availCap(p) > EPS);
    if (!cands.length) break;
    const neutrals = cands.filter((p) => Math.abs(p.grade - target) <= EPS);
    const highs = cands.filter((p) => p.grade > target)
      .sort(simgScore ? (a, b) => simgScore(a) - simgScore(b) : (a, b) => a.grade - b.grade); // Si/Mg fit first when targeted, else closest-to-target
    const lows = cands.filter((p) => p.grade < target)
      .sort(simgScore ? (a, b) => simgScore(a) - simgScore(b) : (a, b) => a.grade - b.grade); // Si/Mg fit first when targeted, else lowest first

    // Floor surplus: how much below-target dilution room is banked before the
    // cumulative average would dip under the minimum target grade.
    const floorSurplus = ni - t * filled;
    // Ceiling headroom: how much above-target room is banked before the cumulative
    // average would breach target + 0.2%.
    const ceilHeadroom = maxT * filled - ni;

    if (neutrals.length) {
      const n = neutrals[0]; const amt = Math.min(need, availCap(n));
      sources.push({ id: n.id, amt, grade: n.grade }); take(n, amt); ni += amt * (n.grade / 100); filled += amt;
      continue;
    }

    if (highs.length && lows.length) {
      // Blend a high and low dome in the exact ratio that hits `target` for this chunk.
      // Landing on target keeps the cumulative average safely within [target, ceiling]
      // without needing any pre-banked headroom — this is what avoids a startup deadlock
      // where neither side could move first. Uses lowest-Ni material as the primary
      // filler (existing priority rule), with just enough high-grade blended in to
      // reach the floor. If either side is supply-constrained (including by the
      // extreme-grade per-barge cap), both amounts scale down together so the ratio —
      // and thus the exact-target result — is preserved.
      const h = highs[0], l = lows[0];
      let f = (target - l.grade) / (h.grade - l.grade); f = Math.max(0, Math.min(1, f));
      const idealAmtH = f * need;
      const idealAmtL = (1 - f) * need;
      const capH = availCap(h), capL = availCap(l);
      let scale = 1;
      if (idealAmtH > EPS && idealAmtH > capH) scale = Math.min(scale, capH / idealAmtH);
      if (idealAmtL > EPS && idealAmtL > capL) scale = Math.min(scale, capL / idealAmtL);
      const amtH = idealAmtH * scale, amtL = idealAmtL * scale;
      if (amtH > EPS) { sources.push({ id: h.id, amt: amtH, grade: h.grade }); take(h, amtH); ni += amtH * (h.grade / 100); filled += amtH; }
      if (amtL > EPS) { sources.push({ id: l.id, amt: amtL, grade: l.grade }); take(l, amtL); ni += amtL * (l.grade / 100); filled += amtL; }
      if (amtH <= EPS && amtL <= EPS) break;
      continue;
    }

    if (highs.length) {
      // No low-grade partner left to balance against — use as much of this high-grade
      // dome as the banked ceiling headroom (and any extreme-grade cap) allows, so
      // on-spec-but-elevated ore isn't wasted.
      const h = highs[0];
      let amtH = Math.min(need, availCap(h));
      if (h.grade / 100 > maxT) {
        const maxByCeiling = Math.max(0, ceilHeadroom / (h.grade / 100 - maxT));
        amtH = Math.min(amtH, maxByCeiling);
      }
      if (amtH <= EPS) break; // no headroom left to add more without breaching the ceiling — stop here
      sources.push({ id: h.id, amt: amtH, grade: h.grade }); take(h, amtH); ni += amtH * (h.grade / 100); filled += amtH;
      continue;
    }

    if (lows.length) {
      // No above-target material left — only dip into the banked floor surplus, never
      // past it (and never past the extreme-grade per-barge cap either).
      const l = lows[0];
      const maxByFloor = Math.max(0, floorSurplus / (t - l.grade / 100));
      const amt = Math.min(need, availCap(l), maxByFloor);
      if (amt <= EPS) break; // no room left to dilute without breaching the floor — stop here
      sources.push({ id: l.id, amt, grade: l.grade }); take(l, amt); ni += amt * (l.grade / 100); filled += amt;
      continue;
    }

    break;
  }
  const merged = {};
  sources.forEach((s) => { if (!merged[s.id]) merged[s.id] = { id: s.id, amt: 0, grade: s.grade }; merged[s.id].amt += s.amt; });
  const grade = filled > 0 ? (ni / filled) * 100 : 0;
  return { sources: Object.values(merged), totalWMT: filled, grade };
}

function statusFor(b, bargeSize, target, tolerance) {
  if (b.totalWMT < bargeSize - 1) return b.totalWMT === 0 ? "unplanned" : "incomplete";
  const dev = Math.abs(b.grade - target);
  if (dev <= tolerance) return "exact";
  return b.grade > target ? "excess" : "deficit";
}

// Generates `count` new barges from a shared pool (mutates pool.remaining as it goes).
// Contractor limits differ by source: inventory contractors are unrestricted (a barge
// can blend across any number of them), while production contractors are still capped
// at 2 per barge, same as the original rule.
function generateFromPool(pool, count, bargeSize, target, tolerance, simgTarget) {
  const EPS = 1e-6;
  const out = [];
  for (let i = 0; i < count; i++) {
    const totalRemaining = pool.reduce((s, p) => s + p.remaining, 0);
    if (totalRemaining < 1e-6) { out.push({ sources: [], totalWMT: 0, grade: 0, status: "unplanned", pair: null }); continue; }

    const inventoryContractors = Array.from(new Set(
      pool.filter((p) => p.remaining > EPS && p.grade > 0 && p.source !== "production").map((p) => p.contractor)
    ));
    const productionPair = pickProductionPair(pool, target);
    const allowedList = Array.from(new Set([...inventoryContractors, ...productionPair]));

    if (!allowedList.length) { out.push({ sources: [], totalWMT: 0, grade: 0, status: "unplanned", pair: null }); continue; }

    const allowed = new Set(allowedList);
    const b = fillOneBarge(pool, bargeSize, target, allowed, simgTarget);
    b.status = statusFor(b, bargeSize, target, tolerance);
    b.pair = allowedList;
    out.push(b);
  }
  return out;
}

function poolFromDomes(domes, subtractBarges) {
  const pool = domes.map((d) => ({ id: d.id, contractor: d.contractor, grade: d.ni, remaining: d.stock, source: d.source || "inventory", simg: d.simg || 0 }));
  subtractBarges.forEach((b) => b.sources.forEach((s) => {
    const p = pool.find((x) => x.id === s.id);
    if (p) p.remaining = Math.max(0, p.remaining - s.amt);
  }));
  return pool;
}

/* ----------------------------- import parsing ----------------------------- */

// Files from this workflow use the European/Indonesian CSV convention: semicolon (;) as
// the field delimiter, comma (,) as the decimal separator — e.g. "1,47" means 1.47, not
// 147. Plain parseFloat("1,47") stops at the comma and silently returns just 1,
// truncating every decimal in the file. Convert comma to dot before parsing to fix this.
const cleanEuroNum = (v) => parseFloat(String(v ?? "").trim().replace(",", ".")) || 0;

function parseDomeCSV(csv, source, forcedContractor) {
  const rows = Papa.parse(csv, { skipEmptyLines: true, delimiter: ";" }).data;
  if (!rows.length) return [];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const [domeId, cid, stock, ni, fe, co, sio2, mgo, al2o3, simg, loc] = rows[i];
    if (!domeId) continue;
    out.push({
      id: domeId, contractor: forcedContractor || cid || "UNKNOWN",
      stock: cleanEuroNum(stock), ni: cleanEuroNum(ni), fe: cleanEuroNum(fe),
      co: cleanEuroNum(co), sio2: cleanEuroNum(sio2), mgo: cleanEuroNum(mgo),
      al2o3: cleanEuroNum(al2o3), simg: cleanEuroNum(simg), location: loc || "",
      source,
    });
  }
  return out;
}

// Reads a simple 2-column barge composition file: Dome ID, WMT (xlsx or csv)
function parseBargeComposition(rows) {
  // The real template (see the Instructions sheet in the .xlsx) puts DATE, BARGE NAME,
  // and TUGBOAT NAME as metadata rows BEFORE the actual "Dome ID"/"WMT" table — e.g.:
  //   Row 0: DATE | 2026-08-01
  //   Row 1: BARGE NAME | BG POLARIS 325
  //   Row 2: TUGBOAT NAME | TB JELAJAH 325
  //   Row 3: Dome ID | WMT      <- the real header
  //   Row 4+: dome data
  // Assuming row 0 is always the header (the old behavior) meant every row's Dome ID
  // lookup silently failed for any file using this metadata-first layout — exactly the
  // "no usable rows" failure. Scan for the actual header row instead of assuming its
  // position, and pull the metadata rows above it along the way.
  if (!rows.length) return { date: null, bargeName: null, tugboatName: null, sources: [] };

  let headerRowIdx = -1;
  let domeIdx = -1, wmtIdx = -1;
  const meta = { date: null, bargeName: null, tugboatName: null };

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i] || [];
    const cells = row.map((c) => String(c || "").toLowerCase().trim());
    const foundDome = cells.findIndex((c) => c.includes("dome"));
    const foundWmt = cells.findIndex((c) => c.includes("wmt") || c.includes("stock") || c.includes("weight"));
    if (foundDome !== -1 && foundWmt !== -1) {
      headerRowIdx = i; domeIdx = foundDome; wmtIdx = foundWmt;
      break;
    }
    // Not the header row — check if it's a metadata label row (label in col A, value in col B)
    const label = cells[0] || "";
    const value = row[1];
    if (value === undefined || value === "") continue;
    if (label.includes("date")) {
      meta.date = value instanceof Date ? value.toISOString().split("T")[0] : String(value).trim();
    } else if (label.includes("barge") && label.includes("name")) {
      meta.bargeName = String(value).trim();
    } else if (label.includes("tugboat") || (label.includes("tb") && label.includes("name"))) {
      meta.tugboatName = String(value).trim();
    }
  }

  if (headerRowIdx === -1) return { ...meta, sources: [] }; // couldn't find a Dome/WMT header anywhere in the first 15 rows

  const sources = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[domeIdx]) continue;
    const amt = cleanEuroNum(row[wmtIdx]);
    if (amt <= 0) continue;
    sources.push({ domeId: String(row[domeIdx]).trim(), amt });
  }
  return { ...meta, sources };
}

/* ----------------------------- small UI bits ----------------------------- */

function BlendRing({ sources, total, size = 64, centerLabel, centerSub }) {
  const buildConic = () => {
    if (!sources.length || total <= 0) return "conic-gradient(#2A3342 0% 100%)";
    let acc = 0;
    const stops = sources.map((s) => {
      const pct = (s.amt / total) * 100;
      const start = acc; acc += pct;
      return `${colorFor(s.id)} ${start}% ${acc}%`;
    });
    return `conic-gradient(${stops.join(", ")})`;
  };
  return (
    <div className="ring-wrap" style={{ width: size, height: size }}>
      <div className="ring-outer" style={{ background: buildConic(), width: size, height: size }} />
      <div className="ring-inner" style={{ width: size * 0.62, height: size * 0.62 }}>
        <span className="ring-label">{centerLabel}</span>
        {centerSub && <span className="ring-sub">{centerSub}</span>}
      </div>
    </div>
  );
}

function StatusBadge({ status }) {
  const map = {
    "exact": { text: "On target", cls: "badge-good", Icon: CheckCircle2 },
    "excess": { text: "Above target", cls: "badge-bad", Icon: AlertTriangle },
    "deficit": { text: "Below target", cls: "badge-bad", Icon: AlertTriangle },
    "unplanned": { text: "Insufficient stock", cls: "badge-bad", Icon: AlertTriangle },
    "incomplete": { text: "Incomplete", cls: "badge-warn", Icon: AlertTriangle },
  };
  const m = map[status] || map["incomplete"];
  const { Icon } = m;
  return <span className={`badge ${m.cls}`}><Icon size={12} strokeWidth={2.5} />{m.text}</span>;
}

function Kpi({ label, value, unit, accent }) {
  return (
    <div className="kpi">
      <div className="kpi-label">{label}</div>
      <div className={`kpi-value ${accent ? "kpi-" + accent : ""}`}>{value}<span className="kpi-unit">{unit}</span></div>
    </div>
  );
}

/* Generic multi-select checkbox dropdown.
 * selected semantics: null = everything included (default); [] = nothing included;
 * [ids] = explicit subset included. This lets the filter start "on" without needing
 * to know the full option list in advance. */
function MultiSelectDropdown({ options, selected, onChange, allLabel = "All", noneLabel = "None selected" }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const isChecked = (opt) => selected === null || selected.includes(opt);
  const toggle = (opt) => {
    if (selected === null) {
      onChange(options.filter((o) => o !== opt));
    } else if (selected.includes(opt)) {
      onChange(selected.filter((s) => s !== opt));
    } else {
      const next = [...selected, opt];
      onChange(next.length === options.length ? null : next);
    }
  };

  let displayText = allLabel;
  if (selected !== null) {
    if (selected.length === 0) displayText = noneLabel;
    else if (selected.length === 1) displayText = selected[0];
    else displayText = `${selected.length} selected`;
  }

  return (
    <div className="multiselect" ref={ref}>
      <button type="button" className="multiselect-btn" onClick={() => setOpen(!open)}>
        <span className="multiselect-btn-text">{displayText}</span>
        <ChevronDown size={13} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform 150ms" }} />
      </button>
      {open && (
        <div className="multiselect-panel">
          <div className="multiselect-actions">
            <button type="button" onClick={() => onChange(null)}>Select all</button>
            <button type="button" onClick={() => onChange([])}>Clear</button>
          </div>
          <div className="multiselect-list">
            {options.map((opt) => (
              <label key={opt} className="multiselect-item">
                <input type="checkbox" checked={isChecked(opt)} onChange={() => toggle(opt)} />
                <span>{opt}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

/* Searchable single-select dropdown for picking a dome by ID or contractor.
 * Replaces a long plain <select> (which requires scrolling through every dome)
 * with a text filter, used in the "Add dome manually" row on a barge. */
function SearchableDomeSelect({ options, value, onChange, placeholder }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const ref = useRef(null);
  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const selected = options.find((o) => o.id === value);
  const q = query.trim().toLowerCase();
  const filtered = q === "" ? options : options.filter((o) =>
    o.id.toLowerCase().includes(q) || (o.contractor || "").toLowerCase().includes(q)
  );

  return (
    <div className="dome-select" ref={ref}>
      <button type="button" className="dome-select-btn" onClick={() => setOpen(!open)}>
        <span className="dome-select-btn-text">
          {selected ? `${selected.id} (${fmt(selected.remaining)} WMT left, ${selected.grade}%)` : (placeholder || "+ Add dome…")}
        </span>
        <ChevronDown size={13} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform 150ms" }} />
      </button>
      {open && (
        <div className="dome-select-panel">
          <div className="dome-select-search">
            <Search size={13} />
            <input autoFocus type="text" placeholder="Search Dome ID or contractor…" value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
          <div className="dome-select-list">
            {filtered.length === 0 && <div className="dome-select-empty">No matches</div>}
            {filtered.map((o) => (
              <div key={o.id} className="dome-select-item" onClick={() => { onChange(o.id); setOpen(false); setQuery(""); }}>
                <span className="dome-select-item-id">{o.id}</span>
                <span className="dome-select-item-meta">{o.contractor} · {fmt(o.remaining)} WMT · {o.grade}%</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function NavButton({ icon: Icon, label, active, onClick, mobile }) {
  return (
    <button className={`navbtn ${active ? "navbtn-active" : ""} ${mobile ? "navbtn-mobile" : ""}`} onClick={onClick}>
      <Icon size={mobile ? 19 : 16} strokeWidth={2.3} />
      <span>{label}</span>
    </button>
  );
}

/* Mobile hamburger nav — a single button that opens a dropdown listing every tab,
 * replacing the old bottom icon bar (which got cramped once Financials/Log/Settings
 * were added on top of the original 4 tabs). Closes on selection or on outside click,
 * same interaction pattern as MultiSelectDropdown/SearchableDomeSelect elsewhere. */
function MobileNavMenu({ tab, setTab, isAdmin, currentUser, handleLogout, userFeatures }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const items = [
    { key: "overview", label: "Overview", icon: LayoutGrid },
    ...(userFeatures.stock ? [{ key: "stock", label: "Stock", icon: Layers }] : []),
    ...(userFeatures.barging ? [{ key: "plan", label: "Barging Plan", icon: Ship }] : []),
    ...(userFeatures.timeline ? [{ key: "timeline", label: "Timeline", icon: Calendar }] : []),
    ...(isAdmin && userFeatures.financials ? [{ key: "financials", label: "Financials", icon: DollarSign }] : []),
    ...(isAdmin && userFeatures.loginLog ? [{ key: "log", label: "Login Log", icon: History }] : []),
    ...(userFeatures.settings || currentUser?.username === "dev" ? [{ key: "settings", label: "Settings", icon: Settings }] : []),
  ];

  const go = (key) => { setTab(key); setOpen(false); };

  return (
    <div className="mobile-nav-menu" ref={ref}>
      <button className="hamburger-btn" onClick={() => setOpen(!open)} aria-label="Menu">
        <Menu size={22} />
      </button>
      {open && (
        <div className="mobile-nav-panel glass">
          <div className="mobile-nav-user">
            <div className="mobile-nav-username">{currentUser?.username}</div>
            <div className="mobile-nav-role">{isAdmin ? "Administrator" : "Operations"}</div>
          </div>
          <div className="mobile-nav-list">
            {items.map((item) => (
              <button key={item.key} className={`mobile-nav-item ${tab === item.key ? "mobile-nav-item-active" : ""}`} onClick={() => go(item.key)}>
                <item.icon size={17} strokeWidth={2.3} />
                <span>{item.label}</span>
              </button>
            ))}
          </div>
          <button className="mobile-nav-logout" onClick={() => { setOpen(false); handleLogout(); }}>
            <LogOut size={16} /> <span>Logout</span>
          </button>
        </div>
      )}
    </div>
  );
}

function StatCard({ title, tag, stats, color }) {
  return (
    <div className="stat-card">
      <div className="stat-card-head">
        <span className="dot" style={{ background: color }} />
        <span className="stat-card-title">{title}</span>
        {tag && <span className="stat-card-tag">{tag}</span>}
      </div>
      <div className="stat-card-grid">
        <div><span className="stat-num">{fmt(stats.domeCount)}</span><span className="stat-lbl">Domes</span></div>
        <div><span className="stat-num">{fmt(stats.stock)}</span><span className="stat-lbl">WMT</span></div>
        <div><span className="stat-num">{fmt(stats.avgNi, 2)}</span><span className="stat-lbl">Ni %</span></div>
        <div><span className="stat-num">{fmt(stats.avgFe, 2)}</span><span className="stat-lbl">Fe %</span></div>
        <div><span className="stat-num">{fmt(stats.avgCo, 3)}</span><span className="stat-lbl">CO %</span></div>
        <div><span className="stat-num">{fmt(stats.avgSio2, 1)}</span><span className="stat-lbl">SiO2 %</span></div>
      </div>
    </div>
  );
}

/* ----------------------------- Site map (self-contained SVG, no external tiles) ----------------------------- */

// Concession boundary polygons, extracted from the site KMZ (734.kmz -> Area Features)
const SITE_POLYGONS = {
  
  "734 ALL": [[-4.421418,122.499439],[-4.421423,122.493310],[-4.425666,122.493307],[-4.425674,122.492387],[-4.431166,122.492393],[-4.431169,122.484110],[-4.432363,122.484109],[-4.432362,122.483081],[-4.434280,122.483080],[-4.434278,122.480610],[-4.442330,122.480614],[-4.442330,122.481443],[-4.443172,122.481442],[-4.443174,122.484390],[-4.443716,122.484389],[-4.443717,122.484939],[-4.445436,122.484938],[-4.445436,122.485749],[-4.446332,122.485749],[-4.446332,122.486722],[-4.447083,122.486721],[-4.447084,122.487388],[-4.452331,122.487394],[-4.452331,122.488557],[-4.455417,122.488861],[-4.455418,122.491439],[-4.454667,122.491439],[-4.454668,122.492386],[-4.453139,122.492387],[-4.453140,122.493721],[-4.439444,122.493721],[-4.439436,122.495686],[-4.436831,122.495688],[-4.436832,122.497139],[-4.433810,122.497141],[-4.433812,122.499439],[-4.421418,122.499439]],
  "734 EXISTING": [[-4.433750,122.487778],[-4.435417,122.487778],[-4.435417,122.488944],[-4.436694,122.488944],[-4.436694,122.489889],[-4.437833,122.489889],[-4.437833,122.492667],[-4.443389,122.492667],[-4.443389,122.491806],[-4.444000,122.491806],[-4.444000,122.489667],[-4.442194,122.489667],[-4.442194,122.485556],[-4.445194,122.485556],[-4.445194,122.486056],[-4.445806,122.486056],[-4.445806,122.486694],[-4.446361,122.486694],[-4.446361,122.487389],[-4.447278,122.487389],[-4.447278,122.491278],[-4.447472,122.491278],[-4.447472,122.491917],[-4.451444,122.491917],[-4.451444,122.491278],[-4.454389,122.491278],[-4.454389,122.492167],[-4.452944,122.492167],[-4.452944,122.493667],[-4.448667,122.493667],[-4.448667,122.493083],[-4.447167,122.493083],[-4.447167,122.492389],[-4.445167,122.492389],[-4.445167,122.493111],[-4.444583,122.493111],[-4.444583,122.493472],[-4.444278,122.493472],[-4.444278,122.494667],[-4.441667,122.494667],[-4.441667,122.494056],[-4.439306,122.494056],[-4.439306,122.494306],[-4.436167,122.494306],[-4.436167,122.495361],[-4.434083,122.495361],[-4.434083,122.494167],[-4.433806,122.494167],[-4.433806,122.493528],[-4.433000,122.493528],[-4.433000,122.493000],[-4.431833,122.493000],[-4.431833,122.490500],[-4.432889,122.490500],[-4.432889,122.490000],[-4.433750,122.490000],[-4.433750,122.487778]]
};

// Bounding box across both polygons + all legend points, with projection helpers
const MAP_BOUNDS = { latMin: -4.455418, latMax: -4.421418, lngMin: 122.48061, lngMax: 122.499439 };
const MAP_VB = { w: 400, h: 725, pad: 0.07 };

function projectLatLng(lat, lng) {
  const { latMin, latMax, lngMin, lngMax } = MAP_BOUNDS;
  const { w, h, pad } = MAP_VB;
  const usableW = w * (1 - 2 * pad), usableH = h * (1 - 2 * pad);
  const x = w * pad + ((lng - lngMin) / (lngMax - lngMin)) * usableW;
  const y = h * pad + ((latMax - lat) / (latMax - latMin)) * usableH;
  return [x, y];
}

function polyToPoints(pts) {
  return pts.map(([lat, lng]) => projectLatLng(lat, lng).join(",")).join(" ");
}

function SiteMap({ domes }) {
  const [hover, setHover] = useState(null);
  const [zoom, setZoom] = useState(1);
  const [panX, setPanX] = useState(0);
  const [panY, setPanY] = useState(0);
  const [isDragging, setIsDragging] = useState(false);
  const [dragStart, setDragStart] = useState(null);
  const mapSvgRef = useRef(null);
  const MIN_ZOOM = 1, MAX_ZOOM = 3;

  const locationStats = useMemo(() => {
    const map = {};
    domes.forEach((d) => {
      const loc = d.location;
      if (!loc) return;
      if (!map[loc]) map[loc] = { domes: 0, stock: 0, ni: 0, fe: 0, co: 0 };
      const s = map[loc];
      s.domes += 1; s.stock += d.stock; s.ni += d.stock * d.ni; s.fe += d.stock * d.fe; s.co += d.stock * d.co;
    });
    return map;
  }, [domes]);

  const maxStock = Math.max(...SITE_POINTS.map((p) => (locationStats[p.name] && locationStats[p.name].stock) || 0), 1);

  const handleZoom = (direction) => {
    setZoom((z) => Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, z + direction * 0.3)));
  };

  const viewBoxWidth = MAP_VB.w / zoom, viewBoxHeight = MAP_VB.h / zoom;
  const maxPanX = (MAP_VB.w - viewBoxWidth) / 2;
  const maxPanY = (MAP_VB.h - viewBoxHeight) / 2;
  const clampedPanX = Math.max(-maxPanX, Math.min(maxPanX, panX));
  const clampedPanY = Math.max(-maxPanY, Math.min(maxPanY, panY));
  const viewBoxX = (MAP_VB.w - viewBoxWidth) / 2 + clampedPanX;
  const viewBoxY = (MAP_VB.h - viewBoxHeight) / 2 + clampedPanY;

  const handleMouseDown = (e) => {
    if (!mapSvgRef.current) return;
    setIsDragging(true);
    setDragStart({ x: e.clientX, y: e.clientY, panX: clampedPanX, panY: clampedPanY });
  };

  const handleMouseMove = (e) => {
    if (!isDragging || !dragStart) return;
    const deltaX = e.clientX - dragStart.x;
    const deltaY = e.clientY - dragStart.y;
    const rect = mapSvgRef.current?.getBoundingClientRect();
    if (!rect) return;
    const panDeltaX = -(deltaX / rect.width) * viewBoxWidth;
    const panDeltaY = -(deltaY / rect.height) * viewBoxHeight;
    setPanX(dragStart.panX + panDeltaX);
    setPanY(dragStart.panY + panDeltaY);
  };

  const handleMouseUp = () => {
    setIsDragging(false);
    setDragStart(null);
  };

  useEffect(() => {
    if (!isDragging) return;
    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };
  }, [isDragging, dragStart, viewBoxWidth, viewBoxHeight]);

  return (
    <section className="glass panel">
      <div className="panel-head"><MapPin size={16} /><span>Site map — stock by stockpile location</span></div>
      <div className="site-map-wrap">
        <svg ref={mapSvgRef} viewBox={`${viewBoxX} ${viewBoxY} ${viewBoxWidth} ${viewBoxHeight}`} className={`site-map-svg ${isDragging ? "dragging" : ""}`} preserveAspectRatio="xMidYMid meet" onMouseDown={handleMouseDown} style={{ cursor: isDragging ? "grabbing" : "grab" }}>
          <defs>
            <pattern id="mapGrid" width="20" height="20" patternUnits="userSpaceOnUse">
              <path d="M 20 0 L 0 0 0 20" fill="none" stroke="rgba(255,255,255,0.035)" strokeWidth="1" />
            </pattern>
          </defs>
          <rect x="0" y="0" width={MAP_VB.w} height={MAP_VB.h} fill="url(#mapGrid)" />
          <polygon points={polyToPoints(SITE_POLYGONS["734 ALL"])} className="poly-all" />
          <polygon points={polyToPoints(SITE_POLYGONS["734 EXISTING"])} className="poly-existing" />

          {SITE_POINTS.map((p) => {
            const stats = locationStats[p.name] || { domes: 0, stock: 0, ni: 0, fe: 0, co: 0 };
            const [x, y] = projectLatLng(p.lat, p.lng);
            const color = p.isJetty ? "#C9A227" : colorFor(p.name);
            const r = p.isJetty ? 5.5 : 6 + 11 * (stats.stock / maxStock);
            return (
              <g key={p.name}
                onMouseEnter={() => setHover({ ...p, stats, x, y })}
                onMouseLeave={() => setHover(null)}
                style={{ cursor: "pointer" }}>
                <circle cx={x} cy={y} r={r + 6} fill={color} opacity={hover && hover.name === p.name ? 0.18 : 0.08} />
                <circle cx={x} cy={y} r={r} fill={color} fillOpacity="0.5" stroke={color} strokeWidth="1.6" />
                <circle cx={x} cy={y} r="2" fill={color} />
              </g>
            );
          })}
        </svg>

        <div className="compass">N ↑</div>

        <div className="zoom-controls">
          <button className="zoom-btn" onClick={() => handleZoom(1)} disabled={zoom >= MAX_ZOOM} title="Zoom in">+</button>
          <button className="zoom-btn" onClick={() => handleZoom(-1)} disabled={zoom <= MIN_ZOOM} title="Zoom out">−</button>
        </div>

        {hover && (
          <div className="map-tip-overlay" style={{ left: `${(hover.x / MAP_VB.w) * 100}%`, top: `${(hover.y / MAP_VB.h) * 100}%` }}>
            <b>{hover.name}</b>
            {hover.isJetty ? (
              <div>Loading / shipping point</div>
            ) : (
              <>
                <div>{hover.stats.domes} domes · {fmt(hover.stats.stock)} WMT</div>
                <div>{hover.stats.stock > 0 ? `Avg Ni ${fmt(hover.stats.ni / hover.stats.stock, 2)}%` : "No stock recorded"}</div>
              </>
            )}
          </div>
        )}
      </div>
      <div className="legend" style={{ marginTop: 10 }}>
        {SITE_POINTS.map((p) => (
          <span key={p.name} className="legend-item">
            <span className="dot" style={{ background: p.isJetty ? "#C9A227" : colorFor(p.name) }} />
            {p.name}{p.isJetty ? " (shipping point)" : ""}
          </span>
        ))}
        <span className="legend-item"><span className="dot" style={{ background: "transparent", border: "1px solid rgba(227,95,12,.5)" }} />734 EXISTING boundary</span>
      </div>
    </section>
  );
}


/* ----------------------------- Overview tab ----------------------------- */

function OverviewTab({ domes, barges, settings }) {
  const totalExisting = domes.reduce((s, d) => s + d.stock, 0);
  const actualBarged = barges.filter((b) => b.finalized).reduce((s, b) => s + b.totalWMT, 0);
  const quotaCoverage = (totalExisting / settings.totalQuota) * 100;
  const overallNi = totalExisting > 0 ? domes.reduce((s, d) => s + d.stock * d.ni, 0) / totalExisting : 0;

  const inventoryStats = useMemo(() => aggregateDomes(domes.filter((d) => d.source === "inventory")), [domes]);
  const productionByContractor = useMemo(() => {
    const prod = domes.filter((d) => d.source === "production");
    const contractors = Array.from(new Set(prod.map((d) => d.contractor))).sort();
    return contractors.map((c) => ({ id: c, ...aggregateDomes(prod.filter((d) => d.contractor === c)) }));
  }, [domes]);

  const ungraded = useMemo(() => domes.filter((d) => d.ni <= 0 && d.stock > 0), [domes]);
  const ungradedStock = ungraded.reduce((s, d) => s + d.stock, 0);

  const heroSources = [inventoryStats, ...productionByContractor].map((c, i) => ({ id: c.id || "Inventory", amt: c.stock, grade: c.avgNi }));

  return (
    <div className="stack">
      {ungradedStock > 0 && (
        <section className="glass banner">
          <AlertTriangle size={18} className="banner-icon" />
          <div>
            <div className="banner-title">{fmt(ungradedStock)} WMT across {ungraded.length} domes has no Ni lab result yet</div>
            <div className="banner-body">
              These domes are recorded with 0% Ni (untested), mostly in IMN-3 and IMN-4. They're excluded
              from auto-generated barging plans until real assay results are entered, so they won't
              silently drag blends down. They still count toward existing stock totals below.
            </div>
          </div>
        </section>
      )}

      <section className="hero glass">
        <BlendRing sources={heroSources} total={totalExisting} size={128} centerLabel={`${fmt(overallNi, 2)}%`} centerSub="blended Ni" />
        <div className="hero-stats">
          <div className="hero-title">2026 Barging Program Overview</div>
          <div className="hero-desc">{domes.length} domes on record, as of {TODAY.toDateString()}.</div>
          <div className="kpi-row">
            <Kpi label="2026 Quota" value={fmt(settings.totalQuota)} unit="WMT" />
            <Kpi label="Actual barged" value={fmt(actualBarged)} unit="WMT" accent="good" />
            <Kpi label="Existing stock" value={fmt(totalExisting)} unit="WMT" />
            <Kpi label="Quota coverage" value={fmt(quotaCoverage, 1)} unit="%" accent={quotaCoverage >= 100 ? "good" : "warn"} />
          </div>
        </div>
      </section>

      <SiteMap domes={domes} />

      <section className="glass panel">
        <div className="panel-head"><Gauge size={16} /><span>Existing stock vs. 2026 quota</span></div>
        <div className="stacked-bar">
          <div className="stacked-seg" style={{ width: `${Math.min(quotaCoverage, 100)}%`, background: "#E35F0C" }} />
        </div>
        <div className="legend">
          <span className="legend-item"><span className="dot" style={{ background: "#E35F0C" }} />{fmt(totalExisting)} WMT on hand</span>
          <span className="legend-item"><span className="dot" style={{ background: "#3A4256" }} />{fmt(Math.max(0, settings.totalQuota - totalExisting))} WMT still to be produced</span>
        </div>
      </section>

      <section className="glass panel">
        <div className="panel-head"><Package size={16} /><span>Inventory</span></div>
        <div className="stat-card-row">
          <StatCard title="Dome Inventory Existing" tag="Inventory" stats={inventoryStats} color="#E35F0C" />
        </div>
      </section>

      <section className="glass panel">
        <div className="panel-head"><Layers size={16} /><span>Production by contractor</span></div>
        <div className="stat-card-row">
          {productionByContractor.map((c) => (
            <StatCard key={c.id} title={c.id} tag="Production" stats={c} color={colorFor(c.id)} />
          ))}
        </div>
      </section>
    </div>
  );
}

/* ----------------------------- Stock tab ----------------------------- */

function StockTab({ domes }) {
  const [search, setSearch] = useState("");
  const [contractorFilter, setContractorFilter] = useState(null); // null = all; [] = none; [ids] = subset
  const [sourceFilter, setSourceFilter] = useState("all");
  const [sortKey, setSortKey] = useState("stock");
  const [sortDir, setSortDir] = useState(-1);
  const [inventoryOpen, setInventoryOpen] = useState(false);
  const [productionOpen, setProductionOpen] = useState(false);

  // Contractor list shown in the checklist follows the Source dropdown, same logic as the
  // Barging Plan generator: Existing Inventory -> inventory contractors only; Production ->
  // IMN-1..4 only; All sources -> everyone.
  const contractors = useMemo(() => {
    const visible = domes.filter((d) => sourceFilter === "all" || d.source === sourceFilter);
    return Array.from(new Set(visible.map((d) => d.contractor))).sort();
  }, [domes, sourceFilter]);

  // Reset contractor selection back to "all" whenever the Source filter changes, so a stale
  // selection from the other source group can't linger invisibly.
  useEffect(() => { setContractorFilter(null); }, [sourceFilter]);

  const filtered = useMemo(() => {
    let rows = domes.filter((d) =>
      d.stock > 0 &&
      (contractorFilter === null || contractorFilter.includes(d.contractor)) &&
      (sourceFilter === "all" || d.source === sourceFilter) &&
      (search === "" || d.id.toLowerCase().includes(search.toLowerCase()) || d.contractor.toLowerCase().includes(search.toLowerCase()))
    );
    rows = rows.map((d) => {
      const initialStock = d.initialStock !== undefined ? d.initialStock : d.stock;
      return { ...d, initialStock, stockOut: Math.max(0, initialStock - d.stock) };
    });
    rows = [...rows].sort((a, b) => (a[sortKey] > b[sortKey] ? 1 : -1) * sortDir);
    return rows;
  }, [domes, search, contractorFilter, sourceFilter, sortKey, sortDir]);

  const totalStock = filtered.reduce((s, d) => s + d.stock, 0);

  // Aggregate totals + weighted averages across whatever's currently filtered/checked —
  // same "exclude unassayed (0% Ni) domes from the average" rule used elsewhere, so a
  // pile of not-yet-tested domes doesn't drag the blended grade down artificially.
  const filteredTotals = useMemo(() => {
    let stockSum = 0;
    const acc = { ni: 0, fe: 0, co: 0, sio2: 0, mgo: 0, al2o3: 0, simg: 0 };
    const weight = { ni: 0, fe: 0, co: 0, sio2: 0, mgo: 0, al2o3: 0, simg: 0 };
    filtered.forEach((d) => {
      stockSum += d.stock;
      // Each field independently excludes domes where THAT field is 0 (unassayed) from
      // both the sum and its own weight — a dome unassayed for Fe shouldn't dilute the
      // Fe average just because it has a valid Co reading. Previously only Ni/Si:Mg did
      // this; Fe/Co/SiO2/MgO/Al2O3 used total stock as the denominator regardless,
      // understating their true average whenever any domes in the filter were unassayed.
      ["ni", "fe", "co", "sio2", "mgo", "al2o3", "simg"].forEach((key) => {
        if (d[key] > 0) { acc[key] += d.stock * d[key]; weight[key] += d.stock; }
      });
    });
    const avg = (key) => (weight[key] > 0 ? acc[key] / weight[key] : 0);
    return {
      stock: stockSum,
      ni: avg("ni"), fe: avg("fe"), co: avg("co"), sio2: avg("sio2"), mgo: avg("mgo"), al2o3: avg("al2o3"), simg: avg("simg"),
    };
  }, [filtered]);
  
  // Contractor summary: split by source (inventory vs production)
  const { inventorySummary, productionSummary } = useMemo(() => {
    const bySource = { inventory: {}, production: {} };
    domes.forEach((d) => {
      const source = d.source || "inventory";
      if (!bySource[source][d.contractor]) {
        bySource[source][d.contractor] = { total: 0, current: 0, domes: 0, source, niSum: 0, niWeight: 0, unassayed: 0 };
      }
      const domeTotal = d.initialStock !== undefined ? d.initialStock : d.stock;
      bySource[source][d.contractor].total += domeTotal;
      bySource[source][d.contractor].current += d.stock;
      bySource[source][d.contractor].domes += 1;
      // Exclude 0% Ni (unassayed placeholder) from the weighted average — both the
      // numerator and the denominator — so unassayed domes don't drag the average down.
      if (d.ni > 0) {
        bySource[source][d.contractor].niSum += domeTotal * d.ni;
        bySource[source][d.contractor].niWeight += domeTotal;
      } else {
        bySource[source][d.contractor].unassayed += domeTotal;
      }
    });
    
    const formatSummary = (sourceData) => 
      Object.entries(sourceData)
        .map(([contractor, data]) => ({
          contractor,
          totalAcquired: data.total,
          actualRemaining: data.current,
          usedInBarging: Math.max(0, data.total - data.current),
          capacityUsed: data.total > 0 ? Math.round(((data.total - data.current) / data.total) * 100) : 0,
          domeCount: data.domes,
          avgNi: data.niWeight > 0 ? data.niSum / data.niWeight : 0,
          unassayedWMT: data.unassayed
        }))
        .sort((a, b) => b.totalAcquired - a.totalAcquired);
    
    return {
      inventorySummary: formatSummary(bySource.inventory),
      productionSummary: formatSummary(bySource.production)
    };
  }, [domes]);

  const productionTargetsView = useMemo(() => {
    return Object.entries(PRODUCTION_TARGETS_2026).map(([contractor, target]) => {
      const actual = productionSummary.find((p) => p.contractor === contractor);
      const actualWMT = actual ? actual.totalAcquired : 0;
      const actualNi = actual ? actual.avgNi : 0;
      const unassayedWMT = actual ? actual.unassayedWMT : 0;
      const pctOfTarget = target.targetWMT > 0 ? Math.min(999, Math.round((actualWMT / target.targetWMT) * 100)) : 0;
      const niDelta = actualNi - target.targetNi;
      return { contractor, ...target, actualWMT, actualNi, unassayedWMT, pctOfTarget, niDelta };
    });
  }, [productionSummary]);
  const [targetsOpen, setTargetsOpen] = useState(true);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(sortDir * -1);
    else { setSortKey(key); setSortDir(-1); }
  };
  const Th = ({ k, children }) => (
    <th onClick={() => toggleSort(k)} className="sortable">{children}{sortKey === k ? (sortDir === 1 ? " ▲" : " ▼") : ""}</th>
  );

  return (
    <div className="stack">
      {inventorySummary.length > 0 && (
        <section className="glass panel">
          <div 
            className="panel-head panel-head-collapsible" 
            onClick={() => setInventoryOpen(!inventoryOpen)}
            style={{ cursor: 'pointer' }}
          >
            <Layers size={16} />
            <span>Stock overview — Existing Inventory</span>
            <ChevronDown size={18} style={{ marginLeft: 'auto', transform: inventoryOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 200ms ease' }} />
          </div>
          {inventoryOpen && (
            <div className="contractor-grid">
              {inventorySummary.map((c) => (
                <div key={c.contractor} className="contractor-card">
                  <div className="card-contractor">{c.contractor}</div>
                  <div className="card-row">
                    <span className="card-label">Total Acquired</span>
                    <span className="card-value">{fmt(c.totalAcquired)} WMT</span>
                  </div>
                  <div className="card-row">
                    <span className="card-label">Actual Remaining</span>
                    <span className="card-value card-highlight">{fmt(c.actualRemaining)} WMT</span>
                  </div>
                  <div className="card-row">
                    <span className="card-label">Used in Barging</span>
                    <span className="card-value card-used">{fmt(c.usedInBarging)} WMT</span>
                  </div>
                  <div className="card-row">
                    <span className="card-label">Capacity Used</span>
                    <span className="card-value">{c.capacityUsed}%</span>
                  </div>
                  <div className="card-row card-domes">
                    <span className="card-label">{c.domeCount} domes</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </section>
      )}

      {productionSummary.length > 0 && (
        <section className="glass panel">
          <div 
            className="panel-head panel-head-collapsible" 
            onClick={() => setProductionOpen(!productionOpen)}
            style={{ cursor: 'pointer' }}
          >
            <Ship size={16} />
            <span>Stock overview — Production (IMN-1–4)</span>
            <ChevronDown size={18} style={{ marginLeft: 'auto', transform: productionOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 200ms ease' }} />
          </div>
          {productionOpen && (
            <div className="contractor-grid">
              {productionSummary.map((c) => (
                <div key={c.contractor} className="contractor-card contractor-card-prod">
                  <div className="card-contractor card-contractor-prod">{c.contractor}</div>
                  <div className="card-row">
                    <span className="card-label">Total Acquired</span>
                    <span className="card-value">{fmt(c.totalAcquired)} WMT</span>
                  </div>
                  <div className="card-row">
                    <span className="card-label">Actual Remaining</span>
                    <span className="card-value card-highlight">{fmt(c.actualRemaining)} WMT</span>
                  </div>
                  <div className="card-row">
                    <span className="card-label">Used in Barging</span>
                    <span className="card-value card-used">{fmt(c.usedInBarging)} WMT</span>
                  </div>
                  <div className="card-row">
                    <span className="card-label">Capacity Used</span>
                    <span className="card-value">{c.capacityUsed}%</span>
                  </div>
                  <div className="card-row card-domes">
                    <span className="card-label">{c.domeCount} domes</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </section>
      )}

      <section className="glass panel">
        <div
          className="panel-head panel-head-collapsible"
          onClick={() => setTargetsOpen(!targetsOpen)}
          style={{ cursor: 'pointer' }}
        >
          <TrendingUp size={16} />
          <span>2026 Production Targets — IMN-1–4</span>
          <ChevronDown size={18} style={{ marginLeft: 'auto', transform: targetsOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 200ms ease' }} />
        </div>
        {targetsOpen && (
          <div className="targets-grid">
            {productionTargetsView.map((t) => (
              <div key={t.contractor} className="target-card">
                <div className="target-header">
                  <span className="target-contractor">{t.contractor}</span>
                  <span className={`target-pct ${t.pctOfTarget >= 100 ? "target-pct-good" : t.pctOfTarget >= 60 ? "target-pct-warn" : "target-pct-bad"}`}>
                    {t.pctOfTarget}%
                  </span>
                </div>
                <div className="target-bar-track">
                  <div className="target-bar-fill" style={{ width: `${Math.min(100, t.pctOfTarget)}%` }} />
                </div>
                <div className="target-row">
                  <span className="target-label">Produced</span>
                  <span className="target-value">{fmt(t.actualWMT)} / {fmt(t.targetWMT)} WMT</span>
                </div>
                <div className="target-row">
                  <span className="target-label">Ni Content</span>
                  <span className="target-value">
                    {t.actualNi > 0 ? fmt(t.actualNi, 2) : "—"}% <span className="target-ni-target">(target {fmt(t.targetNi, 2)}%)</span>
                  </span>
                </div>
                {t.actualNi > 0 && (
                  <div className={`target-ni-badge ${Math.abs(t.niDelta) <= 0.05 ? "ni-badge-good" : "ni-badge-warn"}`}>
                    {t.niDelta >= 0 ? "+" : ""}{fmt(t.niDelta, 2)}% vs target
                  </div>
                )}
                {t.unassayedWMT > 0 && (
                  <div className="target-unassayed-note">{fmt(t.unassayedWMT)} WMT not yet assayed (excluded from Ni%)</div>
                )}
              </div>
            ))}
          </div>
        )}
      </section>

      <section className="glass panel">
        <div className="panel-head"><Package size={16} /><span>Dome-level stock inventory</span></div>
        <div className="filter-toolbar">
          <div className="search-box">
            <Search size={14} />
            <input placeholder="Search dome or contractor…" value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <MultiSelectDropdown
            options={contractors}
            selected={contractorFilter}
            onChange={setContractorFilter}
            allLabel="All contractors"
            noneLabel="No contractors"
          />
          <select value={sourceFilter} onChange={(e) => setSourceFilter(e.target.value)}>
            <option value="all">All sources</option>
            <option value="inventory">Existing inventory</option>
            <option value="production">Production</option>
          </select>
        </div>
        <div className="table-meta">{filtered.length} domes · {fmt(totalStock)} WMT remaining (live — reflects finalized barges)</div>

        <div className="filtered-summary">
          <div className="filtered-summary-label">Summary of domes shown below ({filtered.length})</div>
          <div className="filtered-summary-grid">
            <div className="filtered-summary-item"><span>Total Stock</span><strong>{fmt(filteredTotals.stock)}</strong></div>
            <div className="filtered-summary-item"><span>Ni %</span><strong>{fmt(filteredTotals.ni, 2)}</strong></div>
            <div className="filtered-summary-item"><span>Fe %</span><strong>{fmt(filteredTotals.fe, 2)}</strong></div>
            <div className="filtered-summary-item"><span>Co %</span><strong>{fmt(filteredTotals.co, 3)}</strong></div>
            <div className="filtered-summary-item"><span>SiO2 %</span><strong>{fmt(filteredTotals.sio2, 2)}</strong></div>
            <div className="filtered-summary-item"><span>MgO %</span><strong>{fmt(filteredTotals.mgo, 2)}</strong></div>
            <div className="filtered-summary-item"><span>Al2O3 %</span><strong>{fmt(filteredTotals.al2o3, 2)}</strong></div>
            <div className="filtered-summary-item"><span>Si:Mg</span><strong>{fmt(filteredTotals.simg, 2)}</strong></div>
          </div>
          <div className="note" style={{ margin: "8px 0 0" }}>Weighted averages, each field excluding its own unassayed (0%) domes from the calculation — same as the list below, just visible without scrolling.</div>
        </div>

        <div className="table-wrap table-wrap-tall">
          <table className="data-table">
            <thead>
              <tr>
                <Th k="id">Dome ID</Th><Th k="contractor">Contractor</Th>
                <Th k="initialStock">Initial Stock</Th><Th k="stockOut">Stock Out (Barged)</Th><Th k="stock">Current Stock (WMT)</Th>
                <Th k="ni">Ni %</Th><Th k="fe">Fe %</Th><Th k="co">CO %</Th><Th k="sio2">SiO2 %</Th>
                <Th k="mgo">MgO %</Th><Th k="al2o3">Al2O3 %</Th><Th k="simg">Si/Mg</Th><Th k="location">Location</Th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((d) => (
                <tr key={d.id}>
                  <td>{d.id}</td>
                  <td><span className="dot" style={{ background: colorFor(d.contractor) }} />{d.contractor}</td>
                  <td>{fmt(d.initialStock)}</td>
                  <td className={d.stockOut > 0 ? "cell-stockout" : ""}>{d.stockOut > 0 ? fmt(d.stockOut) : "—"}</td>
                  <td>{fmt(d.stock)}</td><td>{fmt(d.ni, 2)}</td><td>{fmt(d.fe, 2)}</td><td>{fmt(d.co, 3)}</td>
                  <td>{fmt(d.sio2, 2)}</td><td>{fmt(d.mgo, 2)}</td><td>{fmt(d.al2o3, 2)}</td><td>{fmt(d.simg, 2)}</td>
                  <td>{d.location || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

/* ----------------------------- Barging Plan tab ----------------------------- */

function BargeRow({ barge, domesById, pool, onUpdate, onFinalize, onImport, onOpenInvoice, onExportBarge, onCheckStatus, onDataCommitted, isFeatureEnabled }) {
  const [open, setOpen] = useState(false);
  const [addDome, setAddDome] = useState("");
  const [addAmt, setAddAmt] = useState("");

  const availableDomes = pool.filter((p) => p.remaining > 1).sort((a, b) => b.remaining - a.remaining);

  // Weighted-average Si/Mg for this barge, computed live from its current sources —
  // works for generated, manually-built, and imported barges alike (not just barges
  // created via the Si/Mg-targeted generator). Only shown when at least one source
  // dome actually has Si/Mg data recorded.
  const avgSimg = useMemo(() => {
    let weight = 0, sum = 0;
    barge.sources.forEach((s) => {
      const d = domesById[s.id];
      const simg = d ? d.simg : undefined;
      if (simg !== undefined && simg > 0) { weight += s.amt; sum += s.amt * simg; }
    });
    return weight > 0 ? sum / weight : null;
  }, [barge.sources, domesById]);

  const updateSourceAmt = (domeId, amt) => {
    onUpdate(barge.no, barge.sources.map((s) => s.id === domeId ? { ...s, amt: Math.max(0, amt) } : s));
  };
  const removeSource = (domeId) => onUpdate(barge.no, barge.sources.filter((s) => s.id !== domeId));
  const addSourceRow = () => {
    if (!addDome || !addAmt) return;
    const dome = domesById[addDome];
    onUpdate(barge.no, [...barge.sources, { id: addDome, amt: parseFloat(addAmt), grade: dome.ni }]);
    setAddDome(""); setAddAmt("");
  };
  const setShipDate = (v) => onUpdate(barge.no, null, { shipDate: v });
  const setBargeName = (v) => onUpdate(barge.no, null, { bargeName: v });
  const setTugboatName = (v) => onUpdate(barge.no, null, { tugboatName: v });

  const handleFileImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.name.toLowerCase().endsWith(".csv")) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        // This template is a different export path than the Stock tab's CSVs (it can
        // come from Excel's own "Download as CSV", which doesn't reliably use the same
        // semicolon convention) — try semicolon first since that matches the rest of
        // this project's files, then fall back to comma if that finds no usable header.
        let rows = Papa.parse(evt.target.result, { skipEmptyLines: true, delimiter: ";" }).data;
        let comp = parseBargeComposition(rows);
        if (!comp.sources.length) {
          rows = Papa.parse(evt.target.result, { skipEmptyLines: true, delimiter: "," }).data;
          comp = parseBargeComposition(rows);
        }
        applyImportedComposition(comp);
      };
      reader.readAsText(file);
    } else {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const wb = XLSX.read(new Uint8Array(evt.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const comp = parseBargeComposition(rows);
        applyImportedComposition(comp);
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = "";
  };
  const applyImportedComposition = (comp) => {
    if (!comp.sources.length) {
      alert("⚠️ No usable rows found in this file. Check that it has a \"Dome ID\" / \"WMT\" header row somewhere in the first 15 rows, and that Dome IDs match what's on the Stock tab.");
      return;
    }
    const validSources = [];
    const unknownIds = [];
    comp.sources.forEach((c) => {
      const d = domesById[c.domeId];
      if (d) validSources.push({ id: c.domeId, amt: c.amt, grade: d.ni });
      else unknownIds.push(c.domeId);
    });
    if (unknownIds.length) {
      alert(
        `⚠️ Check your file — ${unknownIds.length} Dome ID${unknownIds.length > 1 ? "s were" : " was"} not found in stock and ${unknownIds.length > 1 ? "were" : "was"} skipped:\n\n` +
        unknownIds.slice(0, 20).join(", ") + (unknownIds.length > 20 ? `, +${unknownIds.length - 20} more` : "") +
        `\n\nDouble-check the Dome ID spelling against the Stock tab. ${validSources.length} valid row${validSources.length === 1 ? "" : "s"} were still applied.`
      );
    }
    if (validSources.length) {
      // Apply the metadata (Date/Barge Name/Tugboat Name) alongside the sources in one
      // combined update, rather than the sources plus separate metadata patches — avoids
      // three back-to-back updates each re-triggering their own Sheets write.
      const patch = {};
      if (comp.date) patch.shipDate = comp.date;
      if (comp.bargeName) patch.bargeName = comp.bargeName;
      if (comp.tugboatName) patch.tugboatName = comp.tugboatName;
      const freshBarges = onImport(barge.no, validSources, Object.keys(patch).length ? patch : undefined);
      onDataCommitted?.(freshBarges);
    }
  };

  return (
    <div className={`barge-row glass ${barge.finalized ? "barge-final" : ""}`}>
      <div className="barge-row-top" onClick={() => setOpen(!open)}>
        <div className="barge-row-left">
          <ChevronDown size={16} className={`chevron ${open ? "chevron-open" : ""}`} />
          <span className="barge-no">Barge {String(barge.no).padStart(2, "0")}</span>
          {barge.finalized ? <Lock size={12} className="lock-icon" /> : null}
        </div>
        <input type="date" value={barge.shipDate} onClick={(e) => e.stopPropagation()}
          onChange={(e) => setShipDate(e.target.value)} className="date-input" disabled={barge.finalized} />
        <span className="barge-total-wmt">{fmt(barge.totalWMT)} WMT</span>
        <span className="barge-avg-ni">{barge.totalWMT > 0 ? `${fmt(barge.grade, 2)}%` : "—"}</span>
        {avgSimg !== null && <span className="barge-avg-simg" title="Weighted-average Si/Mg ratio">Si/Mg {fmt(avgSimg, 2)}</span>}
        <StatusBadge status={barge.status} />
        <button className="btn-export-mini" title="Export this barge"
          onClick={(e) => { e.stopPropagation(); onExportBarge(barge); }}>
          <FileDown size={13} /> Export
        </button>
        {barge.finalized && (
          <button className="btn-invoice-mini" title="Generate Invoice"
            onClick={(e) => { e.stopPropagation(); onOpenInvoice(barge); }}>
            <FileText size={13} /> Invoice
          </button>
        )}
      </div>

      {open && (
        <div className="barge-row-detail">
          <div className="barge-info-row">
            <div className="info-field">
              <label>BG (Barge Name)</label>
              <input type="text" className="info-input" placeholder="e.g., BG-001" value={barge.bargeName || ""} disabled={barge.finalized}
                onChange={(e) => setBargeName(e.target.value)} />
            </div>
            <div className="info-field">
              <label>TB (Tugboat Name)</label>
              <input type="text" className="info-input" placeholder="e.g., TB-POWER-01" value={barge.tugboatName || ""} disabled={barge.finalized}
                onChange={(e) => setTugboatName(e.target.value)} />
            </div>
          </div>

          <div className="contractor-rollup">
            {Object.entries(
              barge.sources.reduce((acc, s) => {
                const c = (domesById[s.id] && domesById[s.id].contractor) || "—";
                acc[c] = (acc[c] || 0) + s.amt;
                return acc;
              }, {})
            ).sort((a, b) => b[1] - a[1]).map(([cid, amt]) => (
              <span key={cid} className="rollup-chip">
                <span className="dot" style={{ background: colorFor(cid) }} />{cid} {fmt((amt / (barge.totalWMT || 1)) * 100, 0)}%
              </span>
            ))}
            {barge.sources.length === 0 && <span className="empty-note">No dome data yet — add manually or import a plan below.</span>}
          </div>

          <table className="data-table data-table-compact">
            <thead><tr><th>Dome ID</th><th>Contractor</th><th>WMT</th><th>Ni %</th><th></th></tr></thead>
            <tbody>
              {barge.sources.map((s) => {
                const d = domesById[s.id];
                return (
                  <tr key={s.id}>
                    <td><span className="dot" style={{ background: colorFor(s.id) }} />{s.id}</td>
                    <td>{d ? d.contractor : "—"}</td>
                    <td>
                      <input type="number" className="cell-input" value={Math.round(s.amt)} disabled={barge.finalized}
                        onChange={(e) => updateSourceAmt(s.id, parseFloat(e.target.value) || 0)} />
                    </td>
                    <td>{fmt(s.grade, 2)}</td>
                    <td>{!barge.finalized && <button className="icon-btn" onClick={() => removeSource(s.id)}><Trash2 size={13} /></button>}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {!barge.finalized && (
            <>
              <div className="add-dome-row">
                <SearchableDomeSelect options={availableDomes} value={addDome} onChange={setAddDome} />
                <input type="number" placeholder="WMT" value={addAmt} onChange={(e) => setAddAmt(e.target.value)} />
                <button className="btn-ghost" onClick={addSourceRow}><Plus size={13} /> Add</button>
              </div>
              {isFeatureEnabled("Barging_CreateExcel", true) && (
                <label className="import-barge-label">
                  <FileUp size={13} /> Import final plan for this barge (.xlsx / .csv)
                  <input type="file" accept=".xlsx,.csv" onChange={handleFileImport} />
                </label>
              )}
            </>
          )}

          <div className="barge-row-actions">
            {isFeatureEnabled("Barging_FinalizeBarges", true) && (
              <button className={`btn-toggle ${barge.finalized ? "btn-toggle-on" : ""}`} onClick={() => onFinalize(barge.no)}>
                {barge.finalized ? <><Unlock size={13} /> Reopen</> : <><Lock size={13} /> Finalize</>}
              </button>
            )}
            <button className="btn-status" onClick={() => onCheckStatus(barge)}>
              <AlertTriangle size={13} /> Check Status
            </button>
            <button className="btn-export" onClick={() => onExportBarge(barge)}>
              <FileDown size={13} /> Export
            </button>
            {barge.finalized && (
              <button className="btn-invoice" onClick={() => onOpenInvoice(barge)}>
                <FileText size={13} /> Generate Invoice
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ----------------------------- Timeline tab ----------------------------- */

function TimelineTab({ barges, settings, isDevAccount }) {
  const [hoveredMonth, setHoveredMonth] = useState(null);
  const monthCounts = useMemo(() => {
    const arr = MONTHS.map(() => ({ final: 0, draft: 0, finalWMT: 0, draftWMT: 0 }));
    barges.forEach((b) => {
      const m = new Date(b.shipDate).getMonth();
      if (m < 0 || m > 11) return;
      if (b.finalized) { arr[m].final += 1; arr[m].finalWMT += b.totalWMT || 0; }
      else { arr[m].draft += 1; arr[m].draftWMT += b.totalWMT || 0; }
    });
    return arr;
  }, [barges]);

  const maxCount = Math.max(...monthCounts.map((m) => m.final + m.draft), 1);
  const finalizedTotal = barges.filter((b) => b.finalized).length;
  const currentMonthIdx = TODAY.getMonth();
  const monthsLeft = Math.max(1, 12 - currentMonthIdx);
  const remaining = settings.planTarget - finalizedTotal;

  // Barges that have been through the Loading Report tracker at least once — a barge
  // with no report yet simply won't have these fields, so this naturally only includes
  // barges someone has actually reported progress for.
  const loadingBarges = useMemo(() => barges.filter((b) => b.progressPercent !== undefined).sort((a, b) => a.no - b.no), [barges]);
  const totalQtyOnBoard = loadingBarges.reduce((s, b) => s + (b.qtyOnBoard || 0), 0);
  const totalLoadingCapacity = loadingBarges.reduce((s, b) => s + (b.totalWMT || 0), 0);
  const overallLoadingPct = totalLoadingCapacity > 0 ? (totalQtyOnBoard / totalLoadingCapacity) * 100 : 0;

  return (
    <div className="stack">
      <section className="glass summary-strip">
        <Kpi label="Finalized" value={`${finalizedTotal}`} unit={`/ ${settings.planTarget}`} accent="good" />
        <Kpi label="Created" value={`${barges.length}`} unit={`/ ${settings.planTarget}`} />
        <Kpi label="Months left in 2026" value={`${monthsLeft}`} unit="mo" />
        <Kpi label="Required pace" value={fmt(remaining / monthsLeft, 1)} unit="barges/mo" />
      </section>

      {isDevAccount && loadingBarges.length > 0 && (
        <section className="glass panel">
          <div className="panel-head"><MessageSquare size={16} /><span>Loading Progress</span></div>
          <div className="loading-overall">
            <div className="loading-overall-numbers">
              <span className="loading-overall-value">{fmt(totalQtyOnBoard)}</span>
              <span className="loading-overall-sep">/</span>
              <span className="loading-overall-total">{fmt(totalLoadingCapacity)} WMT</span>
              <span className="loading-overall-pct">{fmt(overallLoadingPct, 1)}%</span>
            </div>
            <div className="loading-bar-track">
              <div className="loading-bar-fill" style={{ width: `${Math.min(100, overallLoadingPct)}%` }} />
            </div>
            <div className="note" style={{ marginTop: 8 }}>Across {loadingBarges.length} barge{loadingBarges.length !== 1 ? "s" : ""} with a loading report on file.</div>
          </div>

          <div className="loading-barge-list">
            {loadingBarges.map((b) => (
              <div key={b.no} className="loading-barge-row">
                <div className="loading-barge-head">
                  <span className="tracker-no">#{String(b.no).padStart(2, "0")}</span>
                  <span className={`loading-status-badge ${b.loadingStatus === "loaded" ? "loading-status-done" : "loading-status-active"}`}>
                    {b.loadingStatus === "loaded" ? "Loaded" : "Loading"}
                  </span>
                  <span className="loading-barge-updated">Updated {b.lastUpdated || "—"}</span>
                </div>
                <div className="loading-bar-track loading-bar-track-sm">
                  <div className="loading-bar-fill" style={{ width: `${Math.min(100, b.progressPercent || 0)}%` }} />
                </div>
                <div className="loading-barge-meta">
                  <span>{fmt(b.qtyOnBoard)} / {fmt(b.totalWMT)} WMT ({fmt(b.progressPercent, 0)}%)</span>
                  <span>Balance: {fmt(b.balanceDue)} WMT</span>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      <section className="glass panel">
        <div className="panel-head"><TrendingUp size={16} /><span>Shipping schedule by month</span></div>
        <div className="month-chart">
          {monthCounts.map((m, i) => (
            <div className="month-col" key={i} onMouseEnter={() => setHoveredMonth(i)} onMouseLeave={() => setHoveredMonth(null)}>
              {hoveredMonth === i && (
                <div className="chart-tooltip">
                  <div className="chart-tooltip-title">{MONTHS[i]}</div>
                  <div className="chart-tooltip-row"><span>Total</span><strong>{fmt(m.finalWMT + m.draftWMT)} WMT</strong></div>
                  <div className="chart-tooltip-row"><span className="chart-tooltip-dot" style={{ background: "#E35F0C" }} />Finalized<strong>{fmt(m.finalWMT)} WMT</strong></div>
                  <div className="chart-tooltip-row"><span className="chart-tooltip-dot" style={{ background: "#C9A227" }} />Draft<strong>{fmt(m.draftWMT)} WMT</strong></div>
                  <div className="chart-tooltip-row chart-tooltip-muted"><span>Barges</span><strong>{m.final + m.draft}</strong></div>
                </div>
              )}
              <div className="month-bars">
                <div className="month-bar month-bar-shipped" style={{ height: `${(m.final / maxCount) * 100}%` }} />
                <div className="month-bar month-bar-other" style={{ height: `${(m.draft / maxCount) * 100}%` }} />
              </div>
              <div className={`month-name ${i === currentMonthIdx ? "month-now" : ""}`}>{MONTHS[i]}</div>
              <div className="month-count">{m.final + m.draft}</div>
            </div>
          ))}
        </div>
        <div className="legend">
          <span className="legend-item"><span className="dot" style={{ background: "#E35F0C" }} />Finalized</span>
          <span className="legend-item"><span className="dot" style={{ background: "#C9A227" }} />Draft / planned</span>
        </div>
      </section>

      <section className="glass panel">
        <div className="panel-head"><Ship size={16} /><span>Barge tracker</span></div>
        {barges.length === 0 && <div className="empty-note">No barges created yet — head to Barging Plan to add or generate some.</div>}
        <div className="tracker-list">
          {barges.map((b) => (
            <div key={b.no} className="tracker-row">
              <span className="tracker-no">#{String(b.no).padStart(2, "0")}</span>
              <span className="tracker-month">{b.shipDate}</span>
              <span className="tracker-grade">{b.totalWMT > 0 ? `${fmt(b.grade, 2)}% Ni` : "—"}</span>
              <StatusBadge status={b.status} />
              {isDevAccount && b.progressPercent !== undefined && (
                <span className={`loading-status-badge ${b.loadingStatus === "loaded" ? "loading-status-done" : "loading-status-active"}`}>
                  {fmt(b.progressPercent, 0)}%
                </span>
              )}
              {b.finalized && <Lock size={12} className="lock-icon" />}
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

/* ============================================================
 * FinancialsTab — admin-only royalty/PNBP calculator.
 * Royalty (IDR) = HMA price (USD/WMT) x Exchange Rate (IDR/USD) x Barge Qty (WMT) x 15% tariff.
 * Each barge uses the HMA/rate that was in effect on ITS OWN shipDate, not today's —
 * that's the whole point of keeping history instead of a single current price.
 * NOTE: tariff is a flat 15% here per the build spec. Real PNBP/royalty tariffs can vary
 * by HMA price bracket under Indonesian mining regulation — flag this if precise bracket
 * rules are needed later; flat-rate is intentionally the simple first pass.
 * ============================================================ */
function FinancialsTab({ barges, hpmHistory, setHpmHistory, exchangeRateHistory, setExchangeRateHistory,
  getHpmOnDate, getExchangeRateOnDate, calculateRoyalty, getHpmTrendPercent, getExRateTrendPercent, exportFinancialData,
  onHpmUpdated, onExchangeRateUpdated }) {

  const updateHpm = () => {
    const newPrice = prompt("Enter new HPM price (USD/WMT):", hpmHistory[0]?.price);
    if (newPrice && !isNaN(parseFloat(newPrice))) {
      const today = new Date().toISOString().split("T")[0];
      const updated = [{ date: today, price: parseFloat(newPrice), unit: "USD/WMT" }, ...hpmHistory].slice(0, 180);
      setHpmHistory(updated);
      onHpmUpdated?.(updated);
    }
  };
  const updateRate = () => {
    const newRate = prompt("Enter new exchange rate (IDR/USD):", exchangeRateHistory[0]?.rate);
    if (newRate && !isNaN(parseInt(newRate))) {
      const today = new Date().toISOString().split("T")[0];
      const updated = [{ date: today, rate: parseInt(newRate), source: "Manual" }, ...exchangeRateHistory].slice(0, 180);
      setExchangeRateHistory(updated);
      onExchangeRateUpdated?.(updated);
    }
  };

  const hpmTrend = getHpmTrendPercent();
  const exTrend = getExRateTrendPercent();
  const sortedBarges = [...barges].sort((a, b) => a.no - b.no);
  const totalRoyalty = barges.reduce((sum, b) => sum + calculateRoyalty(b.shipDate, b.totalWMT || 0), 0);

  return (
    <div className="stack">
      <section className="glass panel">
        <div className="panel-head"><DollarSign size={16} /><span>Royalty &amp; PNBP Calculator</span></div>
        <p className="note" style={{ marginBottom: 16 }}>Royalty = HPM x Exchange Rate x Qty x 15% tariff. HPM is entered manually per period — not calculated from a formula.</p>

        <div className="financial-cards-row">
          <div className="financial-card">
            <div className="card-header">
              <span>HPM (Harga Patokan Mineral)</span>
              <span className="card-unit">USD/WMT</span>
            </div>
            <div className="card-main-value">${hpmHistory[0]?.price?.toFixed(2) || "—"}</div>
            <div className="card-meta">
              <div className="meta-item"><span className="meta-label">Last Updated</span><span className="meta-value">{hpmHistory[0]?.date || "—"}</span></div>
              <div className="meta-item">
                <span className="meta-label">Change vs. Previous</span>
                <span className={`meta-value trend ${hpmTrend >= 0 ? "up" : "down"}`}>{hpmTrend >= 0 ? "↑" : "↓"} {Math.abs(hpmTrend).toFixed(2)}%</span>
              </div>
            </div>
            <div className="card-actions"><button onClick={updateHpm} className="btn-card-action">Update HPM</button></div>
          </div>

          <div className="financial-card">
            <div className="card-header">
              <span>USD to IDR Exchange Rate</span>
              <span className="card-unit">Bank Indonesia / XE.com</span>
            </div>
            <div className="card-main-value">{exchangeRateHistory[0]?.rate?.toLocaleString("id-ID") || "—"} IDR</div>
            <div className="card-meta">
              <div className="meta-item"><span className="meta-label">Last Updated</span><span className="meta-value">{exchangeRateHistory[0]?.date || "—"}</span></div>
              <div className="meta-item">
                <span className="meta-label">Change vs. Previous</span>
                <span className={`meta-value trend ${exTrend >= 0 ? "up" : "down"}`}>{exTrend >= 0 ? "↑" : "↓"} {Math.abs(exTrend).toFixed(2)}%</span>
              </div>
            </div>
            <div className="card-actions"><button onClick={updateRate} className="btn-card-action">Update Rate</button></div>
          </div>
        </div>
      </section>

      <section className="glass panel">
        <div className="panel-head" style={{ justifyContent: "space-between" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 8 }}><DollarSign size={16} /> Royalty by Barge</span>
          <button onClick={exportFinancialData} className="btn-export" title="Export financial data">
            <FileUp size={14} /> Export Data
          </button>
        </div>
        <p className="note" style={{ marginTop: 0, marginBottom: 12 }}>
          Each barge uses the HPM and exchange rate that were in effect on its own ship date, not today's — that's
          the whole point of keeping history instead of a single current figure.
        </p>
        <div className="royalty-table">
          <div className="royalty-header">
            <div>Barge No.</div><div>Date</div><div>Qty (WMT)</div><div>HPM ($/WMT)</div><div>ExRate (IDR)</div><div>Royalty (IDR)</div><div>Status</div>
          </div>
          {sortedBarges.length === 0 && <div className="log-empty">No barges yet.</div>}
          {sortedBarges.map((barge) => {
            const royalty = calculateRoyalty(barge.shipDate, barge.totalWMT || 0);
            const hpm = getHpmOnDate(barge.shipDate);
            return (
              <div key={barge.no} className="royalty-row">
                <div className="cell">#{barge.no}</div>
                <div className="cell">{barge.shipDate || "—"}</div>
                <div className="cell">{fmt(barge.totalWMT)}</div>
                <div className="cell">${hpm.toFixed(2)}</div>
                <div className="cell">{fmt(getExchangeRateOnDate(barge.shipDate))}</div>
                <div className="cell highlight">{royalty.toLocaleString("id-ID", { maximumFractionDigits: 0 })}</div>
                <div className="cell status">{barge.finalized ? "Finalized" : "Draft"}</div>
              </div>
            );
          })}
        </div>
        <div className="royalty-summary">
          <strong>Total Royalty (All Barges):</strong>
          <span className="total">{totalRoyalty.toLocaleString("id-ID", { maximumFractionDigits: 0 })} IDR</span>
        </div>
      </section>
    </div>
  );
}

/* ============================================================
 * LoginLogTab — admin-only view of all login attempts.
 * ============================================================ */
function LoginLogTab({ loginHistory }) {
  const [filterUser, setFilterUser] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [sortOrder, setSortOrder] = useState("desc");

  const successCount = loginHistory.filter((l) => l.status === "Successful").length;
  const failedCount = loginHistory.filter((l) => l.status === "Failed").length;

  const filtered = useMemo(() => {
    let list = loginHistory;
    if (filterUser.trim()) list = list.filter((l) => l.username?.toLowerCase().includes(filterUser.trim().toLowerCase()));
    if (filterStatus) list = list.filter((l) => l.status === filterStatus);
    return [...list].sort((a, b) => {
      const diff = new Date(a.timestamp) - new Date(b.timestamp);
      return sortOrder === "desc" ? -diff : diff;
    });
  }, [loginHistory, filterUser, filterStatus, sortOrder]);

  return (
    <div className="stack">
      <section className="glass panel">
        <div className="panel-head"><History size={16} /><span>Login History</span></div>
        <p className="note" style={{ marginBottom: 16 }}>All login attempts (successful &amp; failed) are tracked here.</p>

        <div className="review-grid" style={{ marginBottom: 16 }}>
          <div className="form-group"><label>Filter by username</label>
            <input className="login-input" placeholder="e.g. operation" value={filterUser} onChange={(e) => setFilterUser(e.target.value)} />
          </div>
          <div className="form-group"><label>Filter by status</label>
            <select className="login-input" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
              <option value="">All statuses</option>
              <option value="Successful">Successful</option>
              <option value="Failed">Failed</option>
            </select>
          </div>
          <div className="form-group"><label>Sort</label>
            <select className="login-input" value={sortOrder} onChange={(e) => setSortOrder(e.target.value)}>
              <option value="desc">Newest first</option>
              <option value="asc">Oldest first</option>
            </select>
          </div>
        </div>

        <div className="log-table">
          <div className="log-header-row">
            <div className="log-cell log-cell-time">Timestamp</div>
            <div className="log-cell log-cell-user">Username</div>
            <div className="log-cell log-cell-status">Status</div>
          </div>
          {filtered.length === 0 && <div className="log-empty">{loginHistory.length === 0 ? "No login history yet." : "No entries match this filter."}</div>}
          {filtered.map((entry, idx) => {
            const timeStr = new Date(entry.timestamp).toLocaleString("en-US", {
              year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit",
            });
            return (
              <div key={idx} className={`log-row ${entry.status === "Successful" ? "log-success" : "log-failed"}`}>
                <div className="log-cell log-cell-time">{timeStr}</div>
                <div className="log-cell log-cell-user">{entry.username}</div>
                <div className="log-cell log-cell-status">
                  <span className={`status-badge ${entry.status === "Successful" ? "badge-success" : "badge-failed"}`}>{entry.status}</span>
                </div>
              </div>
            );
          })}
        </div>

        <p className="note" style={{ textAlign: "right", marginTop: 8 }}>Showing {filtered.length} of {loginHistory.length} entries</p>

        <div className="mini-stats-row">
          <div className="mini-stat-card">
            <div className="mini-stat-label">Total Login Attempts</div>
            <div className="mini-stat-value">{loginHistory.length}</div>
          </div>
          <div className="mini-stat-card">
            <div className="mini-stat-label">Successful Logins</div>
            <div className="mini-stat-value mini-stat-success">{successCount}</div>
          </div>
          <div className="mini-stat-card">
            <div className="mini-stat-label">Failed Attempts</div>
            <div className="mini-stat-value mini-stat-failed">{failedCount}</div>
          </div>
        </div>
      </section>
    </div>
  );
}

/* ============================================================
 * SettingsTab — Data Export (admin only) and Account (everyone).
 * ============================================================ */
/* ============================================================
 * AccountManagement — dev-only. Create/disable/delete users.
 * ============================================================ */
function AccountManagement({ allUsers, onCreateUser, onDisableUser, onDeleteUser }) {
  const [showForm, setShowForm] = useState(false);
  const [newUser, setNewUser] = useState({ username: "", password: "", role: "operation" });

  const handleCreate = async () => {
    if (!newUser.username || !newUser.password) { alert("⚠️ Username and password required"); return; }
    if (allUsers.some((u) => u["Username"] === newUser.username)) { alert("❌ Username already exists"); return; }
    const today = new Date().toISOString().split("T")[0];
    const success = await onCreateUser({
      Username: newUser.username, Password: newUser.password, Role: newUser.role,
      Status: "active", Created_Date: today, Last_Modified: today,
    });
    if (success) { alert(`✅ User ${newUser.username} created`); setNewUser({ username: "", password: "", role: "operation" }); setShowForm(false); }
  };

  return (
    <div className="settings-card" style={{ flexDirection: "column", alignItems: "stretch", gap: 14 }}>
      <div className="card-content" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <div style={{ minWidth: 0, flex: "1 1 200px" }}><h4>Account Management</h4><p>Create, disable, or delete user accounts.</p></div>
        <button onClick={() => setShowForm(!showForm)} className="btn-settings-action btn-sync-sheets" style={{ flexShrink: 0 }}>+ Add User</button>
      </div>

      {showForm && (
        <>
          <div className="review-grid">
            <div className="form-group"><label>Username</label>
              <input className="login-input" value={newUser.username} onChange={(e) => setNewUser((p) => ({ ...p, username: e.target.value }))} /></div>
            <div className="form-group"><label>Password</label>
              <input type="password" className="login-input" value={newUser.password} onChange={(e) => setNewUser((p) => ({ ...p, password: e.target.value }))} /></div>
            <div className="form-group"><label>Role</label>
              <select className="login-input" value={newUser.role} onChange={(e) => setNewUser((p) => ({ ...p, role: e.target.value }))}>
                <option value="operation">Operation</option>
                <option value="admin">Admin</option>
              </select></div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={handleCreate} className="btn-settings-action btn-sync-sheets" style={{ flex: "1 1 100px" }}>Create</button>
            <button onClick={() => setShowForm(false)} className="btn-ghost" style={{ flex: "1 1 100px" }}>Cancel</button>
          </div>
        </>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {allUsers.length === 0 && <p className="note" style={{ margin: 0 }}>No users loaded yet.</p>}
        {allUsers.map((u) => (
          <div key={u["Username"]} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: 12, background: "rgba(0,0,0,.2)", border: "1px solid rgba(255,255,255,.1)", borderRadius: 8, opacity: u["Status"] === "disabled" ? 0.6 : 1, flexWrap: "wrap", gap: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <strong style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12 }}>{u["Username"]}</strong>
              <span className="loading-status-badge loading-status-active">{u["Role"]}</span>
              <span style={{ fontSize: 10, fontWeight: 600, color: u["Status"] === "active" ? "#4ADE80" : "#F87171" }}>
                {u["Status"] === "active" ? "🟢 Active" : "🔴 Disabled"}
              </span>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button onClick={() => onDisableUser(u["Username"], u["Status"] === "active" ? "disabled" : "active")} className="btn-settings-action btn-sync-sheets" style={{ padding: "6px 10px", fontSize: 10 }}>
                {u["Status"] === "active" ? "⛔ Disable" : "✅ Enable"}
              </button>
              <button onClick={() => { if (confirm(`⚠️ Permanently delete user "${u["Username"]}"? This cannot be undone.`)) onDeleteUser(u["Username"]); }}
                className="btn-settings-action btn-logout-action" style={{ padding: "6px 10px", fontSize: 10 }}>🗑️ Delete</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================
 * SubFeatureConfiguration — granular per-user, per-tab feature toggles.
 * Select a user, select a tab, see every sub-feature in that tab as a
 * checkbox. Select All / Deselect All shortcuts. Saves to both the
 * detailed sheet and the tab-level sheet (derived from it) at once.
 * ============================================================ */
function SubFeatureConfiguration({ allUsers, allDetailedFlags, onSave }) {
  const [selectedUser, setSelectedUser] = useState("");
  const [selectedTab, setSelectedTab] = useState("");
  const [localFlags, setLocalFlags] = useState(allDetailedFlags);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => { setLocalFlags(allDetailedFlags); }, [allDetailedFlags]);

  const tabConfig = selectedTab ? FEATURE_STRUCTURE[selectedTab] : null;
  const userFlags = selectedUser ? (localFlags[selectedUser] || {}) : {};
  const isOn = (key) => userFlags[key] === "true" || userFlags[key] === true;
  const allEnabled = tabConfig ? tabConfig.features.every((f) => isOn(f.key)) : false;
  const allDisabled = tabConfig ? tabConfig.features.every((f) => !isOn(f.key)) : false;

  const setFlags = (updater) => setLocalFlags((prev) => ({ ...prev, [selectedUser]: updater(prev[selectedUser] || {}) }));
  const handleToggle = (key) => setFlags((flags) => ({ ...flags, [key]: !isOn(key) }));
  const handleSelectAll = () => setFlags((flags) => ({ ...flags, ...Object.fromEntries(tabConfig.features.map((f) => [f.key, true])) }));
  const handleDeselectAll = () => setFlags((flags) => ({ ...flags, ...Object.fromEntries(tabConfig.features.map((f) => [f.key, false])) }));

  const handleSave = async () => {
    setIsSaving(true);
    const success = await onSave(localFlags);
    setIsSaving(false);
    if (success) alert(`✅ Saved for ${selectedUser} — ${tabConfig.label}`);
    else alert("❌ Save failed — check the sync error detail above.");
  };

  return (
    <div className="settings-card" style={{ flexDirection: "column", alignItems: "stretch", gap: 14 }}>
      <div className="card-content"><h4>Feature Configuration</h4><p>Select a user and tab to fine-tune individual sub-features.</p></div>

      <div className="review-grid">
        <div className="form-group"><label>User</label>
          <select className="login-input" value={selectedUser} onChange={(e) => { setSelectedUser(e.target.value); setSelectedTab(""); }}>
            <option value="">Choose user...</option>
            {allUsers.map((u) => <option key={u["Username"]} value={u["Username"]}>{u["Username"]}</option>)}
          </select>
        </div>
        {selectedUser && (
          <div className="form-group"><label>Tab</label>
            <select className="login-input" value={selectedTab} onChange={(e) => setSelectedTab(e.target.value)}>
              <option value="">Choose tab...</option>
              {Object.entries(FEATURE_STRUCTURE).map(([key, cfg]) => <option key={key} value={key}>{cfg.label}</option>)}
            </select>
          </div>
        )}
      </div>

      {tabConfig && (
        <>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={handleSelectAll} disabled={allEnabled || isSaving} className="btn-settings-action btn-sync-sheets" style={{ flex: "1 1 100px" }}>Select All</button>
            <button onClick={handleDeselectAll} disabled={allDisabled || isSaving} className="btn-settings-action btn-logout-action" style={{ flex: "1 1 100px" }}>Deselect All</button>
          </div>

          <div className="lab-list" style={{ background: "rgba(0,0,0,.2)" }}>
            {tabConfig.features.map((f) => (
              <label key={f.key} className="lab-item" style={{ flexDirection: "row", alignItems: "center", gap: 10, cursor: "pointer" }}>
                <input type="checkbox" checked={isOn(f.key)} onChange={() => handleToggle(f.key)} disabled={isSaving}
                  style={{ width: 15, height: 15, accentColor: "#22D3B8", flexShrink: 0 }} />
                <span className="lab-detail" style={{ color: "#EAF0F6" }}>
                  {f.label}{!WIRED_SUBFEATURES.has(f.key) && <span style={{ color: "#8A97A8", fontSize: "10px" }}> (not yet wired to UI)</span>}
                </span>
              </label>
            ))}
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
            <button onClick={() => setSelectedTab("")} className="btn-ghost" disabled={isSaving} style={{ flex: "1 1 100px", justifyContent: "center" }}>Back</button>
            <button onClick={handleSave} className="btn-settings-action btn-sync-sheets" disabled={isSaving} style={{ flex: "1 1 100px", justifyContent: "center" }}>
              {isSaving ? "Saving…" : "Save Configuration"}
            </button>
          </div>
        </>
      )}
    </div>
  );
}

function SettingsTab({ isAdmin, currentUser, handleLogout, exportAllForGoogleSheets, syncWithSheets, sheetsSyncStatus, lastSyncTime, lastSyncError, exRateFetchStatus, allUsers, allFeatureFlags, setAllUsers, setAllFeatureFlags, writeUsersToSheets, writeFeatureFlagsToSheets, allDetailedFlags, setAllDetailedFlags, writeDetailedFeatureFlagsToSheets, isFeatureEnabled }) {
  return (
    <div className="stack">
      <section className="glass panel" style={{ padding: 0 }}>
        <div className="settings-header">
          <h2>Settings &amp; Admin</h2>
          <p style={{ fontSize: "12px", color: "#8A97A8", marginTop: "4px" }}>Export data, manage account settings</p>
        </div>

        {isAdmin && (
          <div className="settings-section">
            <div className="settings-section-header">
              <h3>Live Sheets Sync</h3>
              <span className="section-badge">Admin Only</span>
            </div>
            <div className="settings-card">
              <div className="card-content">
                <h4>Status: {sheetsSyncStatus}</h4>
                <p>
                  {lastSyncTime ? `Last synced ${lastSyncTime.toLocaleString()}` : "Not synced yet this session"} — pulls Domes, Barges,
                  HPM, and Exchange Rates from your Google Sheet automatically for every logged-in user. Runs on login, every 5 minutes,
                  and immediately whenever you switch back to this tab — no manual step needed. Finalizing/reopening a barge and Excel
                  imports also push changes back to Sheets automatically.
                </p>
              </div>
            </div>
            {lastSyncError && (
              <div className="sync-error-detail">
                <AlertTriangle size={13} /> {lastSyncError}
              </div>
            )}
          </div>
        )}

        {isAdmin && (
          <div className="settings-section">
            <div className="settings-section-header">
              <h3>Exchange Rate Auto-Fetch</h3>
              <span className="section-badge">Admin Only</span>
            </div>
            <div className="settings-card">
              <div className="card-content">
                <h4>{exRateFetchStatus}</h4>
                <p>Checks once/day (recheck runs hourly, only fetches if today's date hasn't been fetched yet) via a free public API. Manual "Update Rate" on the Financials tab always works regardless of this.</p>
              </div>
            </div>
          </div>
        )}

        {currentUser?.username === "dev" && (
          <div className="settings-section dev-admin-section">
            <div className="settings-section-header">
              <h3>Developer Configuration</h3>
              <span className="section-badge dev-badge">Dev Only</span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <AccountManagement
                allUsers={allUsers}
                onCreateUser={async (user) => {
                  const updated = [...allUsers, user];
                  setAllUsers(updated);
                  return await writeUsersToSheets(updated);
                }}
                onDisableUser={async (username, status) => {
                  const updated = allUsers.map((u) => u["Username"] === username ? { ...u, Status: status, Last_Modified: new Date().toISOString().split("T")[0] } : u);
                  setAllUsers(updated);
                  const ok = await writeUsersToSheets(updated);
                  alert(ok ? `✅ User ${username} ${status}` : "❌ Failed to update user");
                }}
                onDeleteUser={async (username) => {
                  const updatedUsers = allUsers.filter((u) => u["Username"] !== username);
                  setAllUsers(updatedUsers);
                  const updatedFlags = { ...allFeatureFlags };
                  delete updatedFlags[username];
                  setAllFeatureFlags(updatedFlags);
                  const updatedDetailed = { ...allDetailedFlags };
                  delete updatedDetailed[username];
                  setAllDetailedFlags(updatedDetailed);
                  const usersOk = await writeUsersToSheets(updatedUsers);
                  const flagsOk = await writeFeatureFlagsToSheets(updatedFlags);
                  const detailedOk = await writeDetailedFeatureFlagsToSheets(updatedDetailed);
                  alert(usersOk && flagsOk && detailedOk ? `✅ User ${username} deleted` : "❌ Failed to fully delete user");
                }}
              />
              <SubFeatureConfiguration allUsers={allUsers} allDetailedFlags={allDetailedFlags} onSave={writeDetailedFeatureFlagsToSheets} />
            </div>
          </div>
        )}

        {isAdmin && isFeatureEnabled("Settings_ExportToSheets", true) && (
          <div className="settings-section">
            <div className="settings-section-header">
              <h3>Data Export</h3>
              <span className="section-badge">Admin Only</span>
            </div>
            <div className="settings-card">
              <div className="card-content">
                <h4>Export to Google Sheets</h4>
                <p>Downloads Domes, Barges, HPM history, Exchange rates, and Login history as 5 CSV files — import each into the matching tab of your Google Sheet.</p>
              </div>
              <button onClick={exportAllForGoogleSheets} className="btn-settings-action btn-export">
                <FileUp size={16} /> <span>Export Data</span>
              </button>
            </div>
          </div>
        )}

        <div className="settings-section">
          <div className="settings-section-header"><h3>Account</h3></div>
          <div className="settings-card">
            <div className="card-content">
              <h4>Logged in as: <strong>{currentUser?.username}</strong></h4>
              <p>Role: {currentUser?.role === "admin" ? "Administrator" : "Operations"}</p>
            </div>
            {isFeatureEnabled("Settings_Logout", true) && (
              <button onClick={handleLogout} className="btn-settings-action btn-logout-action">
                <LogOut size={16} /> <span>Logout</span>
              </button>
            )}
          </div>
        </div>
      </section>
    </div>
  );
}

/* ----------------------------- main app ----------------------------- */


/* ============================================================
 * Invoice generator — data URIs for letterhead / signature / stamp
 * ============================================================ */
const INVOICE_LOGO_HEADER = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBAUEBAYFBQUGBgYHCQ4JCQgICRINDQoOFRIWFhUSFBQXGiEcFxgfGRQUHScdHyIjJSUlFhwpLCgkKyEkJST/2wBDAQYGBgkICREJCREkGBQYJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCT/wAARCAF/AfQDASIAAhEBAxEB/8QAHQABAAMAAwEBAQAAAAAAAAAAAAYHCAEEBQMCCf/EAFwQAAEDAwEDBgcIDAkMAQQDAAEAAgMEBREGBxIhCBMxQVGTFRdUVWGR0RQWIjZxdIGyMjVCUnOUsbPBwtLhNDdEZHJ1g5KhIyQlJiczRVNigqLDZRhDY7Sk8PH/xAAbAQEAAgMBAQAAAAAAAAAAAAAABQYBBAcCA//EADYRAQABAQMHCwMEAwEAAAAAAAABAgMEEQUWNFFSkdEGEhMUMlNxgaGxwRUh4TEzQWEkgvDx/9oADAMBAAIRAxEAPwDVCIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiFAREQEREBEQICIiAiIgIiICIiAiIgIiICIiAiIgIidKAiIgIiICIiAiIgIiICIiCP631LNpSyG4w07Kh3Osj3HuLRh2eOR8igB233DpFnpO+d7FJtsXxPPzqL9Ko0lVPLWULxYXjmWVWEYRq/tcchZMu14u3PtaMZxnX/Sy/HhcfM9J3zvYnjwuHmek753sVZ9C5UT9YvneekcEz9DuPd+s8Vl+PC4eZ6TvnexPHjcPM9J3zvYqzRPrF82/SOB9DuPd+s8VmePC4eZ6TvnexPHhcPM9J3zvYqzXKfWb53npHA+h3Hu/WeK2bHtfrrteaGgfaqaNtTM2IvbK4loJxnGFaQWbdHH/AFss3zyL6y0krNkK9Wt4s6qrWrGYlVOUFzsbta0U2NOETHyIiKcQAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiIINti+J5+dRfpVHcFpm+2Ch1HQ+4bjG58G+1+GvLTkdHEKPeKPSg/klR+Mv9qreVckW16t+ks5jDDD7/8Ai05Hy1YXO79FaROOMz9vL+1DYTCvnxSaU8kqPxl/tTxSaU8kqPxl/tUbm7etcb54JXOe6aqt0cVCkIAr58UelPJKj8Zf7U8UelPJKj8Zf7UzdvWuN88DOe6aqt0cVD4RXydkelPI6j8Zf7U8UelPJKj8Zf7VjN29a43zwM57pqq3RxU/o4gass3zyL6y0kopQ7MNNW6tgraelnbNA8SMJneQHDo4ZUrHQrBke4Wl0s6qbTD7z/Ct5byjZX20prsonCIw+4iIpdCiIiB1IiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIUBERAREQEREBERAREQEREBERARRTaTe6/T+nPdtumEM/Pxs3i0O4HOeBVWeNTVvnNn4uz2KKvmV7G62nR2kTj+v2/9S9xyLb3yz6WzmMMcPvjwX8ioDxqau85s/F4/YnjU1d5zZ+Ls9i1c47tqndHFuZsXvap3zwX+ioDxqat85s/F2exDtU1aP8AibPxeP2JnHdtU7o4mbF62qd88F/os/8AjV1b5zZ+Ls9i5G1TVvnNv4uz2JnHdtVW6OJmxetqnfPBf6Kk9N7RtTXHUFtpKm4NfBPUsje3mWDLSePEBXYpK43+zvdM1WcT9taLv+T7S5VRRazH3+/2ERFutAREQEREBERAREQAmURAREQEREBERAREQEREBERAREQEREBERAREQEROlAREQEREBERAREQEREBERAREQEREBERAREQQbbF8Tz86i/KVRpV5bYSPeeeI/hUX6VRxx2j1qj8odL8o+V/5M6H5z8PyuQuOHaPWnDtHrUGsJ1rlBjtHrXPDtHrQfnCBcnB6x60GO0etYHs6P+Nlm+eRfWC0ks26PI99lm4j+GRdf/UtJK48mv2a/H4UjlV+9R4fIiIrIqwiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAidKICIiAiIgIiICJ8iICIiAiIgJxREBERA6EREBERAREQEREBERB8aqipq6LmqunhqI8h25KwOGeo4K6fvbsh6bPbvxZnsXpIvFVnRVOMw902ldMYUzMPO97dk8z278WZ7Fx727J5nt34sz2L0kWOhs9mNz101ptTved727J5nt34sz2J727If+D278WZ7F6KJ0NnsxuY6a02p3vN97dk8z278WZ7E97dk8z278WZ7F6SJ0NnsxuOmtNqd7oRaftEEjJYrVQRyMO817YGAtPaDhd/oRF6popp7MYPNVdVXanERE+henkREQEREBERAREQETiiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiBAQoiAiIgIiICdSIgIiICIiAiIgIiICIiAiIgjWv55aewF8MskbueYN5ji09fWFW/hOvP8uqu+d7VYu0X4u/28f6VWGVROUdpVTe8In+I+W/dojmO14Trh0V1V3zvauPCdef5dVd872rrIoDpa9qd7Y5sOyLnXD+XVXfO9qeE6/y6q753tXXTCz0te1O8wh2Bc6/y6q753tXPhOv8uqu+d7V1iuE6WvaneYQ7XhOv8tqu+d7U8J1/ltV3zvautlE6WvaneYQ7XhOu8uqu+d7Vx4TrvLqrvne1ddcJ01e1O8wh2fCdf5dVd872p4Sr/Larvne1dZMp01e1O8wh2vCdf5bVd872rjwnX+XVXfO9q664Tpa9qd5hDteE6/y6q753tTwnXeXVXfO9q6yJ0te1O8wh2fCdf5dVd872rjwnX5/htV3zvausidLXtTvMIdrwnXeXVXfO9q/PhSv8uqu+d7V10Tpa9qd5hDs+E67y6q753tTwnXeXVXfO9q6yBOlr1zvMIdjwlX+XVXfO9q/XhOv8uqu+d7V1elcp0te1O8wh2fCdd5bVd872p4Tr/Lqrvne1dZcEp0te1O8wh2vClf0e7ar5eed7VPdnFRNPb6syzyzETAAyPLiPgjtVcKw9mXC3Vn4cfVCmcgWlU3ymJn+J9nwvERzEzRE6F0BHiIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiCL7ROOnv7eP9Kq8q0donxdP4eP9Kq8qg8pNL8o+Ujdew4X0ghlqH7kMUkrsZ3WNLjj6F81KdnPDULvm7/ytUTc7CLe2pspnDGcH2rq5tMy8LwXX+Q1fcu9i58FV/kNX3LvYrrRWnNajvJ3flp9bnUpTwVcPIKvuXexcG13DyCr7l3sV2etPWma1HeTu/J1udSkvBVw8gq+5d7E8F3DyGr7l3sV2oma1HeTu/J1udSk/BVw8hq+5d7E8FV/kNX3LvYrs+lEzWo7yd35OtzqUn4Krx/IavuXexPBVf5BV9y72K7ETNajvJ3HW51KT8FXDyGr7l3sXPgq4eQ1fcu9iutEzWo7yd35OtzqUn4LuHkFX3LvYngu4eQVfcu9iuxFnNajvJ3fk63OpSfgu4eQVfcu9ieCq/P8Aq+5d7Fdn0osZrUd5O78nW51KU8FV/kNX3LvYuPBdef5DV9y72K7ETNajvJ3fk63OpSfgq4eQ1fcu9i58FXDyCr7l3sV1plZzXo7yd35OtzqUp4KuHkNX3LvYngqv8hq+5d7FdaJmvR3k7vydbnUpQ2q4eQ1fcu9i48FXAfyGr7l3sV2Ima1HeTu/J1udSkxa689NDVdy72KfbOaeent9WJoJYiZgQHtLSRujtUu+lFt3DIVN0totorxw/p4tLxNdPNwERFPtcRE9CAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiCL7Rfi7/AG8f6VV6tDaL8XT+Hj/SqvCoPKTS/KPlI3XsOVKNnPxhd83f+VqjAKk+zn4xO+bv/K1aGStLs/GH0tuxKz116u4Udva11XUxQNccNMjg3JXYUK2nfwKh/Cu+qug3+8zdrvVbUxjMI2zp51UUpH75bL50o+9Ce+SzedKLvQqZ6EVVzottiPVt9Up1rm98lm86UfehPfJZvOlH3oVMZXKznRbbEep1SNa5vfJZvOlH3oT3yWbzpR96FTK5CxnRbbEep1SNa5ffJZvOlH3oT3yWbzpR96FTXSh4LOdFtsR6nVI1rl98lm86Ufehc++SzedKPvQqYyucpnRbbEep1SnWuX3x2bzpR96E98lm86UfehUzlMJnRbbEep1SNa5vfJZvOlH3oT3yWbzpR96FTSFM6LbYj1OqU61y++WzedKPvQnvks3nSj70KmVzlM6LbYj1OqU61y++SzedKPvQnvks3nSj70KmsomdFtsR6nVI1rl98lm86UfehPfJZvOlH3oVNBclM6LbYj1OqRrXJ75LN50o+9Ce+WzedKPvQqZK5HpTOi22I9TqlOtcvvks3nSj70LtUlwpK9jn0lRFO1pwTG4OAP0KkckcOpWHsyAFurcf88fVC38m5dtL1bxY1UxETi+drd4opxiUzREVmap1oiICIiAiIgIiICIiAiIgIiIGUREBMoiAiIgIiICIiAiIgIiICIiAiIgi+0X4un8PH+lVcFaO0X4un8PH+lVcqDyk0vyj5SN17D9dSk+zn4xO+bv/ACtUXClGzj4xO+bv/K1aGStLs/GHu27ErQUK2nfwKh/Cu+qpqoVtO/gdD+Fd9VXjLehWnhHvDRsP3IV2UXJCYyubpNxhcr60lLLW1UNLFu85M8Mbk4GT2qReLu+fzTvf3LZsLnbW8TNlRMx/TxVXTT+soyik3i7vnZSd7+5c+Lu+dlJ3v7l9vpd77udzz0tGtGVwQpP4u752Une/uTxd3zspO9/cs/S733c7jpaNaLopP4u752Une/uTxd3zspO9/csfS733c7jpaNaMLnCk3i6vn80739y58Xd8/mne/uWfpd77udx0tGtGEPFSbxd3z+ad7+5c+Lu+dlJ3v7lj6Xe+7ncdLRrRdFJ/F1fOyk739y48XV8/mne/uT6Xe+7ncdLRrRkLlSYbOr5/NO9/cni6vf8ANO9/cn0u993O46WjWjKKT+Lu+fzTvf3J4u752Une/uWfpd77udx0tGtGCFxlSfxdXz+ad7+5PF3fOyk739yfS733c7jpaNaMhWHsx+11b+HH1QvC8XV8z0Unen2KXaLsVZYaOoirOa35JQ9vNu3hjdx2KVyJcLxZXuK7SiYjCfZ8be0pmjCJSNEymFeGiIiICIiAiFEBEQICIiAiIgIiICJ0IgIiICIiAiIgIiICIiAiIgIiICIiCMbROOnT+Hj/AEqriOKtDaL8XT+Hj/SqwVB5SaX5R8pC69hxgKUbOPjC75u/8rVF+ClGzj4xO+bv/K1aGStLs/GH0tuxK0FCtp38DoPwrvqqaqE7Tv4HQfhXfVV4y3oVp4R7w0bDtwr4ouEyubpN6Wnft/bvnDPyq5lTOnft/bfnDPyq5grryW/Zr8fho3vtQIiK0NQREQEX4mlZBE+WVwZGwFznHoAHEleCNoGlj0X2iP8A3r5121nR26ojxl9LOxtLTsUzPhCQoo/4wNLHovtF/fXbteqbJeqh1PbrlTVUzW75ZG7JDc4z/iF5pvNjVOFNcTPjD3VdramOdVRMR4S9VEXQut9tljjZJcq2GkZIS1jpTgOIGcL6VVU0xjVOEPlTRVVPNpjGXfRR7xg6V8+0WP6f7kG0DSxPC+0P99fHrdhtxvh9up2+xO6UhRedatQ2q+OlFtroKsw43+adndznGfUV6K+1FdNcc6mcYfGuiqiebVGEiL8ySMhjdLI9rGMBc5zjgADrJXRGorMf+LW/8ZZ7V6eXoIulBerZVSthguNHLI7oYydrnH5ACu6gIiICIiAiIgIhRAREQEREBE6kQEREBERAREQEREBERAREQERMoCIiAiIAgIiICIiCMbRPi6fw8f6VV6s/aJ8XT+Hj/SqwyqDyk0vyj5SF17DjoUp2cD/WF3zd/wCVqiylOzj4wu+bv/K1aGStMs/GH0tuxKz1Ctp38DoT/wDld9VTVQnad/A6D8K/6qvGW9CtPCPeGjYfuQr1ERc3Sb0tOfb+2/OGflVzKmNOfb+3fOGflVzhXXkt+zX4/DRvfagREVoagiIg6F/+0Vx+ay/UKzCCd1vHHALT2oPtFcfmsv1CswN+xHyBVHlL27Pwn4XTkr2LTxj5foEnrKn+xY51VUD+Zv8ArtVfqwNix/1qqPmb/rsURkvS7PxTWWNCtPBdoVY7cc+DbUP/AM8n1FZwVY7cDi3Wr8PJ9RXLLGh1/wDfzCjZE06z8/aVQgkdZwgJHWuD0oufOlrW2GdN5z0/5H9dWuqn2GcXXn+x/XVsK/5E0Ojz95c3y9p1p5e0OrdbbTXq21VtrGGSmq4nQSsDiN5jhgjI4jgVkTbrs/sez/U9DQ2SKVkFTR8+WTP5zddvubwJGcYA6VsVZg5VY/10s/8AVv8A7XKVQ6X7ANlunDp6ya2fBO+8u517XmUhkfwns4NHA/B7c9KvBV9sEP8AslsOPvJfzz1Gtq21vW+h62ojpdL08NrEgip7pUF0jJiW5+xaRun7LgexBcyLOmi+UNfZLZcI7lSuvl+nqGR22go6fcG7uEuc7dBO6DjtP5V4+pdp+2qwObcLvTy2ikkdhgNvj5nJ6GkneOflOUGo1wTgZPQqu2J7X5do0FVb7pBDBdqNgkcYQQyeMnG8AfsSDgEdHEEdi6O2/U20C3U1dS6fs7Y7G2j36q6ggvaCCHhuXDdwOsAnjwQejoHbDJrnaBedPwUVM220UcklPVse4vmDZGsyR0AHJPqVnrFeye4aytl/qpNEW+OurjSlssb2NeGxb7eOC5vXu9avjQF92vV2qaWDVdkgpbO5khmlbCxpa4MO7xDyeLsdSC3EVebV9sNv2aU8dOyAV93qGl8NLvbrWN6N+Q9Tc9AHE4PyqF2uv286qtjb3SzWq2QTN52ClfGxjpGniMBzXEA9W8QgvdFQmgeUDdI9RDTWvaSGmqDP7m91MZzZhlzjdlZkjBPDeHRw4Y4q8rlcaW02+ouFbMyCmpo3SyyO6GNaMkoOyizx47tbbSNS+BNAUVNQRnec2WqYHyc2OmR5OWsHRwAJ4gcSmrNZbZdljKetvtZabrQzP5sStgBjD8ZDXFoY5pODjqOEGh0Wb7DtN2y68iNZp6itz6WOYRTGmiiBjPAkESPJ6D2Kwtq+0XV2ht6W0aXjrLbHE18tymeXRxuJI3SxuCMcOJOOKCzkWbNG8oq+89eKvUL4awNp2Nt9upYBGZahz8BoPF2MdOc8OgZwvAvO3rahbrsTWCO2SDEgt81vDGlp6B8Mb5B6M5QayRVZtK21t0DarbD4ObNfrhStqBSvcWx04IGS89J+FkADicHiFF7NXbdtY2WK/wBBcbPQU1Q3naen5tjXSM6iA5rsA9W85BfaLO2huUNfqTUzLBrWmhLHT+5ZKlsXNS00m9u/DaPgkZ4HGMdPFaJQEREBERAREQEREBERAREQMIiICIiCMbRPi6fw8f6VV3WrvuFupbpT+56yITRZDt0kjiOjoXme8rT/AJtj/vO9qrOVsjW17t+lomIjDD748G1Y28UU4SqNSjZx8YnfN3/lapr7ytP+bY/7zvauzbtOWq0zmeio2QyEFpcCTwPyla1y5P29hb0WtVUYROP88Hq0vNNVMxEPSUK2mgmjocf8131VNepdK5Wehu7WMrqdszYyXNBJGD9BViyhdqrxd6rKn9Z4tazqimqJlSvT0JjrVue8uwdPg2P+872p7ytP+bY/7zvaqnmxedqn14NzrVOpWmnON/t3zhn5Vcy8in0lZKWeOeGgYyWNwc1wc7gR19K9dWDI2TrS5WdVNpMTjP8ADWtrWK5iYERFMviIiIPP1B9obl81l+oVmFo+CPkC0/qD7Q3H5rL9QrMI+xHyBVHlL27Pwn4XTkr2LTxj5FP9ivxrqPmT/rtUAwrA2K/Gqp+Zu+u1RGS9Ls/FNZY0K08F2joVY7ceFutX4eT6is5Vjtx+11q/DyfUVyyxodfl7wo2Q9Os/P2lUCdKIOlc+dLWrsM+yvP9j+urYVUbDem8f2P66tdX/Imh0efvLm+XtOtPL2gWYOVV8dLP/Vv/ALXLT6zByqvjraP6t/8Aa5SqHW9sE/imsH9CX889eJynseLiPh/xGD8j17mwT+Kawf0Jfzz14fKcG9s4jx5xg/I9BGuShSwml1FUuiYZ2zQRtlLRvNaWuJaD04zg4VrbVKOCt2cakiqI2vYLfNIAR0Oa0uaflBAKq7knn/RupMdVTB9RytjaTx2fajz0eDaj82UGb+TO942mDjhr7fOCO3iwrQm14A7MtTf1fL+RZ85Nf8ZrAOj3BUfqLQm13hsy1MerwfL+RBRPJa469uI6vBj/AM7GtSHAGeAWWuS27Gv7jjo8GP8Azsa1JIznI3NJwHDCDEd71NTaq2lS3q+Plfb5biHSsY3eIpmPwGAdfwGgY9JWhW8pXQoy1sV63W9lFw+ss56ZqG6W2i26etZuMtl1a2feH2LWyFrifkGStwMige0FjIi0jgQ0YIQYs2raltus9cV17skFRHTVEcQzNFuPL2s3SSOPYFfO06sr7hyfWVpLzPPRUMlSR0lpMZeT6O1W3zEQH+6jz/RC+Vfb6W50M9DWQRz01RGYpYnjLXsIwR6kGcOSpLTe+W/RvAFU6jiLM9JYJDvf4lqs7lEPpmbKroKndLjJAIs9POc63GPoyqdotGVdBtaq7ZstvW/U22N8sk1UQI4MENdCXYPODJDeI6fkyppf9lu1DaU6lp9XX2z0VDTv3xHRtLgXYxvbgAy7GcZdwyUHjclFtSbvqB8efcfueEP7Oc33bv07u8rY23DGyrUmOB9yfrtXr6F0JaNn9jZabTG7d3ucmmkwZJ5CMFzj9GABwAXkbbj/ALKtSZ8l/XagpDkw2+krteVs1TTsllpKEywOcM828yNaXD04JGfStMXDT9putVTVlfbKKrqKU5glnha90R/6SRwWb+SuCNb3f+rf/axag9KCtdqWxOg2l3KkuTrrPbaqCLmHuZEJBJHvEjgSMEEnj6ehTOhitOkLHRWx9ZDTUlHAyCN1TK1hLWgAEk448FSW3nbNeLPeptL6cqHUfudjTWVUf+9L3DeDGH7kAEZI45OOGOPsaS5OtnrqGnuus6quu9znY2WSJ1Q4RxkjO7n7JxGeJz9CCktsNTS1O0rUFTQTwzU0lQJGSwPDmOPNtyQRwPHK2lQOL6Knc45JiYSfoCxXtdstv09tAvNrtVM2koqZzGxRNJIaDE0npJPSStpW3jb6b8Ez6oQdlERAREQEREBERARCiAiIgIiICIiAuMhRLahda6z6Y91W+qkppvdEbd9mM4OcjiqjOv8AVPnys9bfYom+5YsrradHXTMzhj9kzcMiWt8suloqiIxw++LROR2pkdqzr4wNU+fKz1t9iDX+qfPtZ62+xaeclhsz6cW7mteNun14NFZHamR2rO3v/wBU+faz1t9ie/8A1T58rPW32JnJYbM+nEzWvG3T68GicjtTI7VnX3/6p8+Vnrb7E8YGqfPlZ62+xM5LDZn04ma1426fXg0VkdoXOR2rOvjA1Sf+OVnrb7Fx4wNU+fKz1t9iZyWGzPpxM1rxt0+vBorI7UyFnXxgap8+Vnrb7E8YGqfPlZ62+xM5LDZn04ma1426fXg0Xkdq4yO1Z0O0DVXn2s9bfYg2gaq8+Vnrb7EzksNmfTiZrXjbp9eC/L+QbFcRn+Sy/UKzEBho+QL359d6mqYXwy3qrdG9pa5pLfhA8COheDwAAChMrZQovlVM0RMYa1gyNky0uVNcWkxOOH6Cn2xX411PzN/12qAqfbFfjXUfMn/XatfJel2fi2Mr6FaeC7gqx24/a61fh5PqKzgqx24/a61fh5PqK5ZY0Ov/AL+YUbIenWfn7SqBchcYQLnzpa1thv2V4/sf11a6qfYZ9lefkh/XVsK/5E0Ojz95c3y9p1p5e0OneLk2z2mtuT4Zp20kD5zFC3ee8NBO60dZOOCyXte1LcNpl/pblSaYvVFHTUvudrJqd7nO+GXZOG4HTjHFbATHy+tSqHUbsI2gVFFbLPomv03d4JWGVjK0xERYJfJ8LIBb2dfFefyhtZz3yln0fQaevL5KSsjlkrBTudE/daThm6CT9l0+grQSY+X1oMw7BNV1Gg56u23DTl6lbdqmANmjpnBsR4s+EHAcPhZyrP22a7Njs1VpyGxXW41F2oZo2zU0RdFDvZZ8IjJz14wrP9frRBjjZVdLhs81Wy91WnLzVwCnkgdHFTPDxvY4jIx9z0elakvNK3XegqqnhZJTi8W5wjZO3ddGZGZaHDqIJGVI/wD+9KdCDFGl75qHY5rB1ZNa3RVkTHU89LUtc1sjCRkAj0gEOGRw61fGi9o+0HaBeaGWl0rTWuwNkBq6mpc/MjMHIjLgMnr4NPpIVuPhjkxvxtdjo3gDhfvCDPG3bYrc6y6T6p0zSOrPdPw62jiH+UD8YMjB90D1gcc8eOeET0jt41hoahjslZRx3CKnaI4mVzXxzRNHQ3e6SB1Agkdq1ngL8PgikcHPiY5w6C5oJQZ6t+p9rW2SVtJQwe9mzFw56thjfH8H/pe74Tz6G4HacLQEQNHRxsqKh0hijG/NJgF2Bxc7GAM4yepffCpnlRXS40WjaCjpTIykrazm6pzSRvAMLmsPoJGcde6gpCa2ag03ri5e8y4VV0qIXyPFZZnOmLonOz8LdB+kHIyOterJth2q2Ih1ddbhCB1VtAwD6d5gVqcl642ZulKy3RTwi8e63y1ERIEj2YAY4DpLQMjh0HParhus9BTW+ea6SU8dE1hMrqkgRhvXvb3DCCldk/KFq9SXunsGp6WmimqjzdNW04LGuk6mPaScE9AIPTwxxXsbfNaOprHXaRprFda2puNI1wqYIS6GMF/QSMkn4PR6QqP0taqfVO2Ong05C9tv8LGqhDAQIaZkm/vegYAx8oC2aPUgyFsh1JcNm2oqu5Vmmb1WxVVN7nLIaZ4cw77XZ4twejC1bp+8M1DZKK6spqikZVxNlEFQ3dkjz1OHUV6GPSfWuUGWOUNs+u9Fq+r1JTUc1RbLg1j5Joml3MyBoaWvx0A7oIPRx9C9jSfKG1XcLfT2Oh0q28XZjGxMnic/D8DAc9gHD0/CA+RaOIBGF+Y4Y4c83GxgPTugDPqQYu2s6W1XZ9RurNVkVFZdI21Lp4WHmt8jBiBHDLMAY7MfKtM7K9f1et7a7n9N3G0MpYo2tnqB/k6g4IO4SATjAP0hTl0bHjD2tcOnBGeK/SAhREBERAREQEREBERAREQETKICIiCDbYviefnUX5SqNKvLbCCdIHAz/nUX6VR5Y7713qVI5QR/leUfK/8AJqf8T/afh+Eyv1zbvvXepNx33rvUoPCVgxhxlc5Tm3H7l3qTm3feu9SYSYw4K4X63Hfeu9SbjvvXeophJjD8ov3uO+9d6iuObd2O9RTCTGH5Rfrm3D7l3qTcP3rvUmEmMPynQv1uO7HepCx2PsXeophLOMPzlcL9c27713qKc27713qTCTGH5Cn+xb411HzN/wBdqge44dLXeoqfbF2kaqqTgj/M3dI/62rfyXE9bs/FG5Xn/CtPBdgVY7cftdafw8n1FZwVY7cftdafw8n1FccsaHX/AN/MKNkPTrPz9pVD1oi461z50tauwz7K8/2P66thVRsM+yvP9j+urXV/yJodHn7y5vl7TrTy9oEXwr21T6GobRPZHVGJ4hfIMta/B3SfRnCzdrvaPtg2d1lNR3m6WnfqIzJHJT0zHseAcHpAI4kdXWpVDtMIqB2J7bL/AKs1gbHqKpp5m1NO91MY4GxkSs+ERw6ct3v7qv7oQEQ8FlfWPKI1e3VF0isFfSw2yKofFTh1Ix5LG/B3snickE/Sg1Qio7ZJqbatrt1Heaq42oWJtUY6gOgayWVrfsg0AcOzOQpTt9uV9tOzuonsclTFIZo2VM1NkSRwHO8QRxAzugkdAJQWQiz5yXrzqS41l4iq6usq7PHEwsdO9z2snLuhrj2tySB2BaDQEREBebqHT1s1TaZ7Vd6RlVSTgbzHcMEdBBHEEHoIXpLp3iSshtNbJb2CSsZBI6BhGQ6QNO6PXhBR9y5K0ArTU2TVFTSNByxlRT77mfI9rmn/AAX7byZa65uYL7resq4WH/dtic4/QZHkD1Kutl+qNb1m062NFxu9VUT1QbXwzve5pjz/AJTnGng0AZ6hggYWvh0BBGNEbOdO7PqN9PZKPcklxz1TK7fmmx0bzuz0DA9Ck6IgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiDgtDhggEelcc2z7xvqXEk0cLd6SRrG9GXEAL5e76TyqDvG+1eZmP5eoiZ/R9ubZ9431JzbPvG+pfH3fSeVQd432p7vpPKoO8b7VjnUnNq1PtzbPvG+pObZ9431L4+EKTyqDvG+1PCFH5VB3jfanOpObVqfbm2feN9Sc2z71vqXx930nlVP3jfanhCj8qg7xvtTnUnNq1PtzbPvG+pObZ9431L4+EKTyqDvG+1Pd9J5VB3jfanOpObVqfbm2feN9Sc2z7xvqXx8IUnlUHeN9qeEKTyqn7xvtTnUnNq1PtzbPvG+pObZ9431L4+EKTyqDvG+1PCFJ5VB3jfamNJzatT7c2z7xvqTm2feN9S+PhCk8qp+8b7U8IUflUHeN9qY0nNq1PtzbPvG+pchjW8Q0D5Avh7vo/KoO8b7U8IUnlUHeN9qc6k5tWp2FWO3H7W2r8PJ9RWN4QpPKqfvG+1VttrqYZrdaxFLHIRO/O64HHwPQo7K9UTc68P++8JTIlMxfrP7a/aVRoOlDx4rkLn7pS1dhow68f2P66tdVTsO6bx/Y/rq1lf8iaHR5+8ub5e0608vaBZ55SNK2r11o2KQb0TzuPafugZ4wR6loZZ95QgMm0jREYP3bD/wDyY1KodUV0ZVbLdpMzacO52y3DnI88OcjDt5v95hA+lbWt1bBdKCmrqZ+/BUxNmjcOtrgCD6is6cqTS7aa9WzU0UeBWMNJOQP/ALjPhMJ+VpI/7VP+Tlqnw/oFlukfmos8ppTx482fhRn1Et/7UEl2t6q952gbtcmP3ah0Xuen48edk+C3HyZJ+hZR1BpQ2TQemLw5pE14lqpzkcRE3m2x+sbzv+5XByhbjNqbVumtn9E4udUStmmx1OedxmfkbzjvpXT5UNvhtdm0hSUjBHTU3PwRsHU1rIwP8AgmnJpcH7MYcdVbUD/yCkurNrOjtG1clBebq1laxoc6liifJJhwyMgDHEdpUV5Mjg7ZqQOq4VA+quOUXpqzTaDr72+20vhSF8DWVgjAlAMjW43ukjBIwUEl2d7T7BtAqK6lsVJWQMoWse4zxNja7eJHwQCfvevC7erNqmkdEVPuS93VsNZuCQU8cT5JC05wcNB6cHpVO8lAnwnqNo/5FP8AWerB2+abs1Zs+vN3qbbSyXGlgYYasxjnWf5RowHdOOJ4elB3tI7atM6xqbkyk91UlNbqcVM1VWhsUe4TjP2RI+nCj9z5TujaOsNPSUt2uEbTg1EULWMPpG+4E+oKoNhehqLXeoa+gudTUttsEDJ5qWF5YKkh+GtcR1A8eHH5FemqthuiazTdZT2+x0tvqmQPdBU0+WvY8NJGTn4QyOIOUEi0RtH09tApZZbLVOdLBjnqaVu5LFnoJb2HtGQu7qnWdg0ZTRVN/uUVDHM4sj3g5xkcBkgAAk8Flbk/XKa37ULSxjiBWslp5Wg8HNMZdx+RzQVq/UOmrPqai9z3q2UlwiaC5jZ4w7cOOkHpB9IQQay7edG33VNHY7XBcJqqvl5ptQaYRszgniXEOPR2Kbam1dZNG0Da++3COhpnv5tj3gnedgndAAJJwCsg7HwG7VdOtP3NaQP7r1sa72K16hovcd3t9LX0xIcI6iMPaD2jPQeJ4oIHbuUDpG86gt9ktbLlVy107YGTcxzcbSegneIOPkC+mrdvujtKV0lvM1Rc6yI7skdCwPbG7sc8kNz6AThZjsNoNdtFpbRR1UtvE10NLHPD9nC0yFuW+kDoWp6HYbs+oaBtKdOU1ThuHTVBc+Vx7S7PA/JhB19EbdtJ63uMdrhdVW+vl4RQVjA3nT2NcCQT6OBVirEO0GyDQW0G5W+2TPa23VDJaZ5Pw2DDZGce0ZAz6Fse4Xl1DpWpvO4HPhoXVe51EiPfwg8PW21zSeg5xSXWudJXFocKSmZzkoB6C7qb/wBxCh8HKf0qZwyrtF8pInfYyvhY4EduA7J+jKpvY9b2652q0Ul9Pu4SOmr6gTfC56Rrd4b2ekbxBx2DHQtN7T9N0OpNCXaiqYY3GKlkmp3lozFIxpc1zT1cRj5CQgidfymdBUYJikulUAP/ALVKWg/S8tVm2i5xXq00Vzp2vbDWQR1EYfjeDXtDgDjrwVknk+T0s+0Kntdwt1HX01xgkaWVMDZRG9rS9rm7wOOgg/KtfQxRwRMihYyONjQ1rGjAaB0AAdAQftERAREQEREBERAREQEREDCIiAiIgg22P4nnIz/nUX5SqNwOwepXlti+J5+dRflKo09Ko/KHSvKPlf8Ak1of+0/DjA7B6lxujsHqXKKDWHA3QfuR6kwM/Yj1IiMYGB2D1Jgfej1JlEMAAdg9SEDsHqRclDB+cDsHqTA7B6kJXIQwfnA7B6lzgdg9SFEZwMDPQPUm6Owepc4QhDBwAOwepc4HYPUiIG6OnA9SAY6MDK5K4QwEyiBBauwz7K8/2P66tdVRsNHwrz/Y/rq1+pX/ACJodHn7y5vl7TrTy9oFn3bzl+1nRDAfuof8apvsWgXODGlziA0DJJ4ALN+2+9W6fa7pKqiuFLLT0Xuc1D45Q5sX+c5O8QcDAGfkUqh1tbZdLnVuzu7UUMYfVQs91U/bzkfwsD5RvD6VQPJv1O2xa7db55gylu8BiJJwBIzL2H1b4+laspK+juUHO0VVT1URP2cMjXtP0glYo2k6bl0Rr662+EvjhinM9M5pwRE/4TcH0AkfQgt3Y9/tD2uai1vM0vpqNxZSl3UX/AZj5Imk/wDcvvysG/6M02eyoqB/4NU52EaXbpnZzbt+Lm6m4D3dMCMEF4G6PoYGqvuVRc6Cro7BSQV1NNURVEznxRyBzmDcAyQDw49qCTcl4/7N5h2XKcf+LF6vKH47Kbr+Fp/zzVHOS/d6CPRdXbX1tOyr8IyPEDpAHlrmMwQ0nJHA+pe3yi7jSQ7M7hSuqYBUSS0+5CXjfd/lWk4b0ngEEC5J/wBs9R/gKf6z1ae3P+KnUXzdv5xiqXkr11LSXi/x1NTBC+aCDm2ySBpfhz84z09I9atPbvX0kGzG+081VAyaaFjY43SAOeecZwA6Sgqfkq/Gy+fMWfnVpG4/a+q/Av8AqlZo5MFbS0OrLv7pqYIeeomiPnHhu+RIMgZ6StH3y4UdDaqqWqqoKeMxP+HLIGj7E9ZKDIGw4/7VdO5/5r/zT1s2T/dP/on8ixbsZqYKDadp6arnjhjbM7ekkcGtGYngZJ6OJC2TWXOipKE1M9ZTRQlpIlfK1rTw7ScIMbbI+O1fT2fLj9V62p1D6FibZXWQUm07T9TVSxwwtrsukkcGtaC1wBJPR0hbQkuNHDSCqlrKdlPwPOulaGevOEGNNHH/AGyWr+vv/cVtXoH0LEmkaunj2s2utfPEymF7Ehle4BgZzxO8T0YwelbUZWU8tP7oZUQuhxnnA8FuPl6EGOuUB/Gtf+rhD+YYtc0NJFX6bp6Sdu/DPRsjkb2tdGAR6iVkHbnVU1x2n36elqIp4iY2h8Tg5pIhYDgj08FrzTFwpLhYqGSiqoKlgp4xvRSBw+wHYgyvWaa1FsF1/Q3Z1K6qt8M7mwVAIDKqIggxn71+6egjpGRkKwNccoe06h0rWWfTNBdZbncIXUxEkGOYDhh3QSXOwSBjtyojtZ1hdbPttbcK+l91RWaaJ9HRzEiN0W6DvD0lxJ3uPEDswphR8qSxsdmbSldA4/ZOglid+XdKDjk9bIbrp+udqrUFM+jl5p0VFSScJGh32Ujx9zw4AHjxJOOCvxQTRW2fSGuKllFQ1ktLXv8AsaSsZzb39u6clrvkBz6FO85QEREBERAREQEREBERAREQEREBERBE9pdmr77ps0dupzUTmeN+4HAcBnJ4lVT4stW+Z397H+0tBIoq+ZIsb1adJXMxP6fbDgmLjlq2udn0VnETGOP3x4s+eLLVvmd/ex/tJ4stW+Z397H+0tB4Rambl32p9ODczovWzTuniz54stW+Z397H+0h2Zat8zv72P8AaWg0TNy77U+nAzovWzTuniz34stW+Z397H+0uRsy1b5nf3sf7S0HhMJm5d9qfTgZ0XrZp9eLPniy1b5nf3sf7SeLHVvmd/ex/tLQeETNy77U+nAzovWzTuniz34sdW+Z397H+0ufFjq3zO/vY/2loPCJm5d9qfTgZ0XrZp3TxZ78WOrfM7+9j/aXPix1b5nf3sf7S0GiZuXfan04GdF62ad08WfPFlq3zO/vY/2k8WWrfM7+9j/aWg0TNy77U+nAzovWzTuniz54stW+Z397H+0niy1b5of3sf7S0GiZuXfan04GdF62ad08WfDsy1b5nf3sf7S48WWrfM7+9j/aWhMImbl32p9OBnRetmndPFnvxZat8zv72P8AaXPiy1b5nf3sf7S0GiZuXfan04GdF62ad08VebKNMXfTrrmbpRmmE/Nc3l7XZxvZ6Ce0Kw0RTF1u1N3sosqP0hCXu9VXm1m2r/WdW51rnb4Ltbaq3VTXOp6uF8Eoa7BLXAg4PVwKrdvJv2etAAoa4AdXu2TH5VaKLYayOaK2f2HZ/S1NLYqeWGKpkEsgklL8uAx19HBfPUuzPSWsK9lwvtmhralkYibI572ncBJA+C4Z4k9PapOiD8xxMijbHG0MYwBrWjoAHQFWUvJy2fzTyTvoa4vlc57iKyQZJOT1+lWeiCt7dyftCWuvpa6noq0T0szJ4nOq3ndc0gjhntC7mp9imj9X3ye93alq5K2oDWvcypewYa0NHAdHAKeIgq//AOm/Z9kH3DX8DkD3bJ7V7GqtjekdZ3bwreKSplquaZDvMqXsAa3OOA+UqcIgq53Ju2euYWGhry301sntUm1hsy05rqmt9PeqeeWO373MCOdzN3IAOcdPBoUrRBWA5OOz7JJoK457a2T2r1rnsZ0jeNP2qwVdLUuoLSXupWCpcHNL+nLuk/SpyiCr28nDZ83+Q1xz21sh/SvXrdjWka/S1Hpeelqja6Od1TDGKlwcHu3s5d0kfCPBTlEFXDk3bPR0UNf+Oye1SWl2X6botGVOjoKeobaKkuMkfPuLjvODjh3SOICliIKuHJv2eg5FDXfjsntX2qbFpnYNpO+6gstFOXPZHmKSoc8SSb27GMnoGX8cdSsteHrbStNrXS9fYaqQxMq48NkAyY3ghzXY68OA4IKG2X0NTtz1LX3TXFX4QpbWxnM0TAIo96Qu4Ybx3QG9Gck4ye2yK/k6bPatpMVsqqJx66erkGPocSP8FTVs0rtQ2M3+W4W2zzVjHDm5H0sRqKeoZnI3g34Te0ZwQpn48tpddDzNBs6l91OGN/3PUOAPbulo/wASgqHaPpN2zXWlTbaOuln9y83U00/2MjcjebnH3QI6R6D1raVonmqrVRz1Ld2eWCN8jcYw4tBP+OVQuitiWo9U6q9920NwZvyid1G8h0k7hjdDg34LGDA+COOBjgtCjoQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBAgRB8K+okpKKoqIoH1EkUTnthZ9lIQCQ0ek4x9KpC48puqs9QILpoG40EuN7m6mp5t2PkMYV7LOPKepDcNY6Uo9/cbURPiLsZ3d6Zjc/RlBZ+yza9b9pza2KKjfbqykIc6mklEhfGeh4IA4ZyCMcOHap+eCxHQ1t52P6/LzGRXWycxyxZw2oiPSP6L24IPyHqWy9O32i1PZaO8W6XnaSriEsbusA9IPYQcgjtCCsdW7erjpC51VLX6EuUdNFO+GKrln5uOoAJw5pLMcQM4yva2cbVrltArubOjbjbLcYXSsuMshdE8ggBoO4Ac5PEHqXjcp9gOzunz0C5Qn/xkU32YY8XWmsdHgyn+oEEnRebctS2SzAm5Xe30YH/AD6hjD6iV97VdqG+UEVwtlVDV0kwJjmidvNfgkHB+UEIO2i6VffLXaml1wuVFRgdc87WflK+Nj1LZtSxTS2a50lwjgk5qR9PIHta7GcZHXghB6aLh72xtL3uDWtGSScADtUUq9q+haKcwT6qtIkacENnDgD6SMhBLEXnWfUVn1BCZ7RdKKvjHS6mmbJj5cHh9K9E9CAii922n6Lskxgr9TWuGVpw5nPhzmn0hucL0LFrHT2pwTZrzQ17mjLmwTBzmjtLekepB7CJ0heVdtU2Gwtc663i30O7xInqGsPqJyg9VF5jdS2Z1mjvfhOkbbJWCRlW+UNic09B3jheTbNqOirxXNoKHU1rmqnHdbGJgC89gzjP0IJSuCQ3iTw68lckgDJ4KieUZtKoorKzStqrYqioq3g1pgkDuZiac7hI6C52OHYD2oL1BDhlpBHoXKq3Y/rbTFu2bWGkrNQ2unqYqbdkimqmNew77uBBOQrJt9yortSMrLfVwVdNJncmheHsdg4OCOB4ghB2UXi3zWenNNHdvN7t9C/GQyeZrXkf0en/AAXStO03Rl7qG01BqW2TTvOGx88GucewB2MoJOidK6tfdaC1xc9X1tNSRj7ueVsY9ZIQdpF5Vo1XYr/TVNVa7tR1lPSuLJpYZQ5kZxni7o6OK8+37StHXW6C10OpLZUVjnbrYmTDLz2NPQT6AUElRfKpq6ejgfPUzxQQxjL5JHBrWjtJPAKKzbXtBQTGF+q7Vvg4O7NvD1jggl6LpWm92y/UoqrVcKWupzw52nlD257MjoK7qAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAs9co9+5rvRpxxyP/ANiNaFWeOUgca80aT2j/APYjQe7yi9mvvgtHvrt8WbhbWEVLWjjNTgk59JZxPyF3YFDOTjtJFlu3vRuEuKK4PLqN7jwin+8+R/1gO1abewSscxwBacggjIIWOtsezx+zjVhkoQ+O11rjUUL2nHMkHLoweotOCPQR2ILr5TfHZxF/WMH1XqZ7Mm52c6baeu2U4/8AAKjtbbQmbQ9htPNVSAXWhuNPDWsHDeduv3ZB6HDj8ocOpXnswIfs600R0eDaf6gQZr276As+gr9b4bSaksrYHzPFRJzha4PxwdjOOPXlX3sF/ilsGSTlkv556qflWfGaw/MpfzgVsbBP4pbB/Ql/PPQVdykNnFmskDdW0ZqRX3C4blQ2STfjO8xzstBGW8WjrwpHyVTvaTvP9Yj80xfflT8dD2z+s2/mpF8OSpw0nef6xH5piDw+VFq+virqHS1NUPho30/uuqaw458lxDGO7WjdJx1kjsUv0DsU0DVaJtk09rguktZSxzSVb5HbznOaCd3dI3QCSBjs48V8Nv2yeu1tDTXyxRia50UZifTFwaaiLO8N0nhvA54HpBPoVF6Y2g6z2YVL6GlqKmkbG4mS210ZMYPX8B2C35W4QWVs60FLorb7VUFNTVZt9NBM+GocxxaY3saWNc/GCRkj/tUi5T1+vVp01baS3STQUddO9lZNES3IDQWxlw6A7JPp3cL2dlW3G3bQp/BVZTeDbyGF7Yd/eiqAOkxnpyOktPHHRnip5qW+2bTtolr79VU9NQswHOnGQ49QDeJcewAEoKd2I7IdL3rQlNeL/Y46urrXyOaZy4BsYcWt3QCMZAznpOVVe0/T79l+0l7bDPPTNhbHWUUged+IOz8HPWA5pHHpHA5V0jlEUt2qXUmj9IXu/SR8MxsEbQOo4AcQPlAVKbabxfL9qyOs1BY3WOrNHG1lIZecPN7z8OJ7SSeHoQa1slYzVOk6KslDo23OhZI9rDgt5yMZAPV0lZO20aEtmgdWQW21y1U0M1Iyo3qpwe8Euc3G8AMj4PXxWpNmnHZ5ponzZT/mws/cqLhtBocebI/zkiCQ6J2NVm0XS1kuWqb/AFUdthpGMt1vocARxDoe4uBG+7pOBn09Qg+2fZTDs1qqCahrJqugrt9rOfA5yN7cEtJGAQQcg4HQVozY/wDxY6a+YR/pVdcq4DwFYD2Vkn5pBKth97qda7LoortJJNJEZrdJKXfDewDAJPbuuAz6FRO2/Z/Ztnl/ttvspqRTz0hmeJ5ecO9vlvA4HDAVycl852c1H9ZT/VjUA5VA/wBcrMf/AI4/nXIPf2bbCNIar0NZ71cBcm1dZBzknM1O63O8RwGOHQrB1DbZNmmya40ulBUGS300jqcvPOyM3n5c7o4lu853R1L67Ev4q9OfNf13KazyxU8L5ppGRRxtLnPeQGtAHEknoCDK+wPQlu19qG8XLUtLLdYKaNhDp5HFssz3HJe7OXEAZwT1/IvQ5Q+zGyaUpbZeLFSNooJ5nU09O1xMe9ulzXNBzg/BcCOjoU+ufKJ0jbKzwXp23119qHPIa23whsb3f9JPF3yhpVe7b9Xap1Vpqi8L6MqNP25laHxzVE2XvfuPAbu4GOBJ6OpBaHJ71TWam0AI6+Z889uqHUfOvOXPYGtczJ6yA7GfQFBOUfs4tdso36xp6isNdWVzIpopZA+LDmO4tBGW/YDhnHEr3uSoN3R13/rI/mmLucqL+Lqm/rOH6kiCrNjmjr5tCtVwsDbibdpxlS2eukiGZaiUsAZGB0boA3uPDOOnhiT0nJpvNs1vRyRXOmmsEE8dQZ3uLajda4O3NwDG9kY3gcY4+heryT/tFqD55F+bV7oMp8o263y67QHWGV1QKGFkXuKlaDuTOe0Zfu/dOLiW+jGO1XDFsB0KdPMtps+5UGENNaJHc+H44v3s4znqxj0L3tdbQ9K6DbDPfqqMVJBdBBHHzk7h1lreoekkD0qFU+3a+6kj39IbO7tc4ScNqJpBHH6wCP8AyQUtsovVfoPafR0TJ3Bk1d4NrIwcMlBeWZI7Q7BHZ9K2WOjisP2N88+1OhlrIfc9S++sdNDnPNvNRlzc9eDkfQtwICIiAiIgdCIiAiIgIiICIiAiIgIiICIiAiIgIiIHUiIgIiICIiAiIgLPHKQLRrzRzi4Ybgn0f5wxaAr4JaqhqIIKh1NNLE5jJmjJjcQQHD5Dx+hUfduTXeL9UsqrrtAqq+oY0NEtRSl7gBx4EycOPHggvYOHH5VFdpehabaBpSqtM242oA52kmcP91MB8E/IeIPoJXx2d6JvOjI6yK66srtQMm3OZFSD/m+7nOCXOPHI9SmKD+f1XT1trqKu3VbJaaSOTmqiBxxh7HHgR14OcLbOzAh2zrTWOjwbT/UChm03YBS6/wBQi9Ul2FqmkjDKlvufnBM4cA/7IYOOB7cBWPpey+9vTdss3P8AP+4aaOn53d3d/daBnGTjOEFAcq6CRl809UFp5t9NNGHdWQ9px6iFYPJ0vtDXbN6G2xVMRrLe+WOaHeG80GRzmnHTghw4/KpVtB2fWvaJZPBtxc+GSN3O09TGAXwvxjIz0gjgR1+oqlI+SneW1vxpomU4OBIymkEmPk3sf4oPX5UOqbRUWS32GnroJrkytFRJBG7eMbAx7cux0HLhwPFd3kqAjSV5LvOI/NMXbufJvs50gyy2mrbDXuqo6iouVVFzkkoaHDdABG6PhZwOzjlfvQvJ/qdF3enuUetbmRHK2SSlpo+ZinA+5eN45B6EHb1RtgqtN7W7ZpOeKjbaalsXOzuB5xrpA4NOc4A3g3q6CVOtUaMsGs6I0l8ttPVsxhr3NxJGe1jxxafkKhO03YTQ7QLq68RXept9e6JsLsxiWJwbnHDgQePUVHoNi+0qKh8GeM2ZtBu7m4BKSG9n2WcejeQVFo+0uotstttlnqHTijvXNRTN+7iZId5xx/0B2fpUx5Ud0rZNY222SF4oYKETRMzwL3vcHO+XDQP/APVb2zXYvY9nT3VzJZLjdXs5s1kzQ3caelrGj7EHrOST2rs7TtlVs2l0UInnfRV9KCIKuNodgHpa5v3TevpBB6OtB1thUVoi2ZWd1p5n4UW9VFmN4z5O/v8Apz29WOpUNyirxbb1tEc621UNUKajjp5XRO3mtkDnktyOBI3hnCmFr5Ll4p6pzJ9WxQUbz/lBSwyB8g7MFwb68qQ6y5NFnudttsGmaplpqKJro3vmjMnuoOOd55BB3s549GDjAwEFibNHA7PNNH/4yn/NhZ/5UA3toFD/AFZH+ckVrbO9j9z0TDM6bWFwqpH0r6aGBgLaanLuh7WOcfhAjI6OtRW8cmm76gqxV3fX9TcKndDDNUUZe7A6v95wHE8PSgsjY9jxY6aH8wj/AEqueVc7/QWnwOOayX80pps02Y3jZ/UFk2r6q6WwQGKKgfEWRxOLgd5uXHHQRgdqjur9gl71ldJ6q469qpqZ08ktPSzUxeyma48Gt+GBwGBnHHCD78mABmzmf+sp/qxqAcqkP9+FlcGncNvcAfSJTn8oVgaD2IXrQd3pqmk11VvtzJudnt7KcsjqOGMEb5A6uOOpSbafsst+0q2U8M9S+jrqQudT1TGb27vY3mubwy04HWCCEHV2E3Cmq9llkbDMx7qaJ0EwB4xva92Qezhg/IVH+UdqF3i4bHbaqOanq69lJUvgkDhgNc4sJHa5oBH0KK23ks3VlQW1uq4YqR32YpYH77x9LgPXlWlR7HNKUWiajR8dLKaCpdzssjpMyul4YlB6A4YGMDHDGEFaclSG0OZfZS2Lwu18bWl2OcEG79z143s5x2DPUu9yp71bve/a7OKuJ1w92ic07XZe2MRuG8R1DLgOPSvJk5LN3p68voNWQMgB+BIad7Jmj/tdjPyEKWN5NthGk6u2Prp57xUubIbrMzec1zTkANzwYeIIzk56eAQdPkqjd0bd88f9In80xdvlREeLqmP/AMnD9SRdHQvJ5u2lrmyrn1rVwwslZK6mtzXRCYtOQHkuII6sY6F6ev8AYhete3aoqKnXNWy3vm52C3vpy+Kn4YwBvgZ6eOOtB4fJPI8Bag+eRfmlfHUqR0/yer/pWUvsu0SroGve10rIKUtbLjtHOYPDI4jrV3Y4Y7UGK9a3Jl52tXGS/vcacXfmJ94/7unZLu7voAYPyrYpq7VZ7Q2pE9JS2yCMFsgc1kLIwOGD0Yx0YVXbTeT1Q60u018tNwFsuFRg1DJIy+KZwGN/gQWuwOOMg9mV4umOTE6CeH3zahdXUMLt4UNK17GP9BLjwH9EA+kIKctdXBctrFJWU0gdBPf2yxvxjea6pyD9IIW3VRGsOTM27akmudhvFNa6Wd4k9zupyeYdgZ3N0jhwyBwwrR0FpCp0ZZ30NXf7jfJpJeddPWOyWnAG63JJDeHQSUEmREQEREBERARE6EBERAREQEREBERAREQEREBERAREQEREBERAREQEREBOtEQECIgIiICIiAidKICIiAiIgIiIGUREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBETCAiIgIiICIiAiIgIiICIiAiIgIiICIiAEREBERAREQEREBERAREQEREBERAREQEREBE606kBERAROnqRARE4ICImQgIiICIiAmETKAmERAREQEREBERAROhEBERAREygIiICJlEBETKAiIgIiICJlMoCInQgIiICIiAiIgIiICImUBEymUBETKAiIgIiICIiAiJlAREQETKIP/9k=";

/* Indonesian number-to-words (terbilang), verified against reference invoice */
const TERBILANG_WORDS = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"];
function terbilang(n) {
  n = Math.floor(Math.abs(n));
  if (n < 12) return TERBILANG_WORDS[n];
  if (n < 20) return terbilang(n - 10) + " Belas";
  if (n < 100) return terbilang(Math.floor(n / 10)) + " Puluh" + (n % 10 !== 0 ? " " + terbilang(n % 10) : "");
  if (n < 200) return "Seratus" + (n - 100 !== 0 ? " " + terbilang(n - 100) : "");
  if (n < 1000) return terbilang(Math.floor(n / 100)) + " Ratus" + (n % 100 !== 0 ? " " + terbilang(n % 100) : "");
  if (n < 2000) return "Seribu" + (n - 1000 !== 0 ? " " + terbilang(n - 1000) : "");
  if (n < 1000000) return terbilang(Math.floor(n / 1000)) + " Ribu" + (n % 1000 !== 0 ? " " + terbilang(n % 1000) : "");
  if (n < 1000000000) return terbilang(Math.floor(n / 1000000)) + " Juta" + (n % 1000000 !== 0 ? " " + terbilang(n % 1000000) : "");
  if (n < 1000000000000) return terbilang(Math.floor(n / 1000000000)) + " Milyar" + (n % 1000000000 !== 0 ? " " + terbilang(n % 1000000000) : "");
  return terbilang(Math.floor(n / 1000000000000)) + " Triliun" + (n % 1000000000000 !== 0 ? " " + terbilang(n % 1000000000000) : "");
}

function formatRupiah(n) {
  const rounded = Math.round(n || 0);
  return "Rp " + rounded.toLocaleString("id-ID") + ",-";
}
function formatUSD(n) {
  return "$" + Number(n || 0).toLocaleString("en-US", { maximumFractionDigits: 2 });
}
/* Kurs (exchange rate) must use Indonesian locale (dot thousands separator), not en-US commas */
function formatKursID(n) {
  return "Rp " + Math.round(n || 0).toLocaleString("id-ID") + ",-";
}
/* Qty in the reference invoice is shown with 2 decimals, Indonesian locale */
function formatQtyID(n) {
  return Number(n || 0).toLocaleString("id-ID", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/* Core invoice calculation, matches the reference template's tax formulas:
 *   DPP Nilai Lain = SubTotal x 11/12
 *   PPN (informational, DIBEBASKAN / exempt per PPN No. 49/2022) = DPP x 12%
 *   PPh 22 = SubTotal x 1.5%
 *   Total = SubTotal - PPh 22   (PPN is NOT added, since it's exempted)
 */
function computeInvoice(inputs) {
  const { qty, price, exchangeRate, paymentType, settlementQty, settlementRate } = inputs;
  const q = Number(qty) || 0;
  const p = Number(price) || 0;
  const fx = Number(exchangeRate) || 0;

  const totalHargaDP = q * p * fx; // full contract value at DP-basis qty/rate
  const dpPaidAmount = totalHargaDP * 0.5;

  let subTotal, totalHargaLine, dpLine;
  if (paymentType === "settlement") {
    const sq = Number(settlementQty) || 0;
    const sfx = Number(settlementRate) || 0;
    const totalHargaFinal = sq * p * sfx;
    subTotal = totalHargaFinal - dpPaidAmount;
    totalHargaLine = totalHargaFinal;
    dpLine = dpPaidAmount;
  } else {
    subTotal = dpPaidAmount;
    totalHargaLine = totalHargaDP;
    dpLine = dpPaidAmount;
  }

  const dpp = subTotal * (11 / 12);
  const ppn = dpp * 0.12;
  const pph22 = subTotal * 0.015;
  const total = subTotal - pph22;

  return { totalHargaLine, dpLine, subTotal, dpp, ppn, pph22, total, dpPaidAmount };
}


/* ============================================================
 * InvoiceModal — form + live preview matching PT Integra Mining
 * Nusantara's invoice template, with print-to-PDF support.
 * ============================================================ */
function InvoiceModal({ barge, onClose }) {
  const todayISO = new Date().toISOString().slice(0, 10);
  const [invoiceNo, setInvoiceNo] = useState("");
  const [invoiceDate, setInvoiceDate] = useState(todayISO);
  const [siNo, setSiNo] = useState("");
  const [kontrakNo, setKontrakNo] = useState("");
  const [kontrakDate, setKontrakDate] = useState(todayISO);
  const [clientName, setClientName] = useState("PT ATLAS DELTA MINERALS");
  const [clientAddress, setClientAddress] = useState("CoHive 101\nJl Mega Kuningan Barat Blok E.4.7 No. 1\nKel. Kuningan Timur, Kec. Setiabudi\nKota Jakarta Selatan, Provinsi DKI Jakarta");
  const [bargeName, setBargeName] = useState(barge.bargeName || "");
  const [tugboatName, setTugboatName] = useState(barge.tugboatName || "");
  const [qty, setQty] = useState(barge.totalWMT || 0);
  const [price, setPrice] = useState("");
  const [exchangeRate, setExchangeRate] = useState("");
  const [paymentType, setPaymentType] = useState("dp"); // 'dp' | 'settlement'
  const [settlementQty, setSettlementQty] = useState(barge.totalWMT || 0);
  const [settlementRate, setSettlementRate] = useState("");

  const calc = useMemo(() => computeInvoice({
    qty, price, exchangeRate, paymentType, settlementQty, settlementRate
  }), [qty, price, exchangeRate, paymentType, settlementQty, settlementRate]);

  const readyToPreview = Number(price) > 0 && Number(exchangeRate) > 0 &&
    (paymentType !== "settlement" || Number(settlementRate) > 0);

  const handlePrint = () => printFitToPage("invoice-print-area");
  const paymentLabel = paymentType === "settlement" ? "Pembayaran Tahap Kedua - Pelunasan" : "Pembayaran Tahap Pertama - 50%";
  const displayQty = paymentType === "settlement" ? settlementQty : qty;
  const displayRate = paymentType === "settlement" ? settlementRate : exchangeRate;

  return (
    <div className="invoice-modal">
      <div className="invoice-panel glass">
        <div className="invoice-head no-print">
          <FileText size={20} style={{ color: "#E35F0C" }} />
          <span>Generate Invoice — Barge #{barge.no}</span>
          <button className="invoice-close" onClick={onClose}><X size={18} /></button>
        </div>

        <div className="invoice-body">
          <div className="invoice-form no-print">
            <div className="form-section-title">Client (Kepada)</div>
            <div className="form-grid">
              <label>Client Name
                <input value={clientName} onChange={(e) => setClientName(e.target.value)} placeholder="PT ATLAS DELTA MINERALS" />
              </label>
              <label>Client Address (one line each)
                <textarea rows={4} value={clientAddress} onChange={(e) => setClientAddress(e.target.value)}
                  placeholder={"CoHive 101\nJl Mega Kuningan Barat Blok E.4.7 No. 1\nKel. Kuningan Timur, Kec. Setiabudi\nKota Jakarta Selatan"} />
              </label>
            </div>

            <div className="form-section-title">Invoice details</div>
            <div className="form-grid">
              <label>No Invoice
                <input value={invoiceNo} onChange={(e) => setInvoiceNo(e.target.value)} placeholder="014/INV/IMN-ADM/VII/2026" />
              </label>
              <label>Tanggal
                <input type="date" value={invoiceDate} onChange={(e) => setInvoiceDate(e.target.value)} />
              </label>
              <label>SI No
                <input value={siNo} onChange={(e) => setSiNo(e.target.value)} placeholder="09/SI-IMN/VII/2026" />
              </label>
              <label>Kontrak No
                <input value={kontrakNo} onChange={(e) => setKontrakNo(e.target.value)} placeholder="073/PKS/IMN-ADM/VII/2026" />
              </label>
              <label>Tanggal Kontrak
                <input type="date" value={kontrakDate} onChange={(e) => setKontrakDate(e.target.value)} />
              </label>
              <label>Barge Name
                <input value={bargeName} onChange={(e) => setBargeName(e.target.value)} />
              </label>
              <label>Tugboat Name
                <input value={tugboatName} onChange={(e) => setTugboatName(e.target.value)} />
              </label>
            </div>

            <div className="form-section-title">Pricing</div>
            <div className="form-grid">
              <label>Qty (WMT) — from finalized barge
                <input type="number" value={qty} onChange={(e) => setQty(e.target.value)} />
              </label>
              <label>Nickel Price (USD / WMT)
                <input type="number" step="0.01" value={price} onChange={(e) => setPrice(e.target.value)} placeholder="e.g. 41" />
              </label>
              <label>Exchange Rate (Rp / USD)
                <input type="number" step="1" value={exchangeRate} onChange={(e) => setExchangeRate(e.target.value)} placeholder="e.g. 18041" />
              </label>
            </div>

            <div className="form-section-title">Payment type</div>
            <div className="radio-row">
              <label className={`radio-pill ${paymentType === "dp" ? "radio-pill-on" : ""}`}>
                <input type="radio" name="paymentType" checked={paymentType === "dp"} onChange={() => setPaymentType("dp")} />
                Down Payment (50%)
              </label>
              <label className={`radio-pill ${paymentType === "settlement" ? "radio-pill-on" : ""}`}>
                <input type="radio" name="paymentType" checked={paymentType === "settlement"} onChange={() => setPaymentType("settlement")} />
                Settlement (Pelunasan)
              </label>
            </div>

            {paymentType === "settlement" && (
              <>
                <div className="form-section-title">Settlement adjustment</div>
                <p className="form-hint">
                  The 50% down payment already invoiced is recalculated from the Qty / Exchange Rate above.
                  Adjust the final measured Qty and settlement-date exchange rate below — the remaining
                  balance (Total Harga at final Qty/Rate minus the DP already paid) becomes this invoice's Sub Total.
                </p>
                <div className="form-grid">
                  <label>Final Qty (WMT)
                    <input type="number" value={settlementQty} onChange={(e) => setSettlementQty(e.target.value)} />
                  </label>
                  <label>Settlement Exchange Rate (Rp / USD)
                    <input type="number" step="1" value={settlementRate} onChange={(e) => setSettlementRate(e.target.value)} placeholder="e.g. 18100" />
                  </label>
                </div>
                <div className="settlement-note">
                  DP already invoiced: <strong>{formatRupiah(calc.dpPaidAmount)}</strong> (50% of {formatUSD(price)} × {fmt(qty)} WMT × Rp {fmt(exchangeRate)})
                </div>
              </>
            )}
          </div>

          <div className="invoice-preview-wrap">
            {!readyToPreview && (
              <div className="invoice-placeholder no-print">
                Enter Nickel Price and Exchange Rate {paymentType === "settlement" ? "(and Settlement Exchange Rate) " : ""}
                to generate the invoice preview.
              </div>
            )}
            {readyToPreview && (
              <div id="invoice-print-area" className="invoice-sheet print-area">
                <div className="inv-header">
                  <img src={INVOICE_LOGO_HEADER} alt="" className="inv-logo" />
                </div>

                <table className="inv-top-table">
                  <colgroup>
                    <col style={{ width: "237.8pt" }} />
                    <col style={{ width: "136.5pt" }} />
                    <col style={{ width: "147.8pt" }} />
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className="inv-kepada" rowSpan={2}>
                        <div>Kepada :</div>
                        <div className="inv-spacer" />
                        <div className="inv-bold">{clientName || "—"}</div>
                        {clientAddress.split("\n").filter(Boolean).map((line, i) => (
                          <div key={i}>{line}</div>
                        ))}
                      </td>
                      <td className="inv-cell-center">
                        <div className="inv-bold">No Invoice</div>
                        <div className="inv-spacer" />
                        <div>{invoiceNo || "—"}</div>
                      </td>
                      <td className="inv-cell-center">
                        <div className="inv-bold">Tanggal</div>
                        <div className="inv-spacer" />
                        <div>{fmtDateID(invoiceDate)}</div>
                      </td>
                    </tr>
                    <tr>
                      <td className="inv-cell-center">
                        <div className="inv-bold">SI No:</div>
                        <div>{siNo || "—"}</div>
                        <div className="inv-spacer" />
                        <div>{tugboatName || "—"}</div>
                        <div>{bargeName || "—"}</div>
                      </td>
                      <td className="inv-cell-center">
                        <div className="inv-bold">Kontrak</div>
                        <div className="inv-spacer" />
                        <div>{kontrakNo || "—"}</div>
                        <div>{fmtDateID(kontrakDate)}</div>
                      </td>
                    </tr>
                  </tbody>
                </table>

                <table className="inv-items-table">
                  <colgroup>
                    <col style={{ width: "54pt" }} />
                    <col style={{ width: "132pt" }} />
                    <col style={{ width: "63pt" }} />
                    <col style={{ width: "56.2pt" }} />
                    <col style={{ width: "74.2pt" }} />
                    <col style={{ width: "142.5pt" }} />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Kode</th>
                      <th>Keterangan</th>
                      <th>Qty</th>
                      <th>Harga</th>
                      <th>Kurs</th>
                      <th>Total Harga</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td></td>
                      <td className="inv-left">
                        Nickel Ore Indonesia<br />
                        <div className="inv-spacer" />
                        {paymentLabel}
                      </td>
                      <td className="inv-center">{formatQtyID(displayQty)}</td>
                      <td className="inv-right">$ {Number(price || 0).toLocaleString("en-US")}</td>
                      <td className="inv-right">{formatKursID(displayRate)}</td>
                      <td className="inv-right">
                        {formatRupiah(calc.totalHargaLine)}<br />
                        <div className="inv-spacer" />
                        {formatRupiah(paymentType === "settlement" ? calc.subTotal : calc.dpLine)}
                      </td>
                    </tr>
                    <tr>
                      <td colSpan={3} className="inv-bank-cell" rowSpan={5}>
                        Harap Pembayaran Melalui Rekening Sebagai Berikut :<br />
                        NAMA : PT Integra Mining Nusantara<br />
                        BANK : Mandiri<br />
                        CABANG : KCP Jakarta Gedung Patra Jasa<br />
                        NO REKENING : 126-00-5886887-7
                      </td>
                      <td colSpan={2} className="inv-left">Sub Total</td>
                      <td className="inv-right">{formatRupiah(calc.subTotal)}</td>
                    </tr>
                    <tr>
                      <td colSpan={2} className="inv-left">DPP Nilai Lainnya</td>
                      <td className="inv-right">{formatRupiah(calc.dpp)}</td>
                    </tr>
                    <tr>
                      <td colSpan={2} className="inv-left">PPN DIBEBASKAN -<br />PP Nomor 49 Tahun 2022</td>
                      <td className="inv-right">{formatRupiah(calc.ppn)}</td>
                    </tr>
                    <tr>
                      <td colSpan={2} className="inv-left">PPh 22 - 1.5%</td>
                      <td className="inv-right">{formatRupiah(calc.pph22)}</td>
                    </tr>
                    <tr>
                      <td colSpan={2} className="inv-left inv-bold">Total</td>
                      <td className="inv-right inv-bold">{formatRupiah(calc.total)}</td>
                    </tr>
                    <tr>
                      <td colSpan={6} className="inv-terbilang">
                        <div className="inv-bold">Terbilang :</div>
                        # {terbilang(Math.round(calc.total))} Rupiah
                      </td>
                    </tr>
                  </tbody>
                </table>

                <table className="inv-sig-table">
                  <tbody>
                    <tr>
                      <td className="inv-sig-box">
                        <div className="inv-bold">PT INTEGRA MINING NUSANTARA</div>
                        <div className="inv-sig-gap" />
                        <div className="inv-bold inv-underline">HARIS FAUZAN SALEH</div>
                        <div className="inv-bold">DIREKTUR UTAMA</div>
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div className="inv-footer">
                  <div className="inv-bold">PT INTEGRA MINING NUSANTARA</div>
                  <div>Patra Jasa Office Lt. 17, Room 1702-1704</div>
                  <div>Kel. Kuningan Timur, Kec. Setiabudi</div>
                  <div>Jakarta Selatan, Indonesia</div>
                </div>
              </div>
            )}
          </div>
        </div>

        {readyToPreview && (
          <div className="invoice-action-bar no-print">
            <button className="btn-primary invoice-print-btn" onClick={handlePrint}>
              <Printer size={14} /> Print / Save as PDF
            </button>
            <button className="btn-download-html" onClick={() => downloadPrintAreaAsHTML("invoice-print-area", `Invoice_${invoiceNo || "Barge_" + barge.no}.html`)}>
              <FileDown size={14} /> Download HTML
            </button>
            <div className="invoice-action-hint">If Print doesn't open a dialog (common in embedded previews), use Download HTML, then open that file in your browser and print from there.</div>
          </div>
        )}
      </div>
    </div>
  );
}

function fmtDateID(iso) {
  if (!iso) return "—";
  const months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return iso;
  return `${d} ${months[m - 1]} ${y}`;
}

/* Downloads the given print-area element as a standalone .html file (all CSS inlined).
 * This is the reliable fallback to "Print / Save as PDF": window.print() can be blocked
 * or behave unpredictably inside a sandboxed preview iframe (which is how this dashboard
 * runs when embedded), whereas a Blob-based file download works regardless. Opening the
 * downloaded file in a normal (non-sandboxed) browser tab and printing from there to PDF
 * always works, since it's then a regular top-level page rather than embedded content. */
function downloadPrintAreaAsHTML(elementId, filename) {
  const el = document.getElementById(elementId);
  if (!el) return;
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${filename}</title><style>${CSS}</style></head>` +
    `<body style="background:#e5e5e5;padding:24px;">${el.outerHTML}</body></html>`;
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* Guarantees single-page print output by MEASURING the report's actual rendered height
 * and scaling it down to fit, instead of relying on hand-tuned padding/font sizes that
 * can't account for every browser's print-DPI quirks. This is the definitive fix for
 * content spilling a few points onto a near-blank second page — rather than guessing at
 * more spacing trims, we directly measure and correct for whatever the real height is.
 * Only shrinks when content would genuinely overflow one page; short barges/invoices
 * print at 100% same as before. */
function printFitToPage(elementId) {
  const el = document.getElementById(elementId);
  if (!el) { window.print(); return; }

  // Reset any leftover scale from a previous print before measuring, so repeated
  // clicks don't compound a stale scale factor into the new measurement.
  el.style.removeProperty("transform");
  el.style.removeProperty("width");
  el.style.removeProperty("transform-origin");
  void el.offsetHeight; // force a reflow so scrollHeight reflects the reset state

  // A4 usable height after the 8mm @page margin, in the ~96dpi CSS-px reference browsers
  // use for print layout. Deliberately conservative (a bit under the true ~1060px) so
  // small cross-browser DPI differences don't tip a "just fits" report back over a page.
  const USABLE_HEIGHT_PX = 980;
  const naturalHeight = el.scrollHeight;

  if (naturalHeight > USABLE_HEIGHT_PX) {
    const scale = USABLE_HEIGHT_PX / naturalHeight;
    // setProperty(..., "important") is required here because the print stylesheet sets
    // width:100% with !important — a plain inline style wouldn't be able to win against
    // that, which would leave the compensating width from taking effect.
    el.style.setProperty("transform", `scale(${scale})`, "important");
    el.style.setProperty("transform-origin", "top left", "important");
    el.style.setProperty("width", `${100 / scale}%`, "important");
  }

  const reset = () => {
    el.style.removeProperty("transform");
    el.style.removeProperty("width");
    el.style.removeProperty("transform-origin");
    window.removeEventListener("afterprint", reset);
  };
  window.addEventListener("afterprint", reset);
  // Safari doesn't always fire `afterprint` reliably after a print-to-PDF flow, so also
  // reset on a fallback timer rather than leaving the on-screen preview shrunk.
  setTimeout(reset, 3000);

  window.print();
}

/* ============================================================
 * BargeExportModal — print-ready export for a SINGLE barge: its
 * header info (no, date, barge/tugboat name, qty, avg Ni, Draft/
 * Final status) plus the full list of domes that make it up.
 * "Print / Save as PDF" uses the browser's print dialog against
 * the shared .print-area mechanism.
 * ============================================================ */
/* ============================================================
 * BargeStatusModal — proactive, read-only diagnostic for a barge.
 * Two independent checks:
 *  1. Stock deficit — would finalizing this barge (or did finalizing it,
 *     if already finalized) require any dome to drop below zero?
 *  2. Lab data completeness — does any source dome have a 0% value for
 *     Ni, Fe, Co, SiO2, MgO, Al2O3, or Si:Mg, meaning it hasn't been lab
 *     assayed yet and the stock database needs updating.
 * Purely informational — no action is taken here. Finalizing itself no
 * longer blocks on this; use this modal beforehand to review first.
 * ============================================================ */
/* ============================================================
 * LoginScreen — full-screen gate shown before isLoggedIn is true.
 * ============================================================ */
/* Shared brand mark — same geometric logo used in the topbar, extracted so the login
 * and welcome screens use the actual brand identity instead of a generic placeholder. */
function BrandMark({ size = 40 }) {
  return (
    <svg width={size} height={size * (46 / 40)} viewBox="0 0 700 820" className="brand-logo">
      <rect x="62" y="57" width="185" height="62" fill="#E35F0C" />
      <rect x="62" y="167" width="185" height="63" fill="#E35F0C" />
      <rect x="62" y="278" width="587" height="62" fill="#BFB12A" />
      <rect x="62" y="278" width="66" height="299" fill="#BFB12A" />
      <rect x="174" y="387" width="475" height="62" fill="#BFB12A" />
      <rect x="174" y="387" width="73" height="190" fill="#BFB12A" />
      <rect x="389" y="501" width="59" height="185" fill="#E35F0C" />
      <rect x="62" y="624" width="386" height="62" fill="#E35F0C" />
      <rect x="495" y="501" width="65" height="294" fill="#E35F0C" />
      <rect x="62" y="734" width="498" height="61" fill="#E35F0C" />
    </svg>
  );
}

function LoginScreen({ onLogin, error }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");

  const handleSubmit = (e) => {
    e.preventDefault();
    onLogin(username, password);
  };

  return (
    <div className="login-container">
      <div className="bg-glow bg-glow-a" />
      <div className="bg-glow bg-glow-b" />
      <div className="login-box">
        <div className="login-panel-brand">
          <div className="login-panel-brand-pattern" />
          <div className="login-panel-brand-content">
            <BrandMark size={48} />
            <div className="login-panel-brand-name">INTEGRA</div>
            <div className="login-panel-brand-tag">Nickel Ore Barging<br />Management System</div>
          </div>
          <div className="login-panel-brand-foot">PT Integra Mining Nusantara</div>
        </div>

        <div className="login-panel-form">
          <div className="login-form-head">
            <div className="login-form-title">Sign in</div>
            <div className="login-form-sub">Enter your credentials to access the dashboard</div>
          </div>

          <form onSubmit={handleSubmit} className="login-form">
            <div className="form-group">
              <label htmlFor="username">Username</label>
              <input id="username" type="text" placeholder="Enter your username" value={username}
                onChange={(e) => setUsername(e.target.value)} autoComplete="username" className="login-input" />
            </div>
            <div className="form-group">
              <label htmlFor="password">Password</label>
              <input id="password" type="password" placeholder="Enter your password" value={password}
                onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" className="login-input" />
            </div>
            {error && <div className="login-error">{error}</div>}
            <button type="submit" className="login-button">Sign In</button>
          </form>

          <div className="login-footer">
            <p>PT Integra Mining Nusantara — 2026</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
 * WelcomeScreen — brief full-screen transition shown for 3s after login.
 * ============================================================ */
function WelcomeScreen({ username, role }) {
  return (
    <div className="welcome-overlay">
      <div className="bg-glow bg-glow-a" />
      <div className="bg-glow bg-glow-b" />
      <div className="welcome-content">
        <div className="welcome-logo"><BrandMark size={64} /></div>
        <h1 className="welcome-title">Welcome to Integra OS</h1>
        <p className="welcome-subtitle">Signed in as <strong>{username}</strong> ({role})</p>
        <div className="welcome-spinner"></div>
      </div>
    </div>
  );
}

/* ============================================================
 * SyncingScreen — shown on first login before the initial Sheets sync
 * completes, since there's no hardcoded fallback data anymore. Prevents
 * ever showing an empty or stale-looking dashboard while data loads.
 * ============================================================ */
function SyncingScreen({ syncFailed, syncError }) {
  return (
    <div className="welcome-overlay">
      <div className="bg-glow bg-glow-a" />
      <div className="bg-glow bg-glow-b" />
      <div className="welcome-content">
        <div className="welcome-logo"><BrandMark size={64} /></div>
        {syncFailed ? (
          <>
            <h1 className="welcome-title" style={{ color: "#F87171" }}>Couldn't load data</h1>
            <p className="welcome-subtitle">Sheets sync failed on first load{syncError ? `: ${syncError}` : ""}. Check your connection and reload, or contact an admin.</p>
          </>
        ) : (
          <>
            <h1 className="welcome-title">Loading data…</h1>
            <p className="welcome-subtitle">Pulling the latest Domes, Barges, HPM, and Exchange Rates from Google Sheets.</p>
            <div className="welcome-spinner"></div>
          </>
        )}
      </div>
    </div>
  );
}

/* ============================================================
 * Loading Report Assistant — dev-account-only for now. Paste a loading
 * report chat, regex-parses it (no API), review/edit the extracted
 * fields, then confirm to update the barge + push to Sheets.
 * ============================================================ */
function LoadingReportModal({ onClose, onSubmit }) {
  const [text, setText] = useState("");

  const handleParse = () => {
    if (!text.trim()) { alert("⚠️ Please paste a loading report"); return; }
    const parsed = parseLoadingReportText(text);
    if (parsed.error) { alert("❌ Parse error: " + parsed.error); return; }
    onSubmit(parsed);
    setText("");
  };

  return (
    <div className="validation-modal">
      <div className="validation-panel glass" style={{ maxWidth: 620 }}>
        <div className="validation-head" style={{ color: "#22D3B8", background: "linear-gradient(135deg, rgba(34,211,184,.08), rgba(255,255,255,.01))", borderBottom: "1px solid rgba(34,211,184,.2)" }}>
          <MessageSquare size={20} style={{ color: "#22D3B8" }} />
          <span>Parse Shipment Loading Report</span>
          <button className="validation-close" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="validation-body">
          <p>Paste the loading report from the operation team's chat. This extracts Shipment #, Qty On Board, Progress %, Balance Due, and Date via pattern matching — instant, no API call.</p>
          <div className="form-group">
            <label>Loading Report Chat</label>
            <textarea value={text} onChange={(e) => setText(e.target.value)} rows={8} className="login-input"
              style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: "11.5px", resize: "vertical", width: "100%" }}
              placeholder="Paste chat report here... (e.g. Shipment 09, Qty on board 1.800 wmt, Progress 17%, Blc: 8.700 wmt, Tgl. 02-08-2026)" />
          </div>
          <div className="validation-actions">
            <button className="btn-ghost" onClick={onClose}>Cancel</button>
            <button className="btn-primary btn-sync-sheets" onClick={handleParse} disabled={!text.trim()}>Parse Report</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function LoadingReportReviewModal({ data, onClose, onConfirm, isSubmitting }) {
  const [edited, setEdited] = useState(data);

  const handleConfirm = () => {
    if (!edited.shipmentNumber) { alert("⚠️ Shipment number required"); return; }
    onConfirm(edited);
  };

  const field = (label, key, type = "number") => (
    <div className="form-group">
      <label>{label}</label>
      <input type={type} className="login-input" value={edited[key] ?? ""}
        onChange={(e) => setEdited((prev) => ({ ...prev, [key]: type === "number" ? (parseFloat(e.target.value) || null) : e.target.value }))} />
    </div>
  );

  return (
    <div className="validation-modal">
      <div className="validation-panel glass" style={{ maxWidth: 560 }}>
        <div className="validation-head" style={{ color: "#4ADE80", background: "linear-gradient(135deg, rgba(74,222,128,.08), rgba(255,255,255,.01))", borderBottom: "1px solid rgba(74,222,128,.2)" }}>
          <CheckCircle2 size={20} style={{ color: "#4ADE80" }} />
          <span>Review Parsed Data</span>
          <button className="validation-close" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="validation-body">
          <p>Review the extracted data below. Edit anything that looks off, then confirm to save.</p>
          {data.missingFields && (
            <div className="lab-list" style={{ marginBottom: 16 }}>
              <div className="lab-item">
                <span className="lab-dome">Not found in the text — fill in manually:</span>
                <span className="lab-detail">{data.missingFields.join(", ")}</span>
              </div>
            </div>
          )}
          <div className="review-grid">
            {field("Shipment Number (Barge #)", "shipmentNumber")}
            {field("Qty On Board (WMT)", "qtyOnBoard")}
            {field("Progress (%)", "progressPercent")}
            {field("Balance Due (WMT)", "balanceDue")}
            {field("Report Date", "reportDate", "date")}
          </div>
          <div className="validation-actions">
            <button className="btn-ghost" onClick={onClose} disabled={isSubmitting}>Cancel</button>
            <button className="btn-primary btn-sync-sheets" onClick={handleConfirm} disabled={isSubmitting}>
              {isSubmitting ? "Saving…" : "Confirm & Save"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function BargeStatusModal({ barge, domes, onClose }) {
  const domesById = useMemo(() => { const m = {}; domes.forEach((d) => (m[d.id] = d)); return m; }, [domes]);

  const stockIssues = useMemo(() => {
    if (barge.finalized) {
      // Already finalized — show what deficit adjustment (if any) was applied at the time.
      const adj = barge.stockAdjustments || {};
      return Object.entries(adj).map(([domeId, deficit]) => ({ domeId, deficit, historical: true }));
    }
    // Still a draft — check whether finalizing right now would push any dome negative.
    const issues = [];
    barge.sources.forEach((s) => {
      const d = domesById[s.id];
      if (d && s.amt > d.stock) {
        issues.push({ domeId: s.id, current: d.stock, requested: s.amt, deficit: s.amt - d.stock, historical: false });
      }
    });
    return issues;
  }, [barge, domesById]);

  const labIssues = useMemo(() => {
    const issues = [];
    barge.sources.forEach((s) => {
      const d = domesById[s.id];
      if (!d) return;
      const zeroFields = LAB_FIELDS.filter(([key]) => (d[key] || 0) === 0).map(([, label]) => label);
      if (zeroFields.length) issues.push({ domeId: s.id, fields: zeroFields });
    });
    return issues;
  }, [barge, domesById]);

  const noIssues = stockIssues.length === 0 && labIssues.length === 0;

  return (
    <div className="validation-modal">
      <div className="validation-panel glass">
        <div className="validation-head">
          <AlertTriangle size={20} style={{ color: noIssues ? "#4ADE80" : "#F87171" }} />
          <span>Barge #{barge.no} Status Check</span>
          <button className="validation-close" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="validation-body">
          {noIssues && (
            <div className="status-ok">
              <CheckCircle2 size={16} /> No stock deficits and all source domes have lab data recorded.
            </div>
          )}

          <div className="status-section-title" style={{ color: "#F87171" }}>
            <AlertTriangle size={13} /> Stock deficit {barge.finalized ? "(applied at finalize)" : "(if finalized now)"}
          </div>
          {stockIssues.length === 0 && <p style={{ marginBottom: 12 }}>None.</p>}
          {stockIssues.length > 0 && (
            <div className="violations-list">
              {stockIssues.map((v) => (
                <div key={v.domeId} className="violation-item">
                  <span className="violation-dome">{v.domeId}</span>
                  <span className="violation-detail">
                    {v.historical
                      ? `This dome was set to 0 remaining when finalized — deficit of ${fmt(v.deficit)} WMT.`
                      : `Current: ${fmt(v.current)} WMT | Requested: ${fmt(v.requested)} WMT | Deficit: ${fmt(v.deficit)} WMT`}
                  </span>
                </div>
              ))}
            </div>
          )}

          <div className="status-section-title" style={{ color: "#FBBF24" }}>
            <AlertTriangle size={13} /> Missing lab data
          </div>
          {labIssues.length === 0 && <p style={{ marginBottom: 0 }}>None.</p>}
          {labIssues.length > 0 && (
            <div className="lab-list">
              {labIssues.map((v) => (
                <div key={v.domeId} className="lab-item">
                  <span className="lab-dome">{v.domeId}</span>
                  <span className="lab-detail">{v.fields.join(", ")} showing 0% — update the stock database with the lab result.</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function BargeExportModal({ barge, domesById, onClose }) {
  const todayISO = new Date().toISOString().slice(0, 10);
  const sources = [...(barge.sources || [])].sort((a, b) => b.amt - a.amt);
  const avgSimg = (() => {
    let weight = 0, sum = 0;
    sources.forEach((s) => {
      const d = domesById[s.id];
      const simg = d ? d.simg : undefined;
      if (simg !== undefined && simg > 0) { weight += s.amt; sum += s.amt * simg; }
    });
    return weight > 0 ? sum / weight : null;
  })();

  const handlePrint = () => printFitToPage("plan-print-area");

  return (
    <div className="invoice-modal">
      <div className="invoice-panel glass" style={{ width: 780 }}>
        <div className="invoice-head no-print">
          <FileDown size={20} style={{ color: "#E35F0C" }} />
          <span>Export Barge #{barge.no}</span>
          <button className="invoice-close" onClick={onClose}><X size={18} /></button>
        </div>

        <div className="invoice-body">
          <div className="invoice-preview-wrap" style={{ flex: 1, padding: "24px" }}>
            <div id="plan-print-area" className="print-area plan-sheet">
              <div className="plan-header">
                <img src={INVOICE_LOGO_HEADER} alt="" className="plan-logo" />
                <div className="plan-header-meta">
                  <div className="plan-title">Barge Plan — #{barge.no}</div>
                  <div className="plan-subtitle">Exported {fmtDateID(todayISO)}</div>
                </div>
              </div>

              <div className="plan-summary">
                <div className="plan-summary-item">
                  <span className="plan-summary-label">Date</span>
                  <span className="plan-summary-value">{barge.shipDate || "—"}</span>
                </div>
                <div className="plan-summary-item">
                  <span className="plan-summary-label">Barge Name</span>
                  <span className="plan-summary-value">{barge.bargeName || <span className="plan-muted">Not yet assigned</span>}</span>
                </div>
                <div className="plan-summary-item">
                  <span className="plan-summary-label">Tugboat Name</span>
                  <span className="plan-summary-value">{barge.tugboatName || <span className="plan-muted">Not yet assigned</span>}</span>
                </div>
                <div className="plan-summary-item">
                  <span className="plan-summary-label">Qty</span>
                  <span className="plan-summary-value">{fmt(barge.totalWMT)} WMT</span>
                </div>
                <div className="plan-summary-item">
                  <span className="plan-summary-label">Avg Ni</span>
                  <span className="plan-summary-value">{barge.totalWMT > 0 ? fmt(barge.grade, 2) : "—"}%</span>
                </div>
                {avgSimg !== null && (
                  <div className="plan-summary-item">
                    <span className="plan-summary-label">Avg Si/Mg</span>
                    <span className="plan-summary-value">{fmt(avgSimg, 2)}</span>
                  </div>
                )}
                <div className="plan-summary-item">
                  <span className="plan-summary-label">Status</span>
                  <span className={`plan-status-badge ${barge.finalized ? "plan-status-final" : "plan-status-draft"}`}>
                    {barge.finalized ? "Final" : "Draft"}
                  </span>
                </div>
              </div>

              <table className="plan-table">
                <thead>
                  <tr>
                    <th style={{ width: "34%" }}>Dome ID</th>
                    <th style={{ width: "22%" }}>Contractor</th>
                    <th style={{ width: "22%" }}>Qty (WMT)</th>
                    <th style={{ width: "22%" }}>Ni Grade %</th>
                  </tr>
                </thead>
                <tbody>
                  {sources.map((s) => (
                    <tr key={s.id}>
                      <td>{s.id}</td>
                      <td>{domesById[s.id]?.contractor || "—"}</td>
                      <td className="plan-right">{fmt(s.amt)}</td>
                      <td className="plan-right">{fmt(s.grade, 2)}</td>
                    </tr>
                  ))}
                  {sources.length === 0 && (
                    <tr><td colSpan={4} className="plan-center plan-muted">No domes assigned to this barge yet.</td></tr>
                  )}
                </tbody>
                {sources.length > 0 && (
                  <tfoot>
                    <tr>
                      <td colSpan={2} className="plan-right plan-footer-bold">Total</td>
                      <td className="plan-right plan-footer-bold">{fmt(barge.totalWMT)}</td>
                      <td className="plan-right plan-footer-bold">{fmt(barge.grade, 2)}</td>
                    </tr>
                  </tfoot>
                )}
              </table>

              <div className="plan-footer">
                <div className="plan-footer-bold">PT INTEGRA MINING NUSANTARA</div>
                <div>Patra Jasa Office Lt. 17, Room 1702-1704</div>
                <div>Kel. Kuningan Timur, Kec. Setiabudi</div>
                <div>Jakarta Selatan, Indonesia</div>
              </div>
            </div>
          </div>
        </div>

        <div className="invoice-action-bar no-print">
          <button className="btn-primary invoice-print-btn" onClick={handlePrint}>
            <Printer size={14} /> Print / Save as PDF
          </button>
          <button className="btn-download-html" onClick={() => downloadPrintAreaAsHTML("plan-print-area", `Barge_${barge.no}_Plan.html`)}>
            <FileDown size={14} /> Download HTML
          </button>
          <div className="invoice-action-hint">If Print doesn't open a dialog (common in embedded previews), use Download HTML, then open that file in your browser and print from there.</div>
        </div>
      </div>
    </div>
  );
}

export default function IntegraDashboard() {
  const [tab, setTab] = useState("overview");
  const [domes, setDomes] = useState(() => withInitialStock(DEFAULT_DOMES));
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [barges, setBarges] = useState([]);
  const [loaded, setLoaded] = useState(false);

  // Authentication & session
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [currentUser, setCurrentUser] = useState(null); // { username, role }
  const [loginError, setLoginError] = useState("");
  const [showWelcome, setShowWelcome] = useState(false);
  const [loginHistory, setLoginHistory] = useState([]);
  const isAdmin = currentUser?.role === "admin";
  const isOperation = currentUser?.role === "operation";

  // Financials (admin-only)
  const [hpmHistory, setHpmHistory] = useState(DEFAULT_HPM_HISTORY);
  const [lastSyncTime, setLastSyncTime] = useState(null);
  const [sheetsSyncStatus, setSheetsSyncStatus] = useState("Ready");
  const [lastSyncError, setLastSyncError] = useState("");
  const [exRateFetchStatus, setExRateFetchStatus] = useState("Not checked yet this session");

  // Feature flags & account management — configuration UI is dev-only, but the
  // resulting tab visibility applies to whoever is logged in (that's the actual
  // feature being tested). Defaults to everything visible — fail-open, not fail-closed
  // — so if the FeatureFlags sheet is unreachable or a user has no row in it yet,
  // nobody gets silently locked out of tabs they should have access to.
  const [userFeatures, setUserFeatures] = useState({
    stock: true, barging: true, timeline: true, financials: true, loginLog: true, settings: true, chatAssistant: true,
  });
  const [allUsers, setAllUsers] = useState([]);
  const [allFeatureFlags, setAllFeatureFlags] = useState({});
  const [userDetailedFeatures, setUserDetailedFeatures] = useState({});
  const [allDetailedFlags, setAllDetailedFlags] = useState({});
  const [exchangeRateHistory, setExchangeRateHistory] = useState(DEFAULT_EXCHANGE_RATE_HISTORY);

  const [statusBarge, setStatusBarge] = useState(null); // barge object currently having its status checked, or null
  const [pendingFinalize, setPendingFinalize] = useState(null); // { bargeNo, violations } awaiting confirm/cancel

  // Loading Report Assistant — dev account only while this is being tried out
  const isDevAccount = currentUser?.username === "dev";
  const [showLoadingReportModal, setShowLoadingReportModal] = useState(false);
  const [showReviewModal, setShowReviewModal] = useState(false);
  const [reviewData, setReviewData] = useState(null);
  const [isSubmittingLoadingReport, setIsSubmittingLoadingReport] = useState(false);
  const [invoiceBarge, setInvoiceBarge] = useState(null); // barge object currently being invoiced, or null
  const [exportBarge, setExportBarge] = useState(null); // barge object currently being exported, or null
  const domesByIdTop = useMemo(() => { const m = {}; domes.forEach((d) => (m[d.id] = d)); return m; }, [domes]);
  const [showImport, setShowImport] = useState(false);
  const [importDate, setImportDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [dataLastUpdated, setDataLastUpdated] = useState(DATA_LAST_UPDATED);
  const [importStatus, setImportStatus] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const d = await window.storage?.get("integra-domes-v3");
        const parsedDomes = d?.value ? JSON.parse(d.value) : null;
        setDomes(withInitialStock(parsedDomes && parsedDomes.length > 0 ? parsedDomes : DEFAULT_DOMES));
      } catch (e) { setDomes(withInitialStock(DEFAULT_DOMES)); }
      try { const s = await window.storage?.get("integra-settings-v3"); if (s?.value) setSettings(JSON.parse(s.value)); } catch (e) {}
      try {
        const b = await window.storage?.get("integra-barges-v3");
        const parsedBarges = b?.value ? JSON.parse(b.value) : null;
        setBarges(parsedBarges && parsedBarges.length > 0 ? parsedBarges : DEFAULT_BARGES);
      } catch (e) { setBarges(DEFAULT_BARGES); }
      setLoaded(true);
    })();
  }, []);
  useEffect(() => { if (loaded) window.storage?.set("integra-domes-v3", JSON.stringify(domes)).catch(() => {}); }, [domes, loaded]);
  useEffect(() => { if (loaded) window.storage?.set("integra-settings-v3", JSON.stringify(settings)).catch(() => {}); }, [settings, loaded]);
  useEffect(() => { if (loaded) window.storage?.set("integra-barges-v3", JSON.stringify(barges)).catch(() => {}); }, [barges, loaded]);

  // Session load on mount. Uses sessionStorage (not localStorage) specifically so the
  // "logout when browser closes" requirement actually holds — sessionStorage clears
  // automatically when the tab/browser closes, localStorage does not.
  useEffect(() => {
    try {
      const saved = sessionStorage.getItem("integraSession");
      if (saved) {
        const session = JSON.parse(saved);
        if (Date.now() - session.loginTime < SESSION_MAX_AGE_MS) {
          setCurrentUser(session.user);
          setIsLoggedIn(true);
        } else {
          sessionStorage.removeItem("integraSession");
        }
      }
    } catch (e) { sessionStorage.removeItem("integraSession"); }
    try {
      const savedHistory = localStorage.getItem("integraLoginHistory");
      if (savedHistory) setLoginHistory(JSON.parse(savedHistory));
    } catch (e) {}
  }, []);

  // Re-checks session age on an interval (not just once at mount) so the 1-hour cap
  // holds even across page reloads — a single mount-time setTimeout would reset every
  // time the page reloads, letting someone stay logged in indefinitely by refreshing.
  useEffect(() => {
    if (!isLoggedIn) return;
    const check = () => {
      try {
        const saved = sessionStorage.getItem("integraSession");
        if (!saved) return;
        const session = JSON.parse(saved);
        if (Date.now() - session.loginTime >= SESSION_MAX_AGE_MS) {
          sessionStorage.removeItem("integraSession");
          setCurrentUser(null);
          setIsLoggedIn(false);
          alert("Your session has expired. Please log in again.");
        }
      } catch (e) {}
    };
    const interval = setInterval(check, 60 * 1000);
    return () => clearInterval(interval);
  }, [isLoggedIn]);

  // Pull fresh data from Google Sheets on login (every user — this used to be
  // admin-only, which meant operation accounts never saw anything an admin pushed;
  // permanently stuck on stale/default data). Re-checks every 5 minutes while the tab
  // stays open, and immediately whenever the tab becomes visible again (switching back
  // from another app/tab) — there's no manual "Sync Now" button, so this is what keeps
  // everyone actually looking at the same data instead of whatever loaded at login.
  useEffect(() => {
    if (!isLoggedIn) return;
    syncWithSheets(false);
    const syncInterval = setInterval(() => syncWithSheets(false), 15 * 60 * 1000);
    const onVisible = () => { if (document.visibilityState === "visible") syncWithSheets(false); };
    document.addEventListener("visibilitychange", onVisible);
    return () => { clearInterval(syncInterval); document.removeEventListener("visibilitychange", onVisible); };
  }, [isLoggedIn]);

  const completeLogin = (username, role) => {
    const user = { username, role };
    const session = { user, loginTime: Date.now() };
    try { sessionStorage.setItem("integraSession", JSON.stringify(session)); } catch (e) {}
    setCurrentUser(user);
    setIsLoggedIn(true);
    setShowWelcome(true);
    loadUserFeaturesFromSheets(username);
    loadUserDetailedFeatures(username);
    if (username === "dev") { loadAllUsersAndFlags(); loadAllDetailedFlags(); }
    setTimeout(() => setShowWelcome(false), 3000);
  };

  // Hybrid login: the 3 core accounts (dev/integra/operation) authenticate instantly
  // against the hardcoded CREDENTIALS object, no network call at all — this is
  // deliberate. Making login itself depend on a live Sheets fetch would turn the
  // Sheets API into a single point of failure for getting into the app at all, and
  // we've hit real Sheets outages multiple times already this project (missing
  // package, malformed key, header mismatches). Only usernames NOT in the hardcoded
  // set fall through to a Sheets lookup — that's how accounts created via the dev
  // Account Management UI get to log in, without weakening the reliability of the
  // three accounts everyone actually depends on day to day.
  const handleLogin = async (username, password) => {
    setLoginError("");
    const logAttempt = (status) => {
      const entry = { timestamp: new Date().toISOString(), username: username || "(unknown)", status };
      setLoginHistory((prev) => {
        const updated = [entry, ...prev].slice(0, 100);
        try { localStorage.setItem("integraLoginHistory", JSON.stringify(updated)); } catch (e) {}
        writeLoginHistoryToSheets(updated);
        return updated;
      });
    };

    const cred = CREDENTIALS[username];
    if (cred) {
      if (cred.password !== password) {
        logAttempt("Failed");
        setLoginError("❌ Invalid username or password. Please try again.");
        return;
      }
      completeLogin(username, cred.role);
      logAttempt("Successful");
      return;
    }

    try {
      const response = await fetch("/api/sheets-read?sheetName=Users");
      if (!response.ok) { logAttempt("Failed"); setLoginError("❌ Invalid username or password."); return; }
      const { data } = await response.json();
      const user = data.find((u) => u["Username"] === username);
      if (!user || user["Password"] !== password) {
        logAttempt("Failed");
        setLoginError("❌ Invalid username or password. Please try again.");
        return;
      }
      if (user["Status"] === "disabled") {
        logAttempt("Failed");
        setLoginError("❌ This account is disabled.");
        return;
      }
      completeLogin(username, user["Role"]);
      logAttempt("Successful");
    } catch (error) {
      logAttempt("Failed");
      setLoginError("❌ Invalid username or password.");
    }
  };

  const handleLogout = () => {
    try { sessionStorage.removeItem("integraSession"); } catch (e) {}
    setCurrentUser(null);
    setIsLoggedIn(false);
    setLoginError("");
  };

  // Financials helpers. Fixed from the original build prompt, which referenced
  // barge.dateCreated/barge.id — neither field exists on this app's barge objects
  // (they're barge.shipDate and barge.no). As written, every barge would have silently
  // fallen back to today's price/rate instead of its own historical date.
  //
  // HPM is looked up directly from manually-entered history (no formula) — each barge
  // uses whatever HPM was in effect on its own shipDate, not today's.
  const getHpmOnDate = (targetDate) => {
    const sorted = [...hpmHistory].sort((a, b) => new Date(b.date) - new Date(a.date));
    const match = sorted.find((h) => new Date(h.date) <= new Date(targetDate));
    return match ? match.price : sorted[sorted.length - 1]?.price || 0;
  };
  const getExchangeRateOnDate = (targetDate) => {
    const sorted = [...exchangeRateHistory].sort((a, b) => new Date(b.date) - new Date(a.date));
    const match = sorted.find((h) => new Date(h.date) <= new Date(targetDate));
    return match ? match.rate : sorted[sorted.length - 1]?.rate || 17990;
  };

  const calculateRoyalty = (bargeDate, qty) => {
    const hpm = getHpmOnDate(bargeDate);
    const exRate = getExchangeRateOnDate(bargeDate);
    return hpm * exRate * qty * ROYALTY_TARIFF;
  };
  const getHpmTrendPercent = () => {
    if (hpmHistory.length < 2) return 0;
    return ((hpmHistory[0].price - hpmHistory[1].price) / hpmHistory[1].price) * 100;
  };
  const getExRateTrendPercent = () => {
    if (exchangeRateHistory.length < 2) return 0;
    return ((exchangeRateHistory[0].rate - exchangeRateHistory[1].rate) / exchangeRateHistory[1].rate) * 100;
  };

  // Auto-fetch today's USD/IDR rate once/day via the /api/fetch-exchange-rate serverless
  // function (see api/fetch-exchange-rate.js — a separate file, deployed alongside this
  // one on Vercel). Fails silently outside that deployment (e.g. in local preview),
  // which is intentional — manual "Update Rate" always remains available as a fallback.
  //
  // Two fixes vs. the original version: (1) this only ran once per login session before
  // — a tab left open across midnight would never re-check, since [isAdmin] only
  // re-fires on login/logout, not on a new day. Added a periodic recheck (same pattern
  // as the Sheets sync) so it actually re-evaluates "has the date changed?" hourly.
  // (2) errors were completely silent before — now surfaced via exRateFetchStatus so
  // it's visible in Settings rather than failing invisibly like the Sheets sync did.
  const runExRateAutoFetch = async () => {
    if (!isAdmin) return;
    try {
      const today = new Date().toISOString().split("T")[0];
      if (localStorage.getItem("lastExRateFetch") === today) return;
      const res = await fetch("/api/fetch-exchange-rate");
      if (!res.ok) throw new Error(await extractErrorDetail(res));
      const data = await res.json();
      if (!data.rate) throw new Error("Response missing rate field");
      const updated = [{ date: data.date, rate: data.rate, source: data.source }, ...exchangeRateHistory].slice(0, 180);
      setExchangeRateHistory(updated);
      writeExchangeRateToSheets(updated); // so the whole team sees the auto-fetched rate too, not just this browser
      localStorage.setItem("lastExRateFetch", today);
      setExRateFetchStatus(`✅ Auto-fetched ${data.date}`);
    } catch (e) {
      console.error("Exchange rate auto-fetch failed:", e);
      setExRateFetchStatus(`❌ Auto-fetch failed: ${e.message}`);
    }
  };
  useEffect(() => {
    if (!isAdmin) return;
    runExRateAutoFetch();
    const exRateInterval = setInterval(runExRateAutoFetch, 60 * 60 * 1000); // recheck hourly
    return () => clearInterval(exRateInterval);
  }, [isAdmin]);

  const exportFinancialData = () => {
    const csv = [
      "Financial Data Export — " + new Date().toISOString(), "",
      "HPM PRICE HISTORY", "Date,Price (USD/WMT)",
      ...hpmHistory.map((h) => `${h.date},${h.price}`), "",
      "EXCHANGE RATE HISTORY", "Date,Rate (IDR/USD),Source",
      ...exchangeRateHistory.map((e) => `${e.date},${e.rate},${e.source || "manual"}`),
    ].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `integra-financials-${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };


  // ---- Google Sheets export suite (Settings tab, admin only) ----
  // One CSV per category, all downloaded together, meant to be imported into a Google
  // Sheet with matching tab names (Domes / Barges / HPMHistory / ExchangeRates /
  // LoginHistory) as a Phase 2 manual backup/handoff step.
  const csvEscape = (cell) => {
    const str = String(cell ?? "");
    return str.includes(",") || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
  };
  const toCSV = (headers, rows) => [headers.join(","), ...rows.map((r) => r.map(csvEscape).join(","))].join("\n");
  // For the live Sheets write path specifically: build the row arrays directly rather
  // than encoding to a CSV string and re-parsing it back into arrays. That round-trip
  // was the actual bug — a genuinely empty cell (e.g. an unnamed draft barge's blank
  // Barge Name/Tugboat Name) has no non-comma character for a regex-based re-splitter
  // to match, so it silently vanished and shifted every later column left by one. Rows
  // built directly, never stringified, can't have this problem.
  const toRows = (headers, rows) => [headers, ...rows.map((r) => r.map((cell) => String(cell ?? "")))];

  const exportDomesCSV = () => toCSV(
    ["Dome ID", "Contractor", "Stock (WMT)", "Initial Stock (WMT)", "Ni %", "Fe %", "Co %", "SiO2 %", "MgO %", "Al2O3 %", "Si:Mg", "Location", "Source"],
    domes.map((d) => [d.id, d.contractor, d.stock, d.initialStock !== undefined ? d.initialStock : d.stock, fmt(d.ni, 2), fmt(d.fe, 2), fmt(d.co, 2), fmt(d.sio2, 2), fmt(d.mgo, 2), fmt(d.al2o3, 2), fmt(d.simg, 2), d.location || "", d.source])
  );
  const exportBargesCSV = () => toCSV(
    ["Barge No", "Ship Date", "Barge Name", "Tugboat Name", "Total WMT", "Grade (Ni %)", "Status", "Finalized", "Sources", "Stock Adjustments", "Qty On Board", "Progress %", "Balance Due", "Last Updated"],
    barges.map((b) => [
      String(b.no).padStart(2, "0"), b.shipDate, b.bargeName || "", b.tugboatName || "",
      fmt(b.totalWMT), fmt(b.grade, 2), b.status, b.finalized ? "Yes" : "No",
      (b.sources || []).map((s) => `${s.id}:${s.amt}WMT`).join("; "),
      b.stockAdjustments && Object.keys(b.stockAdjustments).length ? JSON.stringify(b.stockAdjustments) : "",
      b.qtyOnBoard || "", b.progressPercent || "", b.balanceDue || "", b.lastUpdated || "",
    ])
  );
  const exportHPMHistoryCSV = () => toCSV(
    ["Date", "Price (USD/WMT)"],
    [...hpmHistory].sort((a, b) => new Date(b.date) - new Date(a.date)).map((h) => [h.date, h.price.toFixed(2)])
  );
  const exportExchangeRateHistoryCSV = () => toCSV(
    ["Date", "Rate (IDR/USD)", "Source"],
    [...exchangeRateHistory].sort((a, b) => new Date(b.date) - new Date(a.date)).map((e) => [e.date, e.rate, e.source || "manual"])
  );
  const exportLoginHistoryCSV = () => toCSV(
    ["Timestamp", "Username", "Status"],
    loginHistory.map((l) => [l.timestamp, l.username, l.status])
  );

  const downloadCSV = (filename, csvContent) => {
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  };

  const exportAllForGoogleSheets = () => {
    const today = new Date().toISOString().split("T")[0];
    downloadCSV(`Integra-Domes-${today}.csv`, exportDomesCSV());
    setTimeout(() => downloadCSV(`Integra-Barges-${today}.csv`, exportBargesCSV()), 400);
    setTimeout(() => downloadCSV(`Integra-HPMHistory-${today}.csv`, exportHPMHistoryCSV()), 800);
    setTimeout(() => downloadCSV(`Integra-ExchangeRates-${today}.csv`, exportExchangeRateHistoryCSV()), 1200);
    setTimeout(() => downloadCSV(`Integra-LoginHistory-${today}.csv`, exportLoginHistoryCSV()), 1600);
    alert(
      "✅ Export complete! 5 CSV files downloaded:\n\n" +
      `1. Integra-Domes-${today}.csv\n2. Integra-Barges-${today}.csv\n3. Integra-HPMHistory-${today}.csv\n` +
      `4. Integra-ExchangeRates-${today}.csv\n5. Integra-LoginHistory-${today}.csv\n\n` +
      "Next: import each file into the matching tab of your Google Sheet."
    );
  };

  // ---- Phase 2B: live Google Sheets sync (admin only) ----
  // Fixes applied vs. the original build prompt:
  //  - fetchBargesFromSheets takes the freshly-fetched domes as a PARAMETER instead of
  //    reading the `domes` state variable directly — React state updates aren't
  //    synchronously visible in the same function scope, so the original approach would
  //    have computed barge grades against stale (pre-sync) dome data.
  //  - Location and Al2O3 % are now read back — the original silently dropped both,
  //    which would have wiped dome location data (used on the Timeline map and Stock
  //    table) on every sync.
  //  - Initial Stock and Stock Adjustments now round-trip too, so the "Stock Out"
  //    tracking and the finalize/reopen deficit-reversal feature both survive a sync.
  //
  // cleanNum strips thousands-separator commas before parsing (e.g. "10,489" -> 10489).
  // This is a defensive second layer on top of the UNFORMATTED_VALUE fix in
  // api/sheets-read.js — that fix handles number-*formatted* cells, but if a cell was
  // ever typed as literal text containing a comma, this catches that case too. Without
  // either fix, parseFloat("10,489") silently returns just 10 — it stops at the first
  // non-numeric character — which is exactly why every barge was showing "10 WMT".
  const cleanNum = (v) => parseFloat(String(v ?? "").replace(/,/g, "")) || 0;

  const mapDomeFromSheetRow = (row) => ({
    id: row["Dome ID"], contractor: row["Contractor"],
    stock: cleanNum(row["Stock (WMT)"]),
    initialStock: row["Initial Stock (WMT)"] !== "" && row["Initial Stock (WMT)"] !== undefined
      ? cleanNum(row["Initial Stock (WMT)"]) : cleanNum(row["Stock (WMT)"]),
    ni: cleanNum(row["Ni %"]), fe: cleanNum(row["Fe %"]), co: cleanNum(row["Co %"]),
    sio2: cleanNum(row["SiO2 %"]), mgo: cleanNum(row["MgO %"]),
    al2o3: cleanNum(row["Al2O3 %"]), simg: cleanNum(row["Si:Mg"]),
    location: row["Location"] || "", source: row["Source"] || "inventory",
  });

  const mapBargeFromSheetRow = (row, domesForGrade) => {
    const sourcesStr = row["Sources"] || "";
    const sources = sourcesStr.split(";").map((s) => s.trim()).filter(Boolean).map((s) => {
      const [id, amtStr] = s.split(":");
      const domeId = (id || "").trim();
      const dome = domesForGrade.find((d) => d.id === domeId);
      return { id: domeId, amt: cleanNum(amtStr), grade: dome ? dome.ni : 0 };
    });
    let stockAdjustments;
    try { stockAdjustments = row["Stock Adjustments"] ? JSON.parse(row["Stock Adjustments"]) : undefined; }
    catch (e) { stockAdjustments = undefined; }
    return {
      no: parseInt(String(row["Barge No"] ?? "").replace(/,/g, "")) || 0,
      shipDate: row["Ship Date"] || new Date().toISOString().split("T")[0],
      bargeName: row["Barge Name"] || "", tugboatName: row["Tugboat Name"] || "",
      totalWMT: cleanNum(row["Total WMT"]), grade: cleanNum(row["Grade (Ni %)"]),
      status: row["Status"] || "draft", finalized: row["Finalized"] === "Yes",
      sources, stockAdjustments,
      qtyOnBoard: row["Qty On Board"] !== "" && row["Qty On Board"] !== undefined ? cleanNum(row["Qty On Board"]) : undefined,
      progressPercent: row["Progress %"] !== "" && row["Progress %"] !== undefined ? Math.round(cleanNum(row["Progress %"])) : undefined,
      balanceDue: row["Balance Due"] !== "" && row["Balance Due"] !== undefined ? cleanNum(row["Balance Due"]) : undefined,
      lastUpdated: row["Last Updated"] || undefined,
    };
  };

  // Vercel's own crash page (function throws before reaching our try/catch — e.g. a
  // missing npm package) returns HTML/plain-text, not JSON, so response.json() alone
  // would just silently fail and hide the real cause behind a bare "HTTP 500". Read the
  // raw text first and fall back to a trimmed snippet of it, which usually contains the
  // actual crash reason (e.g. "Cannot find module 'googleapis'").
  const extractErrorDetail = async (response) => {
    if (response.status === 404) return "HTTP 404 — endpoint not found. Check that this API file was actually committed and deployed.";
    const raw = await response.text();
    try {
      const body = JSON.parse(raw);
      return body.details || body.error || `HTTP ${response.status}`;
    } catch (e) {
      // Vercel's crash/error pages are HTML and often include a <style> block — plain
      // tag-stripping alone leaves that block's raw CSS text behind as unreadable
      // noise (exactly what showed up in the UI before this fix). Strip style/script
      // blocks by their full content, not just their tags, before the generic strip.
      const cleaned = raw.replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<script[\s\S]*?<\/script>/gi, " ");
      const snippet = cleaned.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 200);
      return snippet ? `HTTP ${response.status} — ${snippet}` : `HTTP ${response.status} (no response body)`;
    }
  };

  const fetchDomesFromSheets = async () => {
    try {
      const response = await fetch("/api/sheets-read?sheetName=Domes");
      if (!response.ok) throw new Error(await extractErrorDetail(response));
      const { data } = await response.json();
      if (!data.length) { setLastSyncError("Domes tab returned 0 rows — check the tab name and that it has data below the header row."); return null; }
      const transformed = data.map(mapDomeFromSheetRow);
      setDomes(transformed);
      return transformed;
    } catch (error) {
      console.error("Error fetching domes from Sheets:", error);
      setLastSyncError(`Domes: ${error.message}`);
      return null;
    }
  };

  const fetchBargesFromSheets = async (freshDomes) => {
    try {
      const response = await fetch("/api/sheets-read?sheetName=Barges");
      if (!response.ok) throw new Error(await extractErrorDetail(response));
      const { data } = await response.json();
      if (!data.length) { setLastSyncError("Barges tab returned 0 rows — check the tab name and that it has data below the header row."); return null; }
      const transformed = data.map((row) => mapBargeFromSheetRow(row, freshDomes || domes));
      setBarges(transformed);
      return transformed;
    } catch (error) {
      console.error("Error fetching barges from Sheets:", error);
      setLastSyncError(`Barges: ${error.message}`);
      return null;
    }
  };

  const fetchHpmFromSheets = async () => {
    try {
      // The Sheet tab is still named "HMAHistory" from before the dashboard renamed
      // HMA to HPM internally — targeting that exact tab name rather than asking for
      // the Sheet to be renamed.
      const response = await fetch("/api/sheets-read?sheetName=HMAHistory");
      if (!response.ok) throw new Error(await extractErrorDetail(response));
      const { data } = await response.json();
      if (!data.length) { setLastSyncError("HMAHistory tab returned 0 rows — check the tab has data below the header row."); return null; }
      const transformed = data
        .map((row) => ({ date: row["Date"], price: cleanNum(row["Price (USD/WMT)"]), unit: "USD/WMT" }))
        .filter((h) => h.date)
        .sort((a, b) => new Date(b.date) - new Date(a.date));
      // data.length > 0 but transformed is empty means every row failed the .date check —
      // almost always a header-name mismatch (expects exactly "Date" in the header row),
      // not a real "no data" case. Treating this as silent success was the actual bug:
      // an empty array is still truthy, so the old code reported "✅ Synced" while
      // quietly updating nothing.
      if (!transformed.length) { setLastSyncError('HMAHistory: 0 usable rows — check row 1 has a column header exactly "Date" (case-sensitive).'); return null; }
      setHpmHistory(transformed);
      return transformed;
    } catch (error) {
      console.error("Error fetching HPM history from Sheets:", error);
      setLastSyncError(`HPM: ${error.message}`);
      return null;
    }
  };

  const fetchExchangeRateFromSheets = async () => {
    try {
      const response = await fetch("/api/sheets-read?sheetName=ExchangeRates");
      if (!response.ok) throw new Error(await extractErrorDetail(response));
      const { data } = await response.json();
      if (!data.length) { setLastSyncError("ExchangeRates tab returned 0 rows — check the tab has data below the header row."); return null; }
      const transformed = data
        .map((row) => ({ date: row["Date"], rate: cleanNum(row["Rate (IDR/USD)"]), source: row["Source"] || "manual" }))
        .filter((e) => e.date)
        .sort((a, b) => new Date(b.date) - new Date(a.date));
      if (!transformed.length) { setLastSyncError('ExchangeRates: 0 usable rows — check row 1 has a column header exactly "Date" (case-sensitive).'); return null; }
      setExchangeRateHistory(transformed);
      return transformed;
    } catch (error) {
      console.error("Error fetching exchange rate history from Sheets:", error);
      setLastSyncError(`Exchange Rate: ${error.message}`);
      return null;
    }
  };


  const writeToSheets = async (sheetName, rows) => {
    try {
      const response = await fetch("/api/sheets-write", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sheetName, data: rows }),
      });
      if (!response.ok) throw new Error(`Failed to write ${sheetName} to Sheets`);
      return true;
    } catch (error) {
      console.error(`Error writing ${sheetName} to Sheets:`, error);
      return false;
    }
  };
  // Accept an explicit domes/barges array (used right after a state update, where the
  // component's own `domes`/`barges` closure would still be the stale pre-update
  // value) — falls back to current state for the periodic auto-sync case.
  const writeDomesToSheets = (domesOverride) => {
    const list = domesOverride || domes;
    const rows = toRows(
      ["Dome ID", "Contractor", "Stock (WMT)", "Initial Stock (WMT)", "Ni %", "Fe %", "Co %", "SiO2 %", "MgO %", "Al2O3 %", "Si:Mg", "Location", "Source"],
      list.map((d) => [d.id, d.contractor, d.stock, d.initialStock !== undefined ? d.initialStock : d.stock, fmt(d.ni, 2), fmt(d.fe, 2), fmt(d.co, 2), fmt(d.sio2, 2), fmt(d.mgo, 2), fmt(d.al2o3, 2), fmt(d.simg, 2), d.location || "", d.source])
    );
    return writeToSheets("Domes", rows);
  };
  const writeBargesToSheets = (bargesOverride) => {
    const list = bargesOverride || barges;
    const rows = toRows(
      ["Barge No", "Ship Date", "Barge Name", "Tugboat Name", "Total WMT", "Grade (Ni %)", "Status", "Finalized", "Sources", "Stock Adjustments", "Qty On Board", "Progress %", "Balance Due", "Last Updated"],
      list.map((b) => [
        String(b.no).padStart(2, "0"), b.shipDate, b.bargeName || "", b.tugboatName || "",
        fmt(b.totalWMT), fmt(b.grade, 2), b.status, b.finalized ? "Yes" : "No",
        (b.sources || []).map((s) => `${s.id}:${s.amt}WMT`).join("; "),
        b.stockAdjustments && Object.keys(b.stockAdjustments).length ? JSON.stringify(b.stockAdjustments) : "",
        b.qtyOnBoard || "", b.progressPercent || "", b.balanceDue || "", b.lastUpdated || "",
      ])
    );
    return writeToSheets("Barges", rows);
  };
  const writeHpmToSheets = (hpmOverride) => {
    const list = hpmOverride || hpmHistory;
    const rows = toRows(["Date", "Price (USD/WMT)"], [...list].sort((a, b) => new Date(b.date) - new Date(a.date)).map((h) => [h.date, h.price.toFixed(2)]));
    return writeToSheets("HMAHistory", rows);
  };
  const writeExchangeRateToSheets = (rateOverride) => {
    const list = rateOverride || exchangeRateHistory;
    const rows = toRows(["Date", "Rate (IDR/USD)", "Source"], [...list].sort((a, b) => new Date(b.date) - new Date(a.date)).map((e) => [e.date, e.rate, e.source || "manual"]));
    return writeToSheets("ExchangeRates", rows);
  };

  // ---- Feature flags & account management (dev-only configuration, applies to
  // whoever is actually logged in — see the state comment above for why the default
  // is fail-open). ----
  const loadUserFeaturesFromSheets = async (username) => {
    try {
      const response = await fetch("/api/sheets-read?sheetName=FeatureFlags");
      if (!response.ok) return; // keep the fail-open defaults
      const { data } = await response.json();
      const userConfig = data.find((row) => row["Username"] === username);
      if (!userConfig) return; // no row for this user yet — keep fail-open defaults
      setUserFeatures({
        stock: userConfig["Stock"] === "true" || userConfig["Stock"] === true,
        barging: userConfig["Barging"] === "true" || userConfig["Barging"] === true,
        timeline: userConfig["Timeline"] === "true" || userConfig["Timeline"] === true,
        financials: userConfig["Financials"] === "true" || userConfig["Financials"] === true,
        loginLog: userConfig["LoginLog"] === "true" || userConfig["LoginLog"] === true,
        settings: userConfig["Settings"] === "true" || userConfig["Settings"] === true,
        chatAssistant: userConfig["ChatAssistant"] === "true" || userConfig["ChatAssistant"] === true,
      });
    } catch (error) {
      console.error("Error loading user features:", error);
    }
  };

  const loadAllUsersAndFlags = async () => {
    try {
      const usersRes = await fetch("/api/sheets-read?sheetName=Users");
      if (usersRes.ok) { const { data } = await usersRes.json(); setAllUsers(data); }
      const flagsRes = await fetch("/api/sheets-read?sheetName=FeatureFlags");
      if (flagsRes.ok) {
        const { data } = await flagsRes.json();
        const flagsObj = {};
        data.forEach((row) => { flagsObj[row["Username"]] = row; });
        setAllFeatureFlags(flagsObj);
      }
    } catch (error) {
      console.error("Error loading users/flags:", error);
    }
  };

  const writeFeatureFlagsToSheets = async (flagsData) => {
    const boolStr = (v) => (v === true || v === "true" ? "true" : "false");
    const rows = toRows(
      ["Username", "Stock", "Barging", "Timeline", "Financials", "LoginLog", "Settings", "ChatAssistant"],
      Object.entries(flagsData).map(([username, flags]) => [
        username, boolStr(flags["Stock"]), boolStr(flags["Barging"]), boolStr(flags["Timeline"]),
        boolStr(flags["Financials"]), boolStr(flags["LoginLog"]), boolStr(flags["Settings"]), boolStr(flags["ChatAssistant"]),
      ])
    );
    return writeToSheets("FeatureFlags", rows);
  };

  const writeUsersToSheets = async (usersData) => {
    const rows = toRows(
      ["Username", "Password", "Role", "Status", "Created_Date", "Last_Modified"],
      usersData.map((u) => [u["Username"], u["Password"], u["Role"], u["Status"], u["Created_Date"], u["Last_Modified"]])
    );
    return writeToSheets("Users", rows);
  };

  // Login history — previously only ever written to localStorage (per-browser, never
  // shared) and exported as a one-off manual CSV. Nothing ever pushed it live to the
  // LoginHistory Sheet tab, which is exactly why that tab stayed empty despite the
  // dashboard itself showing plenty of entries. Every login attempt now writes here too.
  const writeLoginHistoryToSheets = async (historyOverride) => {
    const list = historyOverride || loginHistory;
    const rows = toRows(["Timestamp", "Username", "Status"], list.map((l) => [l.timestamp, l.username, l.status]));
    return writeToSheets("LoginHistory", rows);
  };
  const fetchLoginHistoryFromSheets = async () => {
    try {
      const response = await fetch("/api/sheets-read?sheetName=LoginHistory");
      if (!response.ok) throw new Error(await extractErrorDetail(response));
      const { data } = await response.json();
      const transformed = data
        .map((row) => ({ timestamp: row["Timestamp"], username: row["Username"], status: row["Status"] }))
        .filter((l) => l.timestamp)
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
        .slice(0, 200);
      setLoginHistory(transformed);
      return transformed;
    } catch (error) {
      console.error("Error fetching login history from Sheets:", error);
      return null; // don't block the rest of sync on this — login log is a nice-to-have, not core data
    }
  };

  // ---- Meta — small key/value sheet for cross-team single values. Currently just
  // DataLastUpdated (the "Data updated" date badge), so an Excel import from any admin's
  // browser updates that date for everyone, not just locally for whoever did the import.
  const writeMetaToSheets = async (updates) => {
    const rows = toRows(["Key", "Value"], Object.entries(updates).map(([k, v]) => [k, v]));
    return writeToSheets("Meta", rows);
  };
  const fetchMetaFromSheets = async () => {
    try {
      const response = await fetch("/api/sheets-read?sheetName=Meta");
      if (!response.ok) throw new Error(await extractErrorDetail(response));
      const { data } = await response.json();
      const row = data.find((r) => r["Key"] === "DataLastUpdated");
      if (row?.Value) setDataLastUpdated(row.Value);
      return true;
    } catch (error) {
      console.error("Error fetching Meta from Sheets:", error);
      return null; // best-effort — falls back to whatever's already in state/localStorage
    }
  };

  // ---- Sub-feature (granular) flags — a finer layer on top of the tab-level flags
  // above. If a user has no row in FeatureFlagsDetailed yet, their tab-level flags are
  // expanded into detailed flags (whole tab on -> every sub-feature on, and vice versa)
  // rather than defaulting to something disconnected from what they already have.
  const convertTabFlagsToDetailed = (tabFlags) => {
    const detailed = {};
    Object.entries(FEATURE_STRUCTURE).forEach(([tabKey, tabConfig]) => {
      const tabEnabled = tabFlags[tabKey] === "true" || tabFlags[tabKey] === true;
      tabConfig.features.forEach((feature) => { detailed[feature.key] = tabEnabled ? "true" : "false"; });
    });
    return detailed;
  };
  const isTabEnabled = (detailedFlags, tabKey) => {
    if (!FEATURE_STRUCTURE[tabKey]) return false;
    return FEATURE_STRUCTURE[tabKey].features.some((f) => detailedFlags[f.key] === "true" || detailedFlags[f.key] === true);
  };

  const loadUserDetailedFeatures = async (username) => {
    try {
      const detailedRes = await fetch("/api/sheets-read?sheetName=FeatureFlagsDetailed");
      if (detailedRes.ok) {
        const { data } = await detailedRes.json();
        const userConfig = data.find((row) => row["Username"] === username);
        if (userConfig) { setUserDetailedFeatures(userConfig); return; }
      }
      // No detailed row for this user — derive detailed flags from their tab-level
      // flags instead of silently defaulting to something disconnected from what
      // they already have.
      const tabRes = await fetch("/api/sheets-read?sheetName=FeatureFlags");
      if (!tabRes.ok) return;
      const { data } = await tabRes.json();
      const userConfig = data.find((row) => row["Username"] === username);
      if (userConfig) setUserDetailedFeatures(convertTabFlagsToDetailed(userConfig));
    } catch (error) {
      console.error("Error loading detailed features:", error);
    }
  };

  const loadAllDetailedFlags = async () => {
    try {
      const detailedRes = await fetch("/api/sheets-read?sheetName=FeatureFlagsDetailed");
      if (detailedRes.ok) {
        const { data } = await detailedRes.json();
        const flagsObj = {};
        data.forEach((row) => { flagsObj[row["Username"]] = row; });
        setAllDetailedFlags(flagsObj);
        return;
      }
      const tabRes = await fetch("/api/sheets-read?sheetName=FeatureFlags");
      if (!tabRes.ok) return;
      const { data } = await tabRes.json();
      const flagsObj = {};
      data.forEach((row) => { flagsObj[row["Username"]] = convertTabFlagsToDetailed(row); });
      setAllDetailedFlags(flagsObj);
    } catch (error) {
      console.error("Error loading all detailed flags:", error);
    }
  };

  const DETAILED_HEADERS = ["Username", ...Object.values(FEATURE_STRUCTURE).flatMap((t) => t.features.map((f) => f.key))];

  const writeDetailedFeatureFlagsToSheets = async (detailedFlagsData) => {
    // Keep the simple tab-level sheet in sync too — derived as "on" if ANY sub-feature
    // in that tab is on, so the fast/simple tab-level toggle stays meaningful even
    // after fine-tuning individual sub-features.
    const tabRows = toRows(
      ["Username", "Stock", "Barging", "Timeline", "Financials", "LoginLog", "Settings", "ChatAssistant"],
      Object.entries(detailedFlagsData).map(([username, flags]) => [
        username,
        isTabEnabled(flags, "Stock") ? "true" : "false", isTabEnabled(flags, "Barging") ? "true" : "false",
        isTabEnabled(flags, "Timeline") ? "true" : "false", isTabEnabled(flags, "Financials") ? "true" : "false",
        isTabEnabled(flags, "LoginLog") ? "true" : "false", isTabEnabled(flags, "Settings") ? "true" : "false",
        flags["ChatAssistant_Widget"] === "true" || flags["ChatAssistant_Widget"] === true ? "true" : "false",
      ])
    );
    const tabOk = await writeToSheets("FeatureFlags", tabRows);

    const detailedRows = toRows(
      DETAILED_HEADERS,
      Object.entries(detailedFlagsData).map(([username, flags]) => [
        username, ...DETAILED_HEADERS.slice(1).map((key) => (flags[key] === true || flags[key] === "true" ? "true" : "false")),
      ])
    );
    const detailedOk = await writeToSheets("FeatureFlagsDetailed", detailedRows);
    return tabOk && detailedOk;
  };

  // What components actually call to decide whether to render something. Checks the
  // detailed flag first; if that specific key was never loaded (e.g. Sheets unreachable,
  // or this key isn't in WIRED_SUBFEATURES so nobody's bothered setting it), falls back
  // to the tab-level flag so nothing breaks.
  const isFeatureEnabled = (featureKey, tabLevelFallback) => {
    if (userDetailedFeatures[featureKey] !== undefined) {
      return userDetailedFeatures[featureKey] === "true" || userDetailedFeatures[featureKey] === true;
    }
    return tabLevelFallback !== false;
  };

  const syncWithSheets = async (force = false) => {
    const lastSyncStr = localStorage.getItem("lastSheetsSyncTime");
    const lastSync = lastSyncStr ? new Date(lastSyncStr) : null;
    const minutesSinceSync = lastSync ? (Date.now() - lastSync.getTime()) / 1000 / 60 : 999;
    if (!force && minutesSinceSync < 15) return;

    setSheetsSyncStatus("Syncing…");
    setLastSyncError("");
    const freshDomes = await fetchDomesFromSheets();
    const freshBarges = freshDomes ? await fetchBargesFromSheets(freshDomes) : null;
    const freshHpm = await fetchHpmFromSheets();
    const freshExRate = await fetchExchangeRateFromSheets();
    fetchMetaFromSheets(); // best-effort, for everyone — the "Data updated" badge is visible to all roles

    // Reconcile stock against actual finalized-barge consumption rather than trusting
    // whatever the Sheet's "Stock (WMT)" column says — see reconcileStock's comment for
    // why this matters. Only "Initial Stock (WMT)" is treated as ground truth here.
    if (freshDomes && freshBarges) {
      const reconciled = freshDomes.map((d) => ({ ...d, ...reconcileStock(d.id, d.initialStock, freshBarges) }));
      const changed = reconciled.some((d, i) => d.stock !== freshDomes[i].stock || d.initialStock !== freshDomes[i].initialStock);
      if (changed) {
        setDomes(reconciled);
        writeDomesToSheets(reconciled); // self-heal the Sheet so next sync doesn't need to re-derive this
      }
    }

    // Feature flags aren't just loaded at login anymore — if dev changes what someone
    // can see mid-session, this is what actually gets it to them without forcing a
    // logout/login. Best-effort: doesn't affect the overall sync status below, since a
    // feature-flag hiccup shouldn't be reported as "sync failed" for stock/barge data.
    if (currentUser?.username) {
      loadUserFeaturesFromSheets(currentUser.username);
      loadUserDetailedFeatures(currentUser.username);
    }
    if (isAdmin) {
      fetchLoginHistoryFromSheets();
      loadAllUsersAndFlags();
      loadAllDetailedFlags();
    }

    if (freshDomes && freshBarges && freshHpm && freshExRate) {
      setLastSyncTime(new Date());
      localStorage.setItem("lastSheetsSyncTime", new Date().toISOString());
      setSheetsSyncStatus("✅ Synced");
    } else {
      setSheetsSyncStatus("❌ Sync failed — using local data");
    }
  };

  // Fires after a discrete, deliberate data change (Excel/CSV import onto a barge) —
  // deliberately NOT wired into every inline edit (typing a quantity, renaming a barge),
  // since those fire on every keystroke and would hammer the Sheets API.
  //
  // Accepts an optional freshBarges array from the caller. This matters: this function's
  // own `barges` closure is fixed at the moment it's called, and setBarges (called just
  // before this, synchronously) doesn't apply until the next render — so without an
  // explicit override, writeBargesToSheets() would silently push the PRE-import data to
  // Sheets. That's exactly what caused barge Excel imports to appear to work locally but
  // never actually reach Sheets, then get overwritten by stale data on the next sync.
  const onDataCommitted = (freshBarges) => {
    setTimeout(() => { writeDomesToSheets(); writeBargesToSheets(freshBarges); }, 500);
  };

  const handleLoadingReportParsed = (parsedData) => {
    if (parsedData.missingFields?.length) console.warn("⚠️ Loading report — missing fields:", parsedData.missingFields);
    setReviewData(parsedData);
    setShowLoadingReportModal(false);
    setShowReviewModal(true);
  };

  const handleLoadingProgressConfirm = async (confirmedData) => {
    setIsSubmittingLoadingReport(true);
    try {
      const bargeExists = barges.some((b) => b.no === confirmedData.shipmentNumber);
      if (!bargeExists) { alert(`❌ Barge #${confirmedData.shipmentNumber} not found`); return; }

      const updatedBarges = barges.map((b) => b.no !== confirmedData.shipmentNumber ? b : {
        ...b,
        qtyOnBoard: confirmedData.qtyOnBoard || 0,
        progressPercent: confirmedData.progressPercent || 0,
        balanceDue: confirmedData.balanceDue || 0,
        lastUpdated: confirmedData.reportDate || new Date().toISOString().split("T")[0],
        // Separate field from the existing Ni-blend-quality `status` (exact/excess/
        // deficit/etc, used by StatusBadge and the Financials royalty table) — reusing
        // that field for loading progress would have broken both.
        loadingStatus: confirmedData.progressPercent === 100 ? "loaded" : "loading",
      });
      setBarges(updatedBarges);
      await writeBargesToSheets(updatedBarges);
      alert(`✅ Loading progress saved! Barge #${confirmedData.shipmentNumber} updated.`);
      setShowReviewModal(false);
      setReviewData(null);
    } catch (error) {
      console.error("Error saving loading progress:", error);
      alert("❌ Error: " + error.message);
    } finally {
      setIsSubmittingLoadingReport(false);
    }
  };

  // finalize / reopen actually mutates the master stock here, since it needs setDomes
  // Does the actual state mutation for finalize/reopen — unconditional, no checks.
  // Called either directly (no deficit) or after explicit confirmation (deficit present).
  const applyFinalizeOrReopen = (no) => {
    const barge = barges.find((b) => b.no === no);
    if (!barge) return;
    const sign = barge.finalized ? 1 : -1; // reopening adds back; finalizing subtracts

    let updatedDomes, updatedBarges;
    if (sign === -1) {
      // Finalizing: subtract normally where stock covers it. Where it doesn't (only
      // reachable after explicit confirmation via the alert modal), clamp that dome to
      // zero rather than negative, and bump its initialStock by the shortfall so the
      // Stock page's ledger stays consistent. The exact shortfall per dome is recorded
      // on the barge itself so reopening can reverse precisely this adjustment.
      const adjustments = {};
      updatedDomes = domes.map((d) => {
        const used = barge.sources.filter((s) => s.id === d.id).reduce((s, x) => s + x.amt, 0);
        if (used <= 0) return d;
        if (d.stock < used) {
          const deficit = used - d.stock;
          adjustments[d.id] = deficit;
          const priorInitial = d.initialStock !== undefined ? d.initialStock : d.stock;
          return { ...d, stock: 0, initialStock: priorInitial + deficit };
        }
        return { ...d, stock: d.stock - used };
      });
      updatedBarges = barges.map((b) => b.no === no ? { ...b, finalized: true, stockAdjustments: adjustments } : b);
    } else {
      // Reopening: reverse exactly what finalize did, including any deficit adjustment
      // that was applied — not just adding the barged amount back blindly.
      const adjustments = barge.stockAdjustments || {};
      updatedDomes = domes.map((d) => {
        const used = barge.sources.filter((s) => s.id === d.id).reduce((s, x) => s + x.amt, 0);
        if (used <= 0) return d;
        const deficit = adjustments[d.id] || 0;
        const priorInitial = d.initialStock !== undefined ? d.initialStock : d.stock;
        return { ...d, stock: Math.max(0, d.stock + used - deficit), initialStock: priorInitial - deficit };
      });
      updatedBarges = barges.map((b) => b.no === no ? { ...b, finalized: false, stockAdjustments: undefined } : b);
    }
    setDomes(updatedDomes);
    setBarges(updatedBarges);
    writeDomesToSheets(updatedDomes); writeBargesToSheets(updatedBarges);
    setPendingFinalize(null);
  };

  // Entry point from the Finalize button: checks for a stock deficit first. If found,
  // stop and show a confirmation modal rather than finalizing (or silently adjusting)
  // right away — the person needs to see exactly which domes are short before deciding.
  const toggleFinalize = (no) => {
    const barge = barges.find((b) => b.no === no);
    if (!barge) return;
    const sign = barge.finalized ? 1 : -1;

    if (sign === -1) {
      const violations = [];
      barge.sources.forEach((s) => {
        const dome = domes.find((d) => d.id === s.id);
        if (dome && s.amt > dome.stock) {
          violations.push({ domeId: s.id, current: dome.stock, requested: s.amt });
        }
      });

      // Also flag any source dome still missing lab data (0% on any assay field).
      // This never blocks finalizing — same as the stock deficit case, it's a
      // confirm-and-proceed alert, since the physical barge may already be loaded
      // and shipped before the lab result comes back. The alert exists so the person
      // is reminded to go update the stock database once the real assay is in.
      const labIssues = [];
      barge.sources.forEach((s) => {
        const dome = domes.find((d) => d.id === s.id);
        if (!dome) return;
        const zeroFields = LAB_FIELDS.filter(([key]) => (dome[key] || 0) === 0).map(([, label]) => label);
        if (zeroFields.length) labIssues.push({ domeId: s.id, fields: zeroFields });
      });

      if (violations.length > 0 || labIssues.length > 0) {
        setPendingFinalize({ bargeNo: no, violations, labIssues });
        return; // wait for explicit confirm/cancel from the alert modal
      }
    }
    applyFinalizeOrReopen(no);
  };

  const handleImport = (e) => {
    const files = e.target.files;
    if (!files || !files.length) return;
    let merged = [...domes];
    let processed = 0;
    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const name = file.name.toLowerCase();
          let source = "inventory", forced = null;
          if (name.includes("imn-1")) { source = "production"; forced = "IMN-1"; }
          else if (name.includes("imn-2")) { source = "production"; forced = "IMN-2"; }
          else if (name.includes("imn-3")) { source = "production"; forced = "IMN-3"; }
          else if (name.includes("imn-4")) { source = "production"; forced = "IMN-4"; }
          else if (name.includes("production")) { source = "production"; }
          const rows = parseDomeCSV(evt.target.result, source, forced);
          // The imported "stock" figure is treated as INITIAL/gross stock, not current
          // remaining — same rule as manual chat-based stock updates. Net out whatever
          // has already been consumed by FINALIZED barges, clamping to 0 (and bumping
          // initialStock by the deficit) rather than silently resetting stock to the raw
          // imported figure, which would have erased that consumption bookkeeping.
          const adjustedRows = rows.map((r) => ({ ...r, ...reconcileStock(r.id, r.stock, barges) }));
          const ids = new Set(adjustedRows.map((r) => r.id));
          merged = merged.filter((d) => !ids.has(d.id)).concat(adjustedRows);
        } catch (err) { console.error(err); }
        processed++;
        if (processed === files.length) {
          setDomes(merged);
          // Standing rule: barge source grades must be re-synced whenever dome stock/lab
          // data is updated from Excel — a barge's stored source grade is a SNAPSHOT
          // taken when it was built, not a live link, so it goes stale the moment the
          // underlying dome's Ni% changes. This was previously only ever done as a
          // one-off manual reconciliation; doing it automatically here is what makes the
          // in-app Import button actually safe to use directly instead of going through
          // chat for every stock update.
          const mergedById = {}; merged.forEach((d) => { mergedById[d.id] = d; });
          const updatedBarges = barges.map((b) => {
            let changed = false;
            const newSources = b.sources.map((s) => {
              const d = mergedById[s.id];
              if (d && Math.abs((d.ni || 0) - s.grade) > 1e-9) { changed = true; return { ...s, grade: d.ni }; }
              return s;
            });
            if (!changed) return b;
            const totalWMT = newSources.reduce((sum, s) => sum + s.amt, 0);
            const niSum = newSources.reduce((sum, s) => sum + s.amt * (s.grade / 100), 0);
            const grade = totalWMT > 0 ? (niSum / totalWMT) * 100 : 0;
            return { ...b, sources: newSources, totalWMT, grade, status: statusFor({ totalWMT, grade }, settings.bargeSize, settings.targetGrade, settings.tolerance) };
          });
          setBarges(updatedBarges);
          setDataLastUpdated(importDate);
          // Push straight to Sheets, same as the loading-report chat assistant does —
          // previously this only updated local state, so the import looked like it
          // worked but silently never reached Sheets. The next sync (or the next
          // person's login) would then pull the OLD sheet data back over it, making the
          // import appear to have been lost.
          if (isAdmin) {
            writeDomesToSheets(merged);
            writeBargesToSheets(updatedBarges);
            writeMetaToSheets({ DataLastUpdated: importDate });
          }
          setImportStatus(`✓ Imported ${merged.length} domes`);
          setTimeout(() => { setShowImport(false); setImportStatus(""); }, 1500);
        }
      };
      reader.readAsText(file);
    });
  };

  // Auth gates — placed after all hooks (React's Rules of Hooks require every hook to
  // run on every render, so these early returns must come after every useState/useEffect
  // above, not before). Each includes its own <style>{CSS}</style> because these bypass
  // the main .app wrapper below, which is otherwise the only place the stylesheet is
  // injected — without this, the login/welcome screens render with zero CSS applied.
  if (!isLoggedIn) {
    return (
      <>
        <style>{CSS}</style>
        <LoginScreen onLogin={handleLogin} error={loginError} />
      </>
    );
  }
  if (showWelcome) {
    return (
      <>
        <style>{CSS}</style>
        <WelcomeScreen username={currentUser?.username} role={currentUser?.role} />
      </>
    );
  }
  if (domes.length === 0 && !lastSyncTime) {
    const failed = sheetsSyncStatus.startsWith("❌");
    return (
      <>
        <style>{CSS}</style>
        <SyncingScreen syncFailed={failed} syncError={failed ? lastSyncError : null} />
      </>
    );
  }

  return (
    <div className="app">
      <style>{CSS}</style>
      <div className="bg-glow bg-glow-a" />
      <div className="bg-glow bg-glow-b" />

      {currentUser?.username === "dev" && (
        <div className="test-badge">🧪 TEST ACCOUNT</div>
      )}

      {isDevAccount && isFeatureEnabled("ChatAssistant_Widget", userFeatures.chatAssistant) && (
        <button className="loading-assistant-btn" onClick={() => setShowLoadingReportModal(true)} title="Parse shipment loading report">
          <MessageSquare size={20} />
          <span>Chat</span>
        </button>
      )}
      {isDevAccount && showLoadingReportModal && (
        <LoadingReportModal onClose={() => setShowLoadingReportModal(false)} onSubmit={handleLoadingReportParsed} />
      )}
      {isDevAccount && showReviewModal && reviewData && (
        <LoadingReportReviewModal data={reviewData} onClose={() => { setShowReviewModal(false); setReviewData(null); }}
          onConfirm={handleLoadingProgressConfirm} isSubmitting={isSubmittingLoadingReport} />
      )}

      {showImport && (
        <div className="import-modal">
          <div className="import-panel glass">
            <div className="import-head">
              <Upload size={20} style={{ color: "#E35F0C" }} />
              <span>Import from Google Sheets</span>
              <button className="import-close" onClick={() => setShowImport(false)}><X size={18} /></button>
            </div>
            <div className="import-body">
              <p>Upload CSVs exported from each sheet (Dome Inventory Existing, Dome Production IMN-1..4).</p>
              <p style={{ fontSize: "10.5px", color: "#667080" }}>Tip: name the file with "imn-1" / "imn-2" etc. so it's tagged to the right contractor.</p>
              <div className="form-group" style={{ marginBottom: 14 }}>
                <label>Data date (shown as "Data updated" on the dashboard)</label>
                <input type="date" className="login-input" value={importDate} onChange={(e) => setImportDate(e.target.value)} />
              </div>
              <label className="import-input-label">
                <input type="file" accept=".csv" multiple onChange={handleImport} />
                <span>Select CSV file(s)</span>
              </label>
              {importStatus && <div className="import-status">{importStatus}</div>}
            </div>
          </div>
        </div>
      )}

      {statusBarge && (
        <BargeStatusModal barge={statusBarge} domes={domes} onClose={() => setStatusBarge(null)} />
      )}

      {pendingFinalize && (
        <div className="validation-modal">
          <div className="validation-panel glass">
            <div className="validation-head">
              <AlertTriangle size={20} style={{ color: "#F87171" }} />
              <span>Confirm Before Finalizing</span>
              <button className="validation-close" onClick={() => setPendingFinalize(null)}><X size={18} /></button>
            </div>
            <div className="validation-body">
              {pendingFinalize.violations.length > 0 && (
                <>
                  <div className="status-section-title" style={{ color: "#F87171" }}>
                    <AlertTriangle size={13} /> Stock deficit
                  </div>
                  <p>Barge #{pendingFinalize.bargeNo} would cause the following domes to drop below zero. If you finalize anyway, each of these domes will be set to <strong>0 remaining</strong> (not negative), and the ledger will be updated to reflect the shortfall. Please verify with your data provider when convenient:</p>
                  <div className="violations-list">
                    {pendingFinalize.violations.map((v) => (
                      <div key={v.domeId} className="violation-item">
                        <span className="violation-dome">{v.domeId}</span>
                        <span className="violation-detail">Current: {fmt(v.current)} WMT | Requested: {fmt(v.requested)} WMT | Deficit: {fmt(v.requested - v.current)} WMT</span>
                      </div>
                    ))}
                  </div>
                </>
              )}

              {pendingFinalize.labIssues.length > 0 && (
                <>
                  <div className="status-section-title" style={{ color: "#FBBF24" }}>
                    <AlertTriangle size={13} /> Missing lab data
                  </div>
                  <p>The following domes in Barge #{pendingFinalize.bargeNo} still show 0% on one or more assay fields. You can finalize before the lab result comes back — just make sure to update the stock database with the real values once it does:</p>
                  <div className="lab-list">
                    {pendingFinalize.labIssues.map((v) => (
                      <div key={v.domeId} className="lab-item">
                        <span className="lab-dome">{v.domeId}</span>
                        <span className="lab-detail">{v.fields.join(", ")} showing 0%</span>
                      </div>
                    ))}
                  </div>
                </>
              )}

              <div className="validation-actions">
                <button className="btn-ghost" onClick={() => setPendingFinalize(null)}>Cancel</button>
                <button className="btn-primary btn-danger" onClick={() => applyFinalizeOrReopen(pendingFinalize.bargeNo)}>
                  Confirm &amp; Finalize Anyway
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {invoiceBarge && (
        <InvoiceModal barge={invoiceBarge} onClose={() => setInvoiceBarge(null)} />
      )}

      {exportBarge && (
        <BargeExportModal barge={exportBarge} domesById={domesByIdTop} onClose={() => setExportBarge(null)} />
      )}

      <header className="topbar">
        <div className="brand">
          <svg width="40" height="46" viewBox="0 0 700 820" className="brand-logo">
            {/* orange dots (top) */}
            <rect x="62" y="57" width="185" height="62" fill="#E35F0C" />
            <rect x="62" y="167" width="185" height="63" fill="#E35F0C" />
            {/* gold zigzag */}
            <rect x="62" y="278" width="587" height="62" fill="#BFB12A" />
            <rect x="62" y="278" width="66" height="299" fill="#BFB12A" />
            <rect x="174" y="387" width="475" height="62" fill="#BFB12A" />
            <rect x="174" y="387" width="73" height="190" fill="#BFB12A" />
            {/* orange bracket (bottom) */}
            <rect x="389" y="501" width="59" height="185" fill="#E35F0C" />
            <rect x="62" y="624" width="386" height="62" fill="#E35F0C" />
            <rect x="495" y="501" width="65" height="294" fill="#E35F0C" />
            <rect x="62" y="734" width="498" height="61" fill="#E35F0C" />
          </svg>
          <div>
            <div className="brand-name">INTEGRA</div>
            <div className="brand-sub">Nickel Ore Barging Plan · 2026</div>
            <div className="data-updated-badge" title="Last time stock/barge Excel data was updated">
              <Calendar size={11} /> Data updated {fmtShortDate(dataLastUpdated)}
            </div>
          </div>
        </div>
        <div className="nav-desktop-wrapper">
          <nav className="nav-desktop">
            <NavButton icon={LayoutGrid} label="Overview" active={tab === "overview"} onClick={() => setTab("overview")} />
            {userFeatures.stock && <NavButton icon={Layers} label="Stock" active={tab === "stock"} onClick={() => setTab("stock")} />}
            {userFeatures.barging && <NavButton icon={Ship} label="Barging Plan" active={tab === "plan"} onClick={() => setTab("plan")} />}
            {userFeatures.timeline && <NavButton icon={Calendar} label="Timeline" active={tab === "timeline"} onClick={() => setTab("timeline")} />}
            {isAdmin && userFeatures.financials && <NavButton icon={DollarSign} label="Financials" active={tab === "financials"} onClick={() => setTab("financials")} />}
            {isAdmin && userFeatures.loginLog && <NavButton icon={History} label="Login Log" active={tab === "log"} onClick={() => setTab("log")} />}
            {(userFeatures.settings || currentUser?.username === "dev") && <NavButton icon={Settings} label="Settings" active={tab === "settings"} onClick={() => setTab("settings")} />}
          </nav>
          {isFeatureEnabled("Stock_ImportExcel", true) && (
            <button className="btn-import" onClick={() => setShowImport(!showImport)}><Upload size={14} /> Import</button>
          )}
          <button className="btn-logout" onClick={handleLogout} title={`Log out (${currentUser?.username})`}>
            <LogOut size={13} /> <span className="btn-logout-label">Logout</span>
          </button>
          <MobileNavMenu tab={tab} setTab={setTab} isAdmin={isAdmin} currentUser={currentUser} handleLogout={handleLogout} userFeatures={userFeatures} />
        </div>
      </header>

      <main className="content">
        {tab === "overview" && <OverviewTab domes={domes} barges={barges} settings={settings} />}
        {tab === "stock" && userFeatures.stock && <StockTab domes={domes} />}
        {tab === "plan" && userFeatures.barging && <PlanTabWired domes={domes} settings={settings} barges={barges} setBarges={setBarges} toggleFinalize={toggleFinalize} onOpenInvoice={setInvoiceBarge} onExportBarge={setExportBarge} onCheckStatus={setStatusBarge} onDataCommitted={onDataCommitted} isFeatureEnabled={isFeatureEnabled} />}
        {tab === "timeline" && userFeatures.timeline && <TimelineTab domes={domes} settings={settings} barges={barges} isDevAccount={isDevAccount} />}
        {tab === "financials" && isAdmin && userFeatures.financials && (
          <FinancialsTab
            barges={barges} hpmHistory={hpmHistory} setHpmHistory={setHpmHistory}
            exchangeRateHistory={exchangeRateHistory} setExchangeRateHistory={setExchangeRateHistory}
            getHpmOnDate={getHpmOnDate} getExchangeRateOnDate={getExchangeRateOnDate}
            calculateRoyalty={calculateRoyalty} getHpmTrendPercent={getHpmTrendPercent}
            getExRateTrendPercent={getExRateTrendPercent} exportFinancialData={exportFinancialData}
            onHpmUpdated={writeHpmToSheets} onExchangeRateUpdated={writeExchangeRateToSheets}
          />
        )}
        {tab === "log" && isAdmin && userFeatures.loginLog && <LoginLogTab loginHistory={loginHistory} />}
        {tab === "settings" && (userFeatures.settings || currentUser?.username === "dev") && (
          <SettingsTab isAdmin={isAdmin} currentUser={currentUser} handleLogout={handleLogout} exportAllForGoogleSheets={exportAllForGoogleSheets}
            syncWithSheets={syncWithSheets} sheetsSyncStatus={sheetsSyncStatus} lastSyncTime={lastSyncTime} lastSyncError={lastSyncError} exRateFetchStatus={exRateFetchStatus}
            allUsers={allUsers} allFeatureFlags={allFeatureFlags} setAllUsers={setAllUsers} setAllFeatureFlags={setAllFeatureFlags}
            writeUsersToSheets={writeUsersToSheets} writeFeatureFlagsToSheets={writeFeatureFlagsToSheets}
            allDetailedFlags={allDetailedFlags} setAllDetailedFlags={setAllDetailedFlags} writeDetailedFeatureFlagsToSheets={writeDetailedFeatureFlagsToSheets}
            isFeatureEnabled={isFeatureEnabled} />
        )}
      </main>
    </div>
  );
}

/* PlanTabWired: thin wrapper so finalize can reach the top-level setDomes */
function PlanTabWired({ domes, settings, barges, setBarges, toggleFinalize, onOpenInvoice, onExportBarge, onCheckStatus, onDataCommitted, isFeatureEnabled }) {
  const [filter, setFilter] = useState("all");
  const [genCount, setGenCount] = useState(5);
  const [genContractors, setGenContractors] = useState(null); // null = all
  const [genTargetNi, setGenTargetNi] = useState(settings.targetGrade);
  const [genQty, setGenQty] = useState(settings.bargeSize);
  const [genSources, setGenSources] = useState({ inventory: true, production: true });
  const [genSimgOp, setGenSimgOp] = useState("none"); // "none" | "lte" | "gte" — optional, skipped when "none"
  const [genSimgValue, setGenSimgValue] = useState("");

  // Contractor list shown in the checklist depends on which Source checkboxes are checked:
  // Inventory only -> inventory contractors; Production only -> IMN-1..4; both -> everyone.
  const activeSourceKeys = useMemo(
    () => Object.entries(genSources).filter(([, v]) => v).map(([k]) => k),
    [genSources.inventory, genSources.production]
  );
  const allContractors = useMemo(() => {
    const visible = domes.filter((d) => activeSourceKeys.length === 0 || activeSourceKeys.includes(d.source || "inventory"));
    return Array.from(new Set(visible.map((d) => d.contractor))).sort();
  }, [domes, activeSourceKeys]);

  // Reset contractor selection back to "all" whenever the Source filter changes,
  // so a stale selection from the other source group can't linger invisibly.
  useEffect(() => { setGenContractors(null); }, [activeSourceKeys.join(",")]);

  const domesById = useMemo(() => { const m = {}; domes.forEach((d) => (m[d.id] = d)); return m; }, [domes]);
  const pool = useMemo(() => poolFromDomes(domes, barges.filter((b) => !b.finalized)), [domes, barges]);

  const filtered = barges.filter((b) => {
    if (filter === "all") return true;
    if (filter === "final") return b.finalized;
    if (filter === "draft") return !b.finalized;
    return b.status === filter;
  });

  const finalCount = barges.filter((b) => b.finalized).length;
  const exactCount = barges.filter((b) => b.status === "exact").length;
  const problemCount = barges.filter((b) => ["excess", "deficit", "unplanned", "incomplete"].includes(b.status)).length;

  const nextNo = () => (barges.length ? Math.max(...barges.map((b) => b.no)) + 1 : 1);

  const addBlankBarge = () => {
    const no = nextNo();
    setBarges((prev) => [...prev, { no, shipDate: defaultShipDate(no, settings.planTarget), bargeName: "", tugboatName: "", sources: [], totalWMT: 0, grade: 0, status: "incomplete", finalized: false }]);
  };

  const generateBarges = () => {
    const n = Math.max(1, parseInt(genCount) || 1);
    const start = nextNo();
    const simgVal = Number(genSimgValue);
    const simgActive = genSimgOp !== "none" && genSimgValue !== "" && !isNaN(simgVal);
    const simgTarget = simgActive ? { op: genSimgOp, value: simgVal } : null;
    const genPool = poolFromDomes(domes, barges.filter((b) => !b.finalized))
      .filter((p) => genContractors === null || genContractors.includes(p.contractor))
      .filter((p) => activeSourceKeys.length === 0 || activeSourceKeys.includes(p.source))
      // Generator-only rules (Stock page is unaffected by these):
      // Rule 1 — skip domes with 250 WMT or less remaining, too small to be a practical draw.
      .filter((p) => p.remaining > 250)
      // Rule 2 — skip domes already barged past 75% of their initial stock, based on real
      // finalized consumption, not draft barges, to avoid over-relying on near-exhausted domes.
      .filter((p) => {
        const d = domesById[p.id];
        const initial = d && d.initialStock !== undefined ? d.initialStock : (d ? d.stock : null);
        if (!initial || initial <= 0) return true;
        const usedPct = (initial - d.stock) / initial;
        return usedPct <= 0.75;
      });
      // Si/Mg is NOT a pool filter — no domes are excluded because of it. It's a targeted
      // output, same as Ni: when stated, it steers which candidate domes the blender
      // reaches for first (see fillOneBarge), aiming the barge's overall Si/Mg toward the
      // target without ever removing a dome from consideration.
    const target = Number(genTargetNi) || settings.targetGrade;
    const qty = Number(genQty) > 0 ? Number(genQty) : settings.bargeSize;
    const generated = generateFromPool(genPool, n, qty, target, settings.tolerance, simgTarget);
    const withNo = generated.map((g, i) => ({ no: start + i, shipDate: defaultShipDate(start + i, settings.planTarget), bargeName: "", tugboatName: "", finalized: false, ...g }));
    setBarges((prev) => [...prev, ...withNo]);
  };

  const onUpdate = (no, sources, patch) => {
    const updated = barges.map((b) => {
      if (b.no !== no) return b;
      const next = { ...b, ...(patch || {}) };
      if (sources) {
        const totalWMT = sources.reduce((s, x) => s + x.amt, 0);
        const niSum = sources.reduce((s, x) => s + x.amt * (x.grade / 100), 0);
        const grade = totalWMT > 0 ? (niSum / totalWMT) * 100 : 0;
        next.sources = sources; next.totalWMT = totalWMT; next.grade = grade;
        next.status = statusFor({ totalWMT, grade }, settings.bargeSize, settings.targetGrade, settings.tolerance);
      }
      return next;
    });
    setBarges(updated);
    return updated; // callers (like the Excel import path) need this to push the CORRECT
    // fresh data to Sheets — reading `barges` state right after calling this would still
    // be the pre-update value, since setState doesn't apply synchronously.
  };

  const removeBarge = (no) => setBarges((prev) => prev.filter((b) => b.no !== no));

  return (
    <div className="stack">
      <section className="glass banner">
        <AlertTriangle size={18} className="banner-icon" />
        <div>
          <div className="banner-title">Barges 1–8: awaiting final submission</div>
          <div className="banner-body">
            Add or generate barges below as you need them — nothing is pre-filled. Once the confirmed plan
            for barges 01–08 is ready, add each one, import or edit its dome breakdown, and click{" "}
            <strong>Finalize</strong>. Finalizing permanently subtracts that stock from Overview and Stock.
          </div>
        </div>
      </section>

      <section className="glass summary-strip">
        <Kpi label="Finalized" value={`${finalCount}`} unit={`/ ${settings.planTarget}`} accent={finalCount > 0 ? "good" : undefined} />
        <Kpi label="Created" value={`${barges.length}`} unit={`/ ${settings.planTarget}`} />
        <Kpi label="Needs attention" value={`${problemCount}`} unit="barges" accent={problemCount > 0 ? "warn" : "good"} />
        <Kpi label="Target grade" value={fmt(settings.targetGrade, 2)} unit="% Ni" />
      </section>

      <section className="glass panel generate-panel">
        <div className="panel-head"><Ship size={16} /><span>Add to the plan</span></div>
        {isFeatureEnabled("Barging_CreateManual", true) && (
          <button className="btn-ghost add-blank-btn" onClick={addBlankBarge}><Plus size={14} /> Add blank barge</button>
        )}

        <div className="gen-form">
          <div className="gen-form-title">Suggested plan generator</div>
          <div className="gen-form-grid">
            <div className="gen-field">
              <label>Contractors</label>
              <MultiSelectDropdown
                options={allContractors}
                selected={genContractors}
                onChange={setGenContractors}
                allLabel="All contractors"
                noneLabel="No contractors"
              />
            </div>
            <div className="gen-field">
              <label>Targeted Ni content (%)</label>
              <input type="number" step="0.01" value={genTargetNi} onChange={(e) => setGenTargetNi(e.target.value)} />
            </div>
            <div className="gen-field">
              <label>Target Qty per barge (WMT)</label>
              <input type="number" min="1" step="1" value={genQty} onChange={(e) => setGenQty(e.target.value)} placeholder={String(settings.bargeSize)} />
            </div>
            <div className="gen-field">
              <label>Source</label>
              <div className="gen-source-checks">
                <label className="gen-check">
                  <input type="checkbox" checked={genSources.inventory}
                    onChange={(e) => setGenSources((s) => ({ ...s, inventory: e.target.checked }))} />
                  Existing inventory
                </label>
                <label className="gen-check">
                  <input type="checkbox" checked={genSources.production}
                    onChange={(e) => setGenSources((s) => ({ ...s, production: e.target.checked }))} />
                  Production
                </label>
              </div>
            </div>
            <div className="gen-field">
              <label># of barges</label>
              <input type="number" min="1" max="20" value={genCount} onChange={(e) => setGenCount(e.target.value)} />
            </div>
            <div className="gen-field">
              <label>Targeted Si/Mg (optional)</label>
              <div className="gen-simg-row">
                <select value={genSimgOp} onChange={(e) => setGenSimgOp(e.target.value)}>
                  <option value="none">Not stated</option>
                  <option value="lte">≤</option>
                  <option value="gte">≥</option>
                </select>
                <input type="number" step="0.01" value={genSimgValue} disabled={genSimgOp === "none"}
                  onChange={(e) => setGenSimgValue(e.target.value)} placeholder="e.g. 3.5" />
              </div>
            </div>
          </div>
          <button className="btn-primary gen-submit" onClick={generateBarges}>Generate suggested plan</button>
        </div>
        <div className="note">Generation rule: Targeted Ni is a floor and target+0.2% is a ceiling — barges are never blended outside that range. Lowest-Ni domes are used first as diluent. A barge can draw from any number of existing-inventory contractors, but at most 2 production contractors (IMN-1–4). Domes below 1% or above 1.5% Ni are hard to source, so each is capped at 500 WMT per barge (spread across multiple barges instead). Targeted Si/Mg (when set) steers which domes are reached for first — it never excludes a dome, and the resulting barge's actual Si/Mg is shown on its row once generated. If there isn't enough on-spec ore left to reach a full barge within all active constraints, the barge is left incomplete rather than filled out of spec.</div>
      </section>

      {barges.length > 0 && (
        <section className="filter-row">
          {["all", "final", "draft", "exact", "excess", "deficit", "unplanned", "incomplete"].map((f) => (
            <button key={f} className={`chip ${filter === f ? "chip-active" : ""}`} onClick={() => setFilter(f)}>
              {f === "all" ? "All" : f === "final" ? "Finalized" : f === "draft" ? "Draft" : f === "exact" ? "On target" : f === "excess" ? "Above target" : f === "deficit" ? "Below target" : f === "unplanned" ? "Insufficient stock" : "Incomplete"}
            </button>
          ))}
        </section>
      )}

      <section className="barge-list">
        {barges.length === 0 && (
          <div className="glass empty-state">
            <Ship size={28} style={{ opacity: 0.4 }} />
            <p>No barges yet. Add one manually, or generate a suggested plan above.</p>
          </div>
        )}
        {filtered.map((b) => (
          <BargeRow key={b.no} barge={b} domesById={domesById} pool={pool}
            onUpdate={onUpdate} onImport={onUpdate} onFinalize={toggleFinalize} onRemove={removeBarge} onOpenInvoice={onOpenInvoice} onExportBarge={onExportBarge} onCheckStatus={onCheckStatus} onDataCommitted={onDataCommitted} isFeatureEnabled={isFeatureEnabled} />
        ))}
      </section>
    </div>
  );
}

/* ----------------------------- styles ----------------------------- */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@500;600&display=swap');

/* Browsers apply a default 8px margin to <body> unless explicitly reset — without this,
 * that default margin shows through as a white/light border around the app on a real
 * deployment (this only stayed hidden in the artifact preview, which resets it at the
 * platform level). Also sets html/body background to match .app's darkest tone, so
 * there's no flash of white before React mounts, and no gap is visible even if .app's
 * own min-height:100vh comes up a pixel short due to mobile browser chrome quirks. */
html, body { margin: 0; padding: 0; background: #070A10; }
#root { margin: 0; padding: 0; }

* { box-sizing: border-box; }
.app {
  min-height: 100vh;
  background: radial-gradient(circle at 20% 0%, #101826 0%, #070A10 55%, #050709 100%);
  color: #EAF0F6;
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  position: relative;
  overflow-x: hidden;
}
.bg-glow { position: fixed; border-radius: 999px; filter: blur(90px); opacity: .18; pointer-events: none; z-index: 0; }
.bg-glow-a { width: 420px; height: 420px; background: #E35F0C; top: -120px; right: -100px; }
.test-badge { position: fixed; top: 10px; right: 10px; background: rgba(248,113,113,.2); color: #F87171;
  padding: 6px 12px; border-radius: 20px; font-size: 10px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .05em; border: 1px solid rgba(248,113,113,.4); z-index: 999; pointer-events: none; }
.bg-glow-b { width: 380px; height: 380px; background: #C9A227; bottom: -140px; left: -80px; }

.topbar { position: relative; z-index: 2; display: flex; align-items: center; justify-content: space-between; padding: 20px 24px; flex-wrap: wrap; gap: 10px; }
.brand { display: flex; align-items: center; gap: 10px; }
.brand-mark { width: 36px; height: 36px; border-radius: 12px; display: flex; align-items: center; justify-content: center;
  background: linear-gradient(145deg, rgba(255,107,53,.25), rgba(255,107,53,.05)); border: 1px solid rgba(255,107,53,.35); color: #FF6B35; }
.brand-logo { width: 44px; height: 44px; display: block; flex-shrink: 0; filter: drop-shadow(0 2px 6px rgba(255,107,53,.15)); }
.brand-name { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 16px; letter-spacing: .04em; }
.brand-sub { font-size: 11px; color: #8A97A8; margin-top: 1px; }
.data-updated-badge { display: inline-flex; align-items: center; gap: 4px; font-size: 10px; color: #667080;
  margin-top: 4px; font-family: 'JetBrains Mono', monospace; }
.data-updated-badge svg { color: #E35F0C; opacity: .8; }

.nav-desktop-wrapper { display: flex; align-items: center; gap: 10px; }
.nav-desktop { display: flex; gap: 4px; background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.08); padding: 4px; border-radius: 14px; }
.navbtn { display: flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 10px; border: none; background: transparent; color: #8A97A8;
  font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 600; cursor: pointer; transition: all .15s ease; white-space: nowrap; }
.navbtn:hover { color: #EAF0F6; }
.navbtn-active { background: rgba(227,95,12,.14); color: #E35F0C; }
.mobile-nav-menu { display: none; position: relative; }
.hamburger-btn { display: flex; align-items: center; justify-content: center; width: 38px; height: 38px;
  background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.14); border-radius: 10px; color: #EAF0F6; cursor: pointer; flex-shrink: 0; }
.hamburger-btn:hover { background: rgba(255,255,255,.12); }
.mobile-nav-panel { position: absolute; top: calc(100% + 8px); right: 0; z-index: 60; width: 240px; max-width: 80vw;
  border-radius: 16px; overflow: hidden; padding: 0; }
.mobile-nav-user { padding: 14px 16px; border-bottom: 1px solid rgba(255,255,255,.08); }
.mobile-nav-username { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 13px; color: #EAF0F6; }
.mobile-nav-role { font-size: 10.5px; color: #8A97A8; margin-top: 2px; text-transform: uppercase; letter-spacing: .02em; }
.mobile-nav-list { display: flex; flex-direction: column; padding: 6px; }
.mobile-nav-item { display: flex; align-items: center; gap: 12px; width: 100%; padding: 11px 12px; background: transparent;
  border: none; border-radius: 9px; color: #B7C0CC; font-size: 13px; font-weight: 600; cursor: pointer; text-align: left; }
.mobile-nav-item:hover { background: rgba(255,255,255,.06); color: #EAF0F6; }
.mobile-nav-item-active { background: rgba(227,95,12,.14); color: #E35F0C; }
.mobile-nav-logout { display: flex; align-items: center; gap: 10px; width: 100%; padding: 12px 16px;
  background: rgba(248,113,113,.08); border: none; border-top: 1px solid rgba(255,255,255,.08); color: #F87171; font-size: 13px; font-weight: 700; cursor: pointer; }
.mobile-nav-logout:hover { background: rgba(248,113,113,.16); }

.btn-import { display: flex; align-items: center; gap: 6px; background: rgba(227,95,12,.14); border: 1px solid rgba(227,95,12,.35);
  color: #E35F0C; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 600; padding: 8px 12px; border-radius: 10px;
  cursor: pointer; transition: all .15s; }
.btn-import:hover { background: rgba(227,95,12,.22); border-color: #E35F0C; }

.content { position: relative; z-index: 1; max-width: 1240px; margin: 0 auto; padding: 8px 20px 40px; }
.stack { display: flex; flex-direction: column; gap: 18px; }

.glass {
  background: linear-gradient(150deg, rgba(255,255,255,.065), rgba(255,255,255,.015));
  border: 1px solid rgba(255,255,255,.09);
  backdrop-filter: blur(22px); -webkit-backdrop-filter: blur(22px);
  border-radius: 22px;
  box-shadow: 0 10px 34px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.05);
}

.hero { padding: 26px; display: flex; align-items: center; gap: 28px; flex-wrap: wrap; }
.hero-stats { flex: 1; min-width: 260px; }
.hero-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 20px; }
.hero-desc { color: #8A97A8; font-size: 13px; margin-top: 4px; margin-bottom: 16px; }

.kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }
.kpi-label { font-size: 11px; color: #8A97A8; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 4px; }
.kpi-value { font-family: 'JetBrains Mono', monospace; font-size: 19px; font-weight: 600; }
.kpi-unit { font-size: 11px; color: #8A97A8; margin-left: 4px; font-family: 'Inter', sans-serif; }
.kpi-good { color: #4ADE80; }
.kpi-warn { color: #FBBF24; }

.panel { padding: 20px; }
.panel-head { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; font-family: 'Space Grotesk', sans-serif; font-weight: 600; font-size: 14px; margin-bottom: 16px; color: #EAF0F6; }
.panel-head svg { color: #E35F0C; }
.panel-head-collapsible { padding: 12px 16px; margin: 0; background: linear-gradient(135deg, rgba(227,95,12,.05), rgba(255,255,255,.01)); 
  border-radius: 12px; border-bottom: 1px solid rgba(227,95,12,.15); 
  user-select: none; transition: background 150ms ease; display: flex; align-items: center; gap: 10px; }
.panel-head-collapsible:hover { background: linear-gradient(135deg, rgba(227,95,12,.1), rgba(255,255,255,.02)); }
.panel-head-collapsible svg:first-child { color: #E35F0C; flex-shrink: 0; }

.stacked-bar { display: flex; height: 14px; border-radius: 8px; overflow: hidden; background: #171E29; margin-bottom: 12px; }
.stacked-seg { transition: width .3s ease; }
.legend { display: flex; gap: 18px; flex-wrap: wrap; margin-top: 4px; }
.legend-item { display: flex; align-items: center; gap: 7px; font-size: 12px; color: #B7C0CC; }
.dot { width: 9px; height: 9px; border-radius: 50%; display: inline-block; flex-shrink: 0; margin-right: 6px; }
.note { font-size: 11.5px; color: #667080; margin-top: 12px; font-style: italic; }
.empty-note { font-size: 12px; color: #667080; font-style: italic; }

.btn-ghost { display: flex; align-items: center; gap: 6px; background: transparent; border: 1px solid rgba(255,255,255,.14);
  color: #B7C0CC; font-size: 12.5px; font-weight: 600; padding: 8px 12px; border-radius: 10px; cursor: pointer; transition: all .15s; }
.btn-ghost:hover { border-color: #E35F0C; color: #E35F0C; }
.btn-primary { display: flex; align-items: center; gap: 6px; background: rgba(227,95,12,.16); border: 1px solid rgba(227,95,12,.4);
  color: #E35F0C; font-size: 12.5px; font-weight: 700; padding: 9px 14px; border-radius: 10px; cursor: pointer; transition: all .15s; white-space: nowrap; }
.btn-primary:hover { background: rgba(227,95,12,.26); }
.btn-danger { background: rgba(248,113,113,.16); border-color: rgba(248,113,113,.45); color: #F87171; }
.btn-danger:hover { background: rgba(248,113,113,.28); }

.ring-wrap { position: relative; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.ring-outer { border-radius: 50%; display: flex; align-items: center; justify-content: center; }
.ring-inner { position: absolute; border-radius: 50%; background: #0A0F16; display: flex; flex-direction: column; align-items: center; justify-content: center; }
.ring-label { font-family: 'JetBrains Mono', monospace; font-weight: 600; font-size: 13px; }
.ring-sub { font-size: 8.5px; color: #8A97A8; margin-top: 1px; }

.summary-strip { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; padding: 18px 22px; }
.filter-row { display: flex; gap: 8px; flex-wrap: wrap; }
.chip { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.1); color: #B7C0CC; font-size: 12px; font-weight: 600;
  padding: 7px 13px; border-radius: 999px; cursor: pointer; transition: all .15s; }
.chip-active { background: rgba(227,95,12,.16); border-color: #E35F0C; color: #E35F0C; }

.badge { display: inline-flex; align-items: center; gap: 4px; font-size: 10.5px; font-weight: 700; padding: 3px 8px; border-radius: 999px; text-transform: uppercase; letter-spacing: .02em; white-space: nowrap; }
.badge-good { background: rgba(74,222,128,.14); color: #4ADE80; }
.badge-warn { background: rgba(251,191,36,.14); color: #FBBF24; }
.badge-bad { background: rgba(248,113,113,.14); color: #F87171; }

.banner { display: flex; gap: 14px; padding: 18px 20px; border: 1px solid rgba(251,191,36,.35); background: linear-gradient(150deg, rgba(251,191,36,.09), rgba(255,255,255,.02)); }
.banner-icon { color: #FBBF24; flex-shrink: 0; margin-top: 2px; }
.banner-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 14.5px; margin-bottom: 6px; color: #FDE68A; }
.banner-body { font-size: 12.5px; color: #D6DDE6; line-height: 1.55; }
.banner-body strong { color: #FBBF24; }

.table-wrap { overflow-x: auto; }
.table-wrap-tall { max-height: 520px; overflow-y: auto; }
.table-meta { font-size: 11px; color: #667080; margin-bottom: 8px; }
.filtered-summary { background: rgba(34,211,184,.05); border: 1px solid rgba(34,211,184,.18); border-radius: 12px;
  padding: 14px 16px; margin-bottom: 14px; }
.filtered-summary-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; color: #22D3B8; margin-bottom: 10px; }
.filtered-summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(90px, 1fr)); gap: 12px; }
.filtered-summary-item { display: flex; flex-direction: column; gap: 3px; }
.filtered-summary-item span { font-size: 10px; color: #8A97A8; text-transform: uppercase; letter-spacing: .02em; }
.filtered-summary-item strong { font-family: 'JetBrains Mono', monospace; font-size: 15px; font-weight: 700; color: #EAF0F6; }

@media (max-width: 640px) {
  .filtered-summary-grid { grid-template-columns: repeat(2, 1fr); }
}
.data-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.data-table th { position: sticky; top: 0; background: #0B1119; text-align: left; padding: 8px 10px; font-size: 10.5px; color: #8A97A8;
  text-transform: uppercase; letter-spacing: .03em; border-bottom: 1px solid rgba(255,255,255,.1); white-space: nowrap; }
.data-table th.sortable { cursor: pointer; user-select: none; }
.data-table th.sortable:hover { color: #E35F0C; }
.data-table td { padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,.05); color: #D6DDE6; white-space: nowrap; }
.cell-stockout { color: #E35F0C !important; font-weight: 600; }
.data-table tr:hover td { background: rgba(255,255,255,.02); }
.data-table-compact td, .data-table-compact th { padding: 6px 8px; font-size: 11.5px; }

.filter-toolbar { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
.search-box { display: flex; align-items: center; gap: 6px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12);
  border-radius: 10px; padding: 7px 10px; color: #8A97A8; flex: 1; min-width: 180px; }
.search-box input { background: transparent; border: none; outline: none; color: #EAF0F6; font-size: 12.5px; width: 100%; }
.filter-toolbar select { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 10px;
  color: #EAF0F6; font-size: 12.5px; padding: 7px 10px; }

.multiselect { position: relative; }
.multiselect-btn { display: flex; align-items: center; gap: 8px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12);
  border-radius: 10px; color: #EAF0F6; font-size: 12.5px; padding: 7px 12px; cursor: pointer; min-width: 160px; justify-content: space-between; }
.multiselect-btn:hover { border-color: rgba(227,95,12,.4); }
.multiselect-btn-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.multiselect-panel { position: absolute; top: calc(100% + 6px); left: 0; z-index: 20; width: 260px; max-width: 80vw;
  background: #171E29; border: 1px solid rgba(255,255,255,.15); border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,.5); overflow: hidden; }
.multiselect-actions { display: flex; gap: 6px; padding: 8px; border-bottom: 1px solid rgba(255,255,255,.08); }
.multiselect-actions button { flex: 1; background: rgba(227,95,12,.12); border: 1px solid rgba(227,95,12,.3); color: #E35F0C;
  border-radius: 7px; padding: 5px 8px; font-size: 11px; font-weight: 700; cursor: pointer; }
.multiselect-actions button:hover { background: rgba(227,95,12,.22); }
.multiselect-list { max-height: 240px; overflow-y: auto; padding: 6px; }
.multiselect-item { display: flex; align-items: center; gap: 8px; padding: 6px 8px; border-radius: 6px; font-size: 12px;
  color: #D6DDE6; cursor: pointer; }
.multiselect-item:hover { background: rgba(255,255,255,.06); }
.multiselect-item input { accent-color: #E35F0C; }

.dome-select { position: relative; flex: 1; min-width: 0; }
.dome-select-btn { width: 100%; display: flex; align-items: center; justify-content: space-between; gap: 8px;
  background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 10px;
  color: #EAF0F6; font-size: 12.5px; padding: 8px 12px; cursor: pointer; box-sizing: border-box; }
.dome-select-btn:hover { border-color: rgba(227,95,12,.4); }
.dome-select-btn-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; text-align: left; }
.dome-select-panel { position: absolute; top: calc(100% + 6px); left: 0; z-index: 20; width: 340px; max-width: 90vw;
  background: #171E29; border: 1px solid rgba(255,255,255,.15); border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,.5); overflow: hidden; }
.dome-select-search { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,.08); color: #8A97A8; }
.dome-select-search input { flex: 1; background: transparent; border: none; outline: none; color: #EAF0F6; font-size: 12.5px; }
.dome-select-list { max-height: 260px; overflow-y: auto; padding: 6px; }
.dome-select-item { display: flex; flex-direction: column; gap: 2px; padding: 7px 10px; border-radius: 7px; cursor: pointer; }
.dome-select-item:hover { background: rgba(227,95,12,.12); }
.dome-select-item-id { font-size: 12px; font-weight: 600; color: #EAF0F6; }
.dome-select-item-meta { font-size: 10.5px; color: #8A97A8; }
.dome-select-empty { padding: 14px; text-align: center; font-size: 12px; color: #667080; font-style: italic; }

.site-map-wrap { position: relative; height: 480px; border-radius: 16px; overflow: hidden;
  background: radial-gradient(circle at 30% 20%, #101C22 0%, #070B10 70%); border: 1px solid rgba(255,255,255,.06); }
.site-map-svg { width: 100%; height: 100%; display: block; user-select: none; -webkit-user-select: none; }
.site-map-svg.dragging { user-select: none; }
.poly-all { fill: rgba(227,95,12,.03); stroke: rgba(255,255,255,.14); stroke-width: 1.2; stroke-dasharray: 3 2; }
.poly-existing { fill: rgba(227,95,12,.07); stroke: rgba(227,95,12,.5); stroke-width: 1.4; }
.compass { position: absolute; top: 12px; right: 14px; font-family: 'JetBrains Mono', monospace; font-size: 11px;
  color: #667080; letter-spacing: .05em; }
.zoom-controls { position: absolute; top: 12px; left: 12px; display: flex; flex-direction: column; gap: 5px; z-index: 10; }
.zoom-btn { display: flex; align-items: center; justify-content: center; width: 32px; height: 32px;
  background: rgba(227,95,12,.12); border: 1px solid rgba(227,95,12,.35); color: #E35F0C;
  border-radius: 8px; font-size: 16px; font-weight: 700; cursor: pointer; transition: all .15s;
  font-family: 'JetBrains Mono', monospace; }
.zoom-btn:hover:not(:disabled) { background: rgba(227,95,12,.22); border-color: #E35F0C; }
.zoom-btn:disabled { opacity: 0.35; cursor: not-allowed; }
.map-tip-overlay { position: absolute; transform: translate(-50%, -120%); background: rgba(10,14,20,.96);
  border: 1px solid rgba(227,95,12,.4); border-radius: 10px; padding: 8px 11px; pointer-events: none;
  white-space: nowrap; box-shadow: 0 8px 22px rgba(0,0,0,.45); z-index: 5; }
.map-tip-overlay b { font-family: 'Space Grotesk', sans-serif; font-size: 12px; color: #EAF0F6; }
.map-tip-overlay div { font-size: 11px; color: #B7C0CC; margin-top: 2px; font-family: 'JetBrains Mono', monospace; }

.stat-card-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; }
.stat-card { background: rgba(255,255,255,.03); border: 1px solid rgba(255,255,255,.07); border-radius: 16px; padding: 16px; }
.stat-card-head { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
.stat-card-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 13.5px; }
.stat-card-tag { margin-left: auto; font-size: 9.5px; color: #667080; text-transform: uppercase; letter-spacing: .04em; border: 1px solid rgba(255,255,255,.1); padding: 2px 6px; border-radius: 999px; }
.stat-card-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
.stat-card-grid > div { display: flex; flex-direction: column; }
.stat-num { font-family: 'JetBrains Mono', monospace; font-size: 15px; font-weight: 600; color: #EAF0F6; }
.stat-lbl { font-size: 9.5px; color: #8A97A8; text-transform: uppercase; letter-spacing: .03em; margin-top: 2px; }

.add-blank-btn { margin-bottom: 16px; }
.gen-form { background: rgba(227,95,12,.05); border: 1px solid rgba(227,95,12,.15); border-radius: 12px; padding: 16px; }
.gen-form-title { font-family: 'Space Grotesk', sans-serif; font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .04em; color: #E35F0C; margin-bottom: 12px; }
.gen-form-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }
.gen-field { display: flex; flex-direction: column; gap: 6px; }
.gen-field label { font-size: 11px; color: #8A97A8; font-weight: 600; }
.gen-field input[type="number"] { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 8px;
  color: #EAF0F6; font-family: 'JetBrains Mono', monospace; font-size: 12.5px; padding: 7px 10px; width: 100%; box-sizing: border-box; }
.gen-source-checks { display: flex; flex-direction: column; gap: 6px; padding-top: 3px; }
.gen-simg-row { display: flex; gap: 6px; }
.gen-simg-row select { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 8px;
  color: #EAF0F6; font-size: 12.5px; padding: 7px 8px; width: 76px; flex-shrink: 0; }
.gen-simg-row input { flex: 1; min-width: 0; }
.gen-simg-row input:disabled { opacity: .4; }
.gen-check { display: flex; align-items: center; gap: 7px; font-size: 12px; color: #D6DDE6; cursor: pointer; }
.gen-check input { accent-color: #E35F0C; }
.gen-submit { margin-top: 16px; width: 100%; justify-content: center; }

.empty-state { padding: 40px 20px; display: flex; flex-direction: column; align-items: center; gap: 10px; color: #8A97A8; font-size: 13px; text-align: center; }

.barge-list { display: flex; flex-direction: column; gap: 8px; }
.barge-row { padding: 0; }
.barge-final { border-color: rgba(74,222,128,.35); }
.barge-row-top { display: flex; align-items: center; gap: 12px; padding: 12px 16px; cursor: pointer; flex-wrap: wrap; }
.barge-row-left { display: flex; align-items: center; gap: 8px; min-width: 120px; }
.chevron { transition: transform .15s; color: #8A97A8; }
.chevron-open { transform: rotate(180deg); }
.barge-no { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 13.5px; }
.lock-icon { color: #4ADE80; }
.date-input { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 8px; color: #EAF0F6;
  font-family: 'JetBrains Mono', monospace; font-size: 11.5px; padding: 5px 8px; }
.barge-total-wmt { font-family: 'JetBrains Mono', monospace; font-size: 12.5px; color: #B7C0CC; margin-left: auto; }
.barge-avg-ni { font-family: 'JetBrains Mono', monospace; font-size: 12.5px; color: #E35F0C; font-weight: 600; }
.barge-avg-simg { font-family: 'JetBrains Mono', monospace; font-size: 11.5px; color: #8A97A8; font-weight: 600;
  padding: 2px 8px; border: 1px solid rgba(255,255,255,.14); border-radius: 6px; }

.barge-row-detail { padding: 4px 16px 16px; border-top: 1px solid rgba(255,255,255,.07); }
.barge-info-row { display: flex; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; }
.info-field { display: flex; flex-direction: column; gap: 4px; flex: 0 1 auto; }
.info-field label { font-size: 10px; text-transform: uppercase; letter-spacing: .03em; color: #8A97A8; font-weight: 600; }
.info-input { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 8px; color: #EAF0F6;
  font-family: 'JetBrains Mono', monospace; font-size: 12px; padding: 7px 10px; min-width: 180px; }
.info-input:disabled { opacity: 0.5; cursor: not-allowed; }
.contractor-rollup { display: flex; gap: 8px; flex-wrap: wrap; margin: 10px 0 12px; }
.rollup-chip { display: flex; align-items: center; font-size: 10.5px; font-weight: 600; color: #B7C0CC;
  background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.08); border-radius: 999px; padding: 4px 9px; }
.cell-input { width: 80px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 6px;
  color: #EAF0F6; font-family: 'JetBrains Mono', monospace; font-size: 11.5px; padding: 4px 6px; }
.icon-btn { background: transparent; border: none; color: #F87171; cursor: pointer; padding: 4px; }
.icon-btn:hover { color: #FCA5A5; }

.add-dome-row { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }
.add-dome-row select { flex: 1; min-width: 200px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12);
  border-radius: 8px; color: #EAF0F6; font-size: 11.5px; padding: 6px 8px; }
.add-dome-row input { width: 90px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 8px;
  color: #EAF0F6; font-size: 11.5px; padding: 6px 8px; }

.import-barge-label { display: flex; align-items: center; gap: 6px; margin-top: 10px; font-size: 11px; color: #E35F0C;
  border: 1px dashed rgba(227,95,12,.4); border-radius: 8px; padding: 8px 10px; cursor: pointer; width: fit-content; }
.import-barge-label input { display: none; }

.barge-row-actions { margin-top: 12px; display: flex; justify-content: flex-end; gap: 10px; }
.btn-invoice { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 9px;
  background: linear-gradient(135deg, #E35F0C, #C94E08); border: none; color: #fff; font-size: 12px; font-weight: 700;
  cursor: pointer; transition: filter .15s; }
.btn-invoice:hover { filter: brightness(1.1); }
.btn-invoice-mini { display: inline-flex; align-items: center; gap: 5px; padding: 5px 10px; border-radius: 7px;
  background: rgba(227,95,12,.15); border: 1px solid rgba(227,95,12,.4); color: #E35F0C; font-size: 11px; font-weight: 700;
  cursor: pointer; margin-left: 8px; flex-shrink: 0; }
.btn-invoice-mini:hover { background: rgba(227,95,12,.25); }
.btn-export-mini { display: inline-flex; align-items: center; gap: 5px; padding: 5px 10px; border-radius: 7px;
  background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.18); color: #B7C0CC; font-size: 11px; font-weight: 700;
  cursor: pointer; margin-left: 8px; flex-shrink: 0; }
.btn-export-mini:hover { background: rgba(255,255,255,.12); }
.btn-status { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 9px;
  background: rgba(251,191,36,.1); border: 1px solid rgba(251,191,36,.3); color: #FBBF24; font-size: 12px; font-weight: 700;
  cursor: pointer; }
.btn-status:hover { background: rgba(251,191,36,.2); }
.btn-export { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 9px;
  background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.18); color: #D6DDE6; font-size: 12px; font-weight: 700;
  cursor: pointer; }
.btn-export:hover { background: rgba(255,255,255,.12); }
.btn-toggle { display: flex; align-items: center; gap: 6px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.15);
  color: #B7C0CC; font-size: 12px; font-weight: 600; padding: 7px 12px; border-radius: 9px; cursor: pointer; }
.btn-toggle-on { background: rgba(74,222,128,.14); border-color: rgba(74,222,128,.4); color: #4ADE80; }

.month-chart { display: flex; align-items: flex-end; gap: 6px; height: 140px; margin-bottom: 8px; }
.month-col { position: relative; flex: 1; display: flex; flex-direction: column; align-items: center; height: 100%; justify-content: flex-end; }
.chart-tooltip { position: absolute; bottom: calc(100% + 10px); left: 50%; transform: translateX(-50%); z-index: 30;
  background: rgba(15,19,26,.98); border: 1px solid rgba(255,255,255,.14); border-radius: 10px; padding: 10px 12px;
  min-width: 148px; box-shadow: 0 8px 24px rgba(0,0,0,.5); pointer-events: none; animation: tooltipFadeIn .12s ease-out; }
.chart-tooltip::after { content: ""; position: absolute; top: 100%; left: 50%; transform: translateX(-50%);
  border: 6px solid transparent; border-top-color: rgba(15,19,26,.98); }
@keyframes tooltipFadeIn { from { opacity: 0; transform: translateX(-50%) translateY(4px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
.chart-tooltip-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 12px; color: #EAF0F6;
  margin-bottom: 8px; padding-bottom: 8px; border-bottom: 1px solid rgba(255,255,255,.1); }
.chart-tooltip-row { display: flex; align-items: center; gap: 6px; font-size: 11px; color: #B7C0CC; margin-bottom: 4px; white-space: nowrap; }
.chart-tooltip-row:last-child { margin-bottom: 0; }
.chart-tooltip-row strong { margin-left: auto; color: #EAF0F6; font-family: 'JetBrains Mono', monospace; font-weight: 600; }
.chart-tooltip-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
.chart-tooltip-muted { padding-top: 4px; margin-top: 4px; border-top: 1px solid rgba(255,255,255,.08); color: #8A97A8; }
.month-col:first-child .chart-tooltip, .month-col:nth-child(2) .chart-tooltip { left: 0; transform: none; }
.month-col:first-child .chart-tooltip::after, .month-col:nth-child(2) .chart-tooltip::after { left: 20px; transform: none; }
.month-col:last-child .chart-tooltip, .month-col:nth-last-child(2) .chart-tooltip { left: auto; right: 0; transform: none; }
.month-col:last-child .chart-tooltip::after, .month-col:nth-last-child(2) .chart-tooltip::after { left: auto; right: 20px; transform: none; }
.month-bars { width: 100%; display: flex; flex-direction: column-reverse; align-items: center; height: 100px; justify-content: flex-start; }
.month-bar { width: 60%; border-radius: 3px 3px 0 0; }
.month-bar-shipped { background: #E35F0C; }
.month-bar-other { background: #C9A227; }
.month-name { font-size: 10px; color: #8A97A8; margin-top: 6px; }
.month-now { color: #E35F0C; font-weight: 700; }
.month-count { font-size: 9.5px; color: #667080; font-family: 'JetBrains Mono', monospace; }

.tracker-list { display: flex; flex-direction: column; gap: 4px; max-height: 420px; overflow-y: auto; }
.tracker-row { display: flex; align-items: center; gap: 12px; background: rgba(255,255,255,.02); border: 1px solid rgba(255,255,255,.06);
  border-radius: 10px; padding: 9px 12px; }
.tracker-no { font-family: 'JetBrains Mono', monospace; font-size: 12px; color: #8A97A8; width: 42px; }
.tracker-month { font-size: 11.5px; color: #B7C0CC; width: 90px; font-family: 'JetBrains Mono', monospace; }
.tracker-grade { font-family: 'JetBrains Mono', monospace; font-size: 12px; color: #8A97A8; width: 70px; margin-left: auto; }

.import-modal { position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 50; background: rgba(0,0,0,.5);
  display: flex; align-items: center; justify-content: center; padding: 20px; }
.import-panel { max-width: 420px; max-height: 80vh; overflow-y: auto; padding: 0; border-radius: 20px; }
.import-head { display: flex; align-items: center; gap: 10px; padding: 22px; border-bottom: 1px solid rgba(255,255,255,.1);
  font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 15px; position: relative; }
.import-close { position: absolute; right: 20px; top: 20px; background: transparent; border: none; color: #8A97A8; cursor: pointer; }
.import-body { padding: 20px; }
.import-body p { font-size: 12px; color: #B7C0CC; margin-bottom: 8px; }
.import-input-label { display: block; position: relative; margin: 16px 0; }
.import-input-label input { display: none; }
.import-input-label span { display: block; text-align: center; padding: 10px 16px; border: 2px dashed rgba(227,95,12,.4);
  border-radius: 10px; color: #E35F0C; font-size: 12px; font-weight: 600; cursor: pointer; }
.import-status { margin-top: 8px; padding: 8px 12px; background: rgba(74,222,128,.15); border: 1px solid rgba(74,222,128,.3);
  border-radius: 8px; color: #4ADE80; font-size: 11px; font-weight: 600; }

/* Contractor overview cards */
.contractor-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 16px; padding: 18px 0; }
.contractor-card { background: linear-gradient(135deg, rgba(227,95,12,.08), rgba(255,255,255,.02)); border: 1px solid rgba(227,95,12,.2);
  border-radius: 14px; padding: 16px; display: flex; flex-direction: column; gap: 10px; }
.contractor-card-prod { background: linear-gradient(135deg, rgba(59,130,246,.08), rgba(255,255,255,.02)); border: 1px solid rgba(59,130,246,.2); }
.card-contractor { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 13px; color: #E35F0C; text-transform: uppercase;
  letter-spacing: .02em; }
.card-contractor-prod { color: #3B82F6; }
.card-row { display: flex; justify-content: space-between; align-items: center; }
.card-label { font-size: 11px; color: #8A97A8; text-transform: uppercase; letter-spacing: .02em; font-weight: 600; }
.card-value { font-family: 'JetBrains Mono', monospace; font-size: 13px; font-weight: 600; color: #B7C0CC; }
.card-highlight { color: #4ADE80; }
.card-used { color: #FBBF24; }
.card-domes { border-top: 1px solid rgba(255,255,255,.08); padding-top: 10px; }

/* 2026 Production Targets */
.targets-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; padding: 18px 0; }
.target-card { background: linear-gradient(135deg, rgba(34,211,184,.07), rgba(255,255,255,.02)); border: 1px solid rgba(34,211,184,.2);
  border-radius: 14px; padding: 16px; display: flex; flex-direction: column; gap: 10px; }
.target-header { display: flex; justify-content: space-between; align-items: center; }
.target-contractor { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 13px; color: #EAF0F6;
  text-transform: uppercase; letter-spacing: .02em; }
.target-pct { font-family: 'JetBrains Mono', monospace; font-size: 13px; font-weight: 700; padding: 2px 8px; border-radius: 6px; }
.target-pct-good { color: #4ADE80; background: rgba(74,222,128,.12); }
.target-pct-warn { color: #FBBF24; background: rgba(251,191,36,.12); }
.target-pct-bad { color: #F87171; background: rgba(248,113,113,.12); }
.target-bar-track { height: 7px; border-radius: 4px; background: rgba(255,255,255,.08); overflow: hidden; }
.target-bar-fill { height: 100%; border-radius: 4px; background: linear-gradient(90deg, #22D3B8, #4ADE80); transition: width 300ms ease; }
.target-row { display: flex; justify-content: space-between; align-items: baseline; }
.target-label { font-size: 11px; color: #8A97A8; text-transform: uppercase; letter-spacing: .02em; font-weight: 600; }
.target-value { font-family: 'JetBrains Mono', monospace; font-size: 12.5px; font-weight: 600; color: #B7C0CC; }
.target-ni-target { font-size: 10.5px; color: #667080; font-weight: 400; }
.target-ni-badge { align-self: flex-start; font-size: 10.5px; font-weight: 700; padding: 3px 9px; border-radius: 6px;
  font-family: 'JetBrains Mono', monospace; }
.ni-badge-good { color: #4ADE80; background: rgba(74,222,128,.12); }
.ni-badge-warn { color: #FBBF24; background: rgba(251,191,36,.12); }
.target-unassayed-note { font-size: 10px; color: #8A97A8; font-style: italic; }

/* Barge status check modal */
.validation-modal { position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 50; background: rgba(0,0,0,.6);
  display: flex; align-items: center; justify-content: center; padding: 20px; }
.validation-panel { max-width: 560px; max-height: 80vh; overflow-y: auto; padding: 0; border-radius: 20px;
  border: 1px solid rgba(248,113,113,.3); }
.validation-head { display: flex; align-items: center; gap: 12px; padding: 22px; border-bottom: 1px solid rgba(248,113,113,.2);
  background: linear-gradient(135deg, rgba(248,113,113,.08), rgba(255,255,255,.01));
  font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 15px; color: #F87171; position: relative; }
.validation-close { position: absolute; right: 20px; top: 20px; background: transparent; border: none; color: #8A97A8; cursor: pointer; }
.validation-body { padding: 22px; }
.validation-body p { font-size: 12.5px; color: #D6DDE6; margin-bottom: 16px; line-height: 1.5; }
.violations-list { display: flex; flex-direction: column; gap: 10px; margin-bottom: 16px;
  padding: 12px; background: rgba(248,113,113,.08); border: 1px solid rgba(248,113,113,.15); border-radius: 10px; }
.violation-item { display: flex; flex-direction: column; gap: 4px; }
.violation-dome { font-family: 'JetBrains Mono', monospace; font-size: 12px; font-weight: 600; color: #F87171; }
.violation-detail { font-size: 11px; color: #B7C0CC; line-height: 1.4; }
.validation-actions { display: flex; gap: 10px; }
.status-section-title { display: flex; align-items: center; gap: 7px; font-family: 'Space Grotesk', sans-serif; font-weight: 700;
  font-size: 12.5px; text-transform: uppercase; letter-spacing: .02em; margin: 18px 0 8px; }
.status-section-title:first-child { margin-top: 0; }
.status-ok { display: flex; align-items: center; gap: 8px; padding: 12px; background: rgba(74,222,128,.08);
  border: 1px solid rgba(74,222,128,.2); border-radius: 10px; color: #4ADE80; font-size: 12px; font-weight: 600; margin-bottom: 12px; }
.lab-list { display: flex; flex-direction: column; gap: 10px; margin-bottom: 12px;
  padding: 12px; background: rgba(251,191,36,.08); border: 1px solid rgba(251,191,36,.18); border-radius: 10px; }
.lab-item { display: flex; flex-direction: column; gap: 4px; }
.lab-dome { font-family: 'JetBrains Mono', monospace; font-size: 12px; font-weight: 600; color: #FBBF24; }
.lab-detail { font-size: 11px; color: #B7C0CC; line-height: 1.4; }

@media (max-width: 900px) {
  .kpi-row { grid-template-columns: repeat(2, 1fr); }
  .summary-strip { grid-template-columns: repeat(2, 1fr); }
  .contractor-grid { grid-template-columns: repeat(2, 1fr); }
  .targets-grid { grid-template-columns: repeat(2, 1fr); }
  .gen-form-grid { grid-template-columns: repeat(2, 1fr); }
}

@media (max-width: 640px) {
  .nav-desktop, .btn-import, .btn-logout { display: none !important; }
  .mobile-nav-menu { display: block; }
  .topbar { padding: 16px 16px 4px; }
  .hero { flex-direction: column; align-items: flex-start; padding: 20px; }
  .kpi-row { grid-template-columns: repeat(2, 1fr); width: 100%; }
  .barge-row-top { gap: 8px; }
  .barge-total-wmt { margin-left: 0; }
  .tracker-month { display: none; }
  .import-panel { max-width: 90vw; }
  .validation-panel { max-width: 90vw; }
  .contractor-grid { grid-template-columns: 1fr; }
  .targets-grid { grid-template-columns: 1fr; }
  .gen-form-grid { grid-template-columns: 1fr; }
  .invoice-body { flex-direction: column; }
  .invoice-form { width: 100%; max-width: none; }
  .invoice-panel { max-width: 96vw; max-height: 92vh; }
  .invoice-action-bar { flex-direction: column; align-items: stretch; padding: 12px 16px; }
  .invoice-print-btn { max-width: none; }
  .btn-download-html { width: 100%; justify-content: center; }
  .invoice-action-hint { text-align: center; }
}

/* ============================================================
 * Invoice generator
 * ============================================================ */
.invoice-modal { position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 60; background: rgba(0,0,0,.65);
  display: flex; align-items: center; justify-content: center; padding: 16px; }
.invoice-panel { width: 1180px; max-width: 96vw; max-height: 94vh; overflow-y: auto; padding: 0; border-radius: 18px; }
.invoice-head { display: flex; align-items: center; gap: 10px; padding: 18px 22px; border-bottom: 1px solid rgba(255,255,255,.1);
  font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 15px; position: sticky; top: 0; z-index: 5;
  background: #171E29; border-radius: 18px 18px 0 0; color: #EAF0F6; }
.invoice-close { position: absolute; right: 18px; top: 16px; background: transparent; border: none; color: #8A97A8; cursor: pointer; padding: 6px; border-radius: 8px; }
.invoice-close:hover { background: rgba(255,255,255,.08); color: #EAF0F6; }
.invoice-body { display: flex; gap: 0; align-items: flex-start; }
.invoice-form { width: 380px; flex-shrink: 0; padding: 20px; border-right: 1px solid rgba(255,255,255,.08); }
.form-section-title { font-family: 'Space Grotesk', sans-serif; font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .04em; color: #E35F0C; margin: 18px 0 8px; }
.form-section-title:first-child { margin-top: 0; }
.form-grid { display: grid; grid-template-columns: 1fr; gap: 10px; }
.form-grid label { display: flex; flex-direction: column; gap: 4px; font-size: 11px; color: #8A97A8; font-weight: 600; }
.form-grid input, .form-grid textarea { background: #171E29; border: 1px solid rgba(255,255,255,.12); border-radius: 8px; padding: 8px 10px;
  color: #EAF0F6; font-size: 12.5px; font-family: 'JetBrains Mono', monospace; }
.form-grid textarea { resize: vertical; font-family: inherit; line-height: 1.5; }
.form-grid input:focus, .form-grid textarea:focus { outline: none; border-color: #E35F0C; }
.form-hint { font-size: 11px; color: #8A97A8; line-height: 1.5; margin: 6px 0 10px; }
.radio-row { display: flex; flex-direction: column; gap: 8px; }
.radio-pill { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-radius: 10px;
  border: 1px solid rgba(255,255,255,.12); font-size: 12.5px; font-weight: 600; color: #B7C0CC; cursor: pointer; transition: all .15s; }
.radio-pill input { accent-color: #E35F0C; }
.radio-pill-on { border-color: #E35F0C; background: rgba(227,95,12,.1); color: #EAF0F6; }
.settlement-note { font-size: 11.5px; color: #B7C0CC; background: rgba(227,95,12,.08); border: 1px solid rgba(227,95,12,.2);
  border-radius: 8px; padding: 10px 12px; margin-top: 4px; line-height: 1.5; }
.invoice-print-btn { width: 100%; max-width: 320px; justify-content: center; }
.invoice-action-bar { position: sticky; bottom: 0; z-index: 5; background: #171E29; border-top: 1px solid rgba(255,255,255,.1);
  padding: 14px 22px; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; border-radius: 0 0 18px 18px; }
.invoice-action-hint { font-size: 11px; color: #8A97A8; }
.btn-download-html { display: inline-flex; align-items: center; gap: 6px; padding: 10px 16px; border-radius: 9px;
  background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.18); color: #D6DDE6; font-size: 12.5px; font-weight: 700; cursor: pointer; }
.btn-download-html:hover { background: rgba(255,255,255,.12); }
.invoice-preview-wrap { flex: 1; padding: 24px; display: flex; justify-content: center; background: rgba(0,0,0,.2); min-height: 400px; }
.invoice-placeholder { color: #667080; font-size: 13px; font-style: italic; align-self: center; text-align: center; max-width: 320px; }

/* The invoice "paper" itself — matches the real Google-Docs HTML export point-for-point:
 * Arial font, 10pt body / 10pt bold labels, exact pt column widths, #999999 header band. */
.invoice-sheet { background: #ffffff; color: #000000; width: 522pt; max-width: 100%; padding: 24pt 20pt;
  font-family: Arial, sans-serif; font-size: 10pt; line-height: 1.3; box-shadow: 0 8px 30px rgba(0,0,0,.4); border-radius: 4px;
  box-sizing: border-box; }
.inv-header { margin-bottom: 8pt; }
.inv-logo { height: 56pt; }
.inv-bold { font-weight: 700; }
.inv-underline { text-decoration: underline; }
.inv-spacer { height: 8pt; }

.inv-top-table { border-collapse: collapse; margin-bottom: 6pt; table-layout: fixed; }
.inv-top-table td { border: 1pt solid #000; padding: 5pt; vertical-align: top; font-size: 10pt; }
.inv-top-table td.inv-kepada { text-align: left; border-top-color: #fff; border-left-color: #fff; border-bottom-color: #fff; border-right-color: #000; }
.inv-kepada div { margin: 0; }
.inv-cell-center { text-align: center; }
.inv-cell-center div { margin: 0; }

.inv-items-table { border-collapse: collapse; table-layout: fixed; margin-bottom: 6pt; }
.inv-items-table th, .inv-items-table td { border: 1pt solid #000; padding: 5pt; font-size: 10pt; vertical-align: top; }
.inv-items-table th { font-weight: 700; text-align: center; background: #999999; color: #000; }
.inv-left { text-align: left; }
.inv-center { text-align: center; }
.inv-right { text-align: right; white-space: nowrap; }
.inv-bank-cell { font-size: 10pt; line-height: 1.6; text-align: left; vertical-align: top; }
.inv-terbilang { font-size: 10pt; padding: 5pt; line-height: 1.6; text-align: left; }

.inv-sig-table { margin-left: auto; margin-right: 0; border-collapse: collapse; margin-top: 4pt; }
.inv-sig-box { border: 1pt solid #fff; padding: 8pt 14pt; text-align: center; min-width: 190pt; font-size: 10pt; }
.inv-sig-gap { height: 60pt; }
.inv-footer { margin-top: 10pt; text-align: right; font-size: 7pt; color: #000; line-height: 1.4; }

/* Export to PDF button (barging plan filter row) */
.btn-export-pdf { display: inline-flex; align-items: center; gap: 6px; margin-left: auto; padding: 7px 13px; border-radius: 9px;
  background: rgba(227,95,12,.12); border: 1px solid rgba(227,95,12,.35); color: #E35F0C; font-size: 12px; font-weight: 700; cursor: pointer; }
.btn-export-pdf:hover { background: rgba(227,95,12,.22); }

/* Barging Plan export report sheet */
.plan-sheet { background: #ffffff; color: #000000; width: 780px; max-width: 100%; padding: 10pt 16pt;
  font-family: Arial, sans-serif; font-size: 10pt; line-height: 1.3; box-shadow: 0 8px 30px rgba(0,0,0,.4); border-radius: 4px; box-sizing: border-box; }
.plan-header { display: flex; align-items: center; gap: 14pt; margin-bottom: 6pt; }
.plan-logo { height: 32pt; }
.plan-title { font-size: 14pt; font-weight: 700; }
.plan-subtitle { font-size: 8.5pt; color: #444; margin-top: 1pt; }
.plan-summary { display: flex; flex-wrap: wrap; gap: 16pt; padding: 6pt 0; border-top: 1pt solid #000; border-bottom: 1pt solid #000; margin-bottom: 8pt; }
.plan-summary-item { display: flex; flex-direction: column; gap: 1pt; }
.plan-summary-label { font-size: 7.5pt; text-transform: uppercase; letter-spacing: .03em; color: #555; }
.plan-summary-value { font-size: 10pt; font-weight: 700; }
.plan-table { width: 100%; border-collapse: collapse; margin-bottom: 8pt; table-layout: fixed; }
.plan-table th, .plan-table td { border: 1pt solid #000; padding: 2.5pt 6pt; font-size: 8.5pt; vertical-align: middle; line-height: 1.2; }
.plan-table th { font-weight: 700; text-align: center; background: #999999; }
.plan-center { text-align: center; }
.plan-right { text-align: right; }
.plan-muted { color: #888; font-style: italic; }
.plan-status-badge { display: inline-block; padding: 1pt 7pt; border-radius: 4pt; font-size: 8pt; font-weight: 700; border: 1pt solid #000; }
.plan-status-final { background: #d7f4de; }
.plan-status-draft { background: #fdf0d0; }
.plan-footer { margin-top: 4pt; text-align: right; font-size: 6.5pt; color: #000; line-height: 1.3; }
.plan-footer-bold { font-weight: 700; }

@page {
  size: A4 portrait;
  margin: 8mm;
}

@media print {
  /* Remove the rest of the dashboard from the page flow entirely (display:none, not
   * visibility:hidden) — visibility:hidden keeps the full-height layout intact even
   * though it's invisible, which makes the print engine paginate against the entire
   * dashboard's height instead of just the small modal, producing blank leading pages
   * before the real content shows up. display:none collapses that height to zero. */
  .bg-glow, .topbar, .content { display: none !important; }
  .no-print { display: none !important; }

  /* Neutralize the modal's on-screen chrome (dark backdrop, centering, fixed
   * positioning, scroll clipping) so the print-area flows naturally from the top
   * of the page and paginates across as many pages as it actually needs. */
  .invoice-modal { position: static !important; background: none !important; padding: 0 !important; display: block !important; }
  .invoice-panel { position: static !important; box-shadow: none !important; max-height: none !important;
    overflow: visible !important; width: auto !important; max-width: none !important; border-radius: 0 !important; }
  .invoice-body { display: block !important; }
  .invoice-preview-wrap { padding: 0 !important; background: none !important; display: block !important; min-height: 0 !important; }
  .print-area { box-shadow: none !important; border: none !important; width: 100% !important; }

  /* The sheet's on-screen width is set in px (screen unit) while its internal spacing
   * uses pt (print unit) — mixing the two can cause small rounding overflows once a
   * real print engine takes over layout. Forcing 100% width for print sidesteps that
   * entirely, letting the page's own @page margin box define the usable width. */
  .invoice-sheet, .plan-sheet { width: 100%; }

  /* Keep the closing address block intact as one unit rather than letting the print
   * engine split it (or push just a sliver of it) onto a trailing page. */
  .plan-footer, .inv-footer { page-break-inside: avoid; break-inside: avoid; }
  .plan-table tr { page-break-inside: avoid; break-inside: avoid; }
}

/* ======================== LOGIN SCREEN ======================== */
.login-container { position: fixed; inset: 0; background: radial-gradient(circle at 20% 0%, #101826 0%, #070A10 55%, #050709 100%);
  display: flex; align-items: center; justify-content: center; z-index: 1000; padding: 20px; overflow: hidden; }
.login-box { position: relative; z-index: 1; width: 100%; max-width: 760px; min-height: 460px; display: flex;
  border-radius: 24px; overflow: hidden; background: linear-gradient(150deg, rgba(255,255,255,.065), rgba(255,255,255,.015));
  border: 1px solid rgba(255,255,255,.09); box-shadow: 0 24px 60px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.05); }

/* Left: branding panel — same orange/gold geometry as the logo, scaled up as an ambient
 * background pattern, standing in for the reference's photo panel but built from
 * Integra's own visual identity instead of generic stock imagery. */
.login-panel-brand { position: relative; flex: 0 0 42%; padding: 36px 32px; display: flex; flex-direction: column;
  justify-content: space-between; background: linear-gradient(160deg, #171E29 0%, #0D1219 100%); overflow: hidden; }
.login-panel-brand-pattern { position: absolute; inset: -20% -30% auto auto; width: 340px; height: 340px;
  background: radial-gradient(circle, rgba(227,95,12,.18) 0%, transparent 70%); pointer-events: none; }
.login-panel-brand-pattern::after { content: ""; position: absolute; inset: 60% -40% -40% auto; width: 300px; height: 300px;
  background: radial-gradient(circle, rgba(191,178,42,.14) 0%, transparent 70%); }
.login-panel-brand-content { position: relative; z-index: 1; }
.login-panel-brand-name { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 22px; letter-spacing: .05em;
  color: #EAF0F6; margin-top: 18px; }
.login-panel-brand-tag { font-size: 12.5px; color: #8A97A8; margin-top: 8px; line-height: 1.5; }
.login-panel-brand-foot { position: relative; z-index: 1; font-size: 10.5px; color: #55606E; letter-spacing: .02em; }

.login-panel-form { flex: 1; padding: 44px 40px; display: flex; flex-direction: column; justify-content: center; }
.login-form-head { margin-bottom: 28px; }
.login-form-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 22px; color: #EAF0F6; }
.login-form-sub { font-size: 12.5px; color: #8A97A8; margin-top: 6px; }
.login-form { display: flex; flex-direction: column; gap: 16px; }
.form-group { display: flex; flex-direction: column; gap: 6px; }
.form-group label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .03em; color: #B7C0CC; }
.login-input { background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); border-radius: 10px;
  padding: 12px 14px; font-size: 14px; color: #EAF0F6; font-family: 'JetBrains Mono', monospace; box-sizing: border-box; }
.login-input:focus { outline: none; border-color: rgba(227,95,12,.5); box-shadow: 0 0 0 3px rgba(227,95,12,.12); }
.login-input::placeholder { color: #667080; }
.login-error { background: rgba(248,113,113,.12); border: 1px solid rgba(248,113,113,.3); border-radius: 8px;
  padding: 10px 12px; color: #FCA5A5; font-size: 12px; font-weight: 500; animation: shake .3s ease-in-out; }
@keyframes shake { 0%, 100% { transform: translateX(0); } 25% { transform: translateX(-4px); } 75% { transform: translateX(4px); } }
.login-button { display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #E35F0C, #C94E08);
  border: none; border-radius: 10px; padding: 13px 16px; font-size: 13px; font-weight: 700; color: #fff; cursor: pointer;
  text-transform: uppercase; letter-spacing: .05em; margin-top: 8px; font-family: 'Space Grotesk', sans-serif;
  box-shadow: 0 8px 20px rgba(227,95,12,.25); }
.login-button:hover { filter: brightness(1.08); }
.login-footer { text-align: center; margin-top: 24px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,.06); }
.login-footer p { font-size: 10.5px; color: #667080; margin: 0; letter-spacing: .02em; }

/* ======================== WELCOME SCREEN ======================== */
.welcome-overlay { position: fixed; inset: 0; background: radial-gradient(circle at 20% 0%, #101826 0%, #070A10 55%, #050709 100%);
  display: flex; align-items: center; justify-content: center; z-index: 1001; animation: fadeIn .3s ease-in-out; overflow: hidden; }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
.welcome-content { position: relative; z-index: 1; text-align: center; animation: slideUp .5s ease-out; }
@keyframes slideUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
.welcome-logo { display: flex; justify-content: center; margin-bottom: 20px; animation: bounce .6s ease-in-out; filter: drop-shadow(0 4px 16px rgba(227,95,12,.25)); }
@keyframes bounce { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-10px); } }
.welcome-title { font-family: 'Space Grotesk', sans-serif; font-size: 30px; font-weight: 700; color: #EAF0F6; margin: 0 0 8px; letter-spacing: -.5px; }
.welcome-subtitle { font-size: 13px; color: #B7C0CC; margin: 0 0 28px; }
.welcome-subtitle strong { color: #E35F0C; }
.welcome-spinner { width: 34px; height: 34px; border: 3px solid rgba(227,95,12,.2); border-top-color: #E35F0C;
  border-radius: 50%; animation: spin .8s linear infinite; margin: 0 auto; }
@keyframes spin { to { transform: rotate(360deg); } }

/* ======================== LOGOUT BUTTON ======================== */
.btn-logout { display: inline-flex; align-items: center; gap: 6px; padding: 8px 12px; border-radius: 9px;
  background: rgba(248,113,113,.1); border: 1px solid rgba(248,113,113,.3); color: #F87171; font-size: 12px; font-weight: 700; cursor: pointer; }
.btn-logout:hover { background: rgba(248,113,113,.18); }

/* ======================== FINANCIALS TAB ======================== */
.financial-cards-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
.financial-card { background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.1); border-radius: 16px;
  padding: 20px; display: flex; flex-direction: column; gap: 14px; }
.card-header { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 6px; font-size: 11.5px; color: #B7C0CC;
  font-weight: 600; text-transform: uppercase; letter-spacing: .03em; }
.card-unit { font-size: 10px; color: #8A97A8; font-weight: 500; }
.card-main-value { font-family: 'JetBrains Mono', monospace; font-size: 26px; font-weight: 700; color: #E35F0C; }
.card-meta { display: flex; flex-direction: column; gap: 8px; padding: 12px 0; border-top: 1px solid rgba(255,255,255,.08); border-bottom: 1px solid rgba(255,255,255,.08); }
.meta-item { display: flex; justify-content: space-between; align-items: center; font-size: 11px; }
.meta-label { color: #8A97A8; text-transform: uppercase; letter-spacing: .02em; font-weight: 600; }
.meta-value { font-family: 'JetBrains Mono', monospace; color: #B7C0CC; font-size: 12px; font-weight: 600; }
.meta-value.trend.up { color: #4ADE80; }
.meta-value.trend.down { color: #F87171; }
.card-actions { display: flex; gap: 8px; }
.btn-card-action { flex: 1; background: rgba(227,95,12,.12); border: 1px solid rgba(227,95,12,.3); border-radius: 8px;
  color: #E35F0C; font-size: 11px; font-weight: 700; padding: 8px 12px; cursor: pointer; text-transform: uppercase; letter-spacing: .02em; }
.btn-card-action:hover { background: rgba(227,95,12,.22); }

.royalty-table { background: rgba(255,255,255,.02); border: 1px solid rgba(255,255,255,.08); border-radius: 10px;
  overflow-x: auto; margin: 16px 0; }
.royalty-header, .royalty-row { display: grid; grid-template-columns: 70px 100px 90px 100px 110px 150px 90px; gap: 10px; align-items: center; }
.royalty-header { background: rgba(255,255,255,.06); padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,.08);
  font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .02em; color: #8A97A8; }
.royalty-row { padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,.04); font-size: 12px; color: #B7C0CC;
  font-family: 'JetBrains Mono', monospace; }
.royalty-row:last-child { border-bottom: none; }
.royalty-row .cell { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.royalty-row .cell.highlight { color: #E35F0C; font-weight: 700; }
.royalty-row .cell.status { color: #F5B841; text-transform: uppercase; font-size: 10px; font-weight: 600; }
.royalty-summary { display: flex; justify-content: space-between; align-items: center; background: rgba(74,222,128,.1);
  border: 1px solid rgba(74,222,128,.3); padding: 14px 16px; border-radius: 10px; font-size: 12px; color: #4ADE80; font-weight: 600; }
.royalty-summary .total { font-family: 'JetBrains Mono', monospace; font-size: 14px; font-weight: 700; }

/* ======================== LOGIN LOG TAB ======================== */
.log-table { background: rgba(255,255,255,.02); border: 1px solid rgba(255,255,255,.08); border-radius: 10px; overflow: hidden; }
.log-header-row, .log-row { display: grid; grid-template-columns: 1fr 1fr 100px; gap: 12px; align-items: center; }
.log-header-row { background: rgba(255,255,255,.06); padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,.08);
  font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .02em; color: #8A97A8; }
.log-row { padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,.04); font-size: 12px; color: #B7C0CC; font-family: 'JetBrains Mono', monospace; }
.log-row.log-success { background: rgba(74,222,128,.04); }
.log-row.log-failed { background: rgba(248,113,113,.04); }
.log-cell { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.log-cell-time { color: #8A97A8; }
.log-cell-user { color: #E35F0C; font-weight: 600; }
.log-cell-status { text-align: center; }
.status-badge { display: inline-block; padding: 3px 8px; border-radius: 6px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .02em; }
.badge-success { background: rgba(74,222,128,.2); color: #4ADE80; }
.badge-failed { background: rgba(248,113,113,.2); color: #FCA5A5; }
.log-empty { padding: 32px 16px; text-align: center; color: #8A97A8; font-size: 12px; }

/* Shared "mini stat card" pattern, used by both the Login Log summary and elsewhere.
 * Deliberately NOT named .stat-card — that class already exists for the Overview page's
 * summary cards, and reusing it would have silently overridden that unrelated UI. */
.mini-stats-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-top: 16px; }
.mini-stat-card { background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.1); border-radius: 10px; padding: 16px; text-align: center; }
.mini-stat-label { font-size: 10.5px; font-weight: 600; text-transform: uppercase; letter-spacing: .02em; color: #8A97A8; margin-bottom: 8px; }
.mini-stat-value { font-family: 'JetBrains Mono', monospace; font-size: 26px; font-weight: 700; color: #E35F0C; }
.mini-stat-value.mini-stat-success { color: #4ADE80; }
.mini-stat-value.mini-stat-failed { color: #F87171; }

/* ======================== SETTINGS TAB ======================== */
.settings-header { padding: 24px 20px 16px; border-bottom: 1px solid rgba(255,255,255,.1); }
.settings-header h2 { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 18px; color: #EAF0F6; margin: 0; }
.settings-section { padding: 20px; border-bottom: 1px solid rgba(255,255,255,.08); }
.settings-section:last-child { border-bottom: none; }
.dev-admin-section { background: rgba(248,113,113,.06); border-left: 3px solid rgba(248,113,113,.3); }
.dev-badge { background: rgba(248,113,113,.3) !important; color: #F87171 !important; }
.settings-section-header { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; }
.settings-section-header h3 { font-size: 13px; font-weight: 700; color: #22D3B8; text-transform: uppercase; letter-spacing: .03em; margin: 0; }
.section-badge { font-size: 9px; font-weight: 600; background: rgba(248,113,113,.2); color: #F87171;
  padding: 4px 8px; border-radius: 4px; text-transform: uppercase; letter-spacing: .02em; }
.settings-card { background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.1); border-radius: 10px;
  padding: 16px; display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; }
.settings-card .card-content h4 { font-size: 13px; font-weight: 600; color: #EAF0F6; margin: 0 0 4px; }
.settings-card .card-content p { font-size: 11px; color: #8A97A8; margin: 0; line-height: 1.4; }
.btn-settings-action { display: flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 600; padding: 8px 12px;
  border-radius: 8px; cursor: pointer; text-transform: uppercase; letter-spacing: .02em; white-space: nowrap; border: 1px solid; flex-shrink: 0; }
.btn-settings-action.btn-export { background: rgba(34,211,184,.15); border-color: rgba(34,211,184,.35); color: #22D3B8; }
.btn-settings-action.btn-export:hover { background: rgba(34,211,184,.25); border-color: rgba(34,211,184,.5); }
.btn-settings-action.btn-logout-action { background: rgba(248,113,113,.15); border-color: rgba(248,113,113,.35); color: #F87171; }
.btn-settings-action.btn-logout-action:hover { background: rgba(248,113,113,.25); border-color: rgba(248,113,113,.5); }
.btn-settings-action.btn-sync-sheets { background: rgba(74,222,128,.15); border-color: rgba(74,222,128,.35); color: #4ADE80; }
.btn-settings-action.btn-sync-sheets:hover { background: rgba(74,222,128,.25); border-color: rgba(74,222,128,.5); }
.sync-error-detail { display: flex; align-items: flex-start; gap: 8px; margin-top: 12px; padding: 10px 12px;
  background: rgba(248,113,113,.1); border: 1px solid rgba(248,113,113,.25); border-radius: 8px;
  color: #FCA5A5; font-size: 11.5px; line-height: 1.4; font-family: 'JetBrains Mono', monospace; }

/* ======================== LOADING REPORT ASSISTANT (dev account only) ======================== */
.loading-assistant-btn { position: fixed; bottom: 24px; right: 24px; width: 56px; height: 56px; border-radius: 50%;
  background: linear-gradient(135deg, #22D3B8 0%, #14B8A6 100%); border: none; color: #0A0E14; font-size: 10px; font-weight: 700;
  cursor: pointer; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px;
  box-shadow: 0 4px 12px rgba(34,211,184,.35); transition: transform .15s, box-shadow .15s; z-index: 400; }
.loading-assistant-btn:hover { transform: translateY(-3px); box-shadow: 0 8px 20px rgba(34,211,184,.45); }

/* Loading Progress section (Timeline tab) */
.loading-overall { margin-bottom: 18px; }
.loading-overall-numbers { display: flex; align-items: baseline; gap: 6px; margin-bottom: 8px; font-family: 'JetBrains Mono', monospace; }
.loading-overall-value { font-size: 22px; font-weight: 700; color: #22D3B8; }
.loading-overall-sep { color: #667080; font-size: 16px; }
.loading-overall-total { font-size: 15px; color: #B7C0CC; }
.loading-overall-pct { margin-left: auto; font-size: 15px; font-weight: 700; color: #EAF0F6; }
.loading-bar-track { width: 100%; height: 10px; background: rgba(255,255,255,.08); border-radius: 6px; overflow: hidden; }
.loading-bar-track-sm { height: 6px; margin: 6px 0; }
.loading-bar-fill { height: 100%; background: linear-gradient(90deg, #22D3B8, #14B8A6); border-radius: 6px; transition: width .3s; }
.loading-barge-list { display: flex; flex-direction: column; gap: 12px; padding-top: 14px; border-top: 1px solid rgba(255,255,255,.08); }
.loading-barge-row { padding: 12px 14px; background: rgba(255,255,255,.03); border: 1px solid rgba(255,255,255,.08); border-radius: 10px; }
.loading-barge-head { display: flex; align-items: center; gap: 10px; margin-bottom: 4px; }
.loading-barge-updated { margin-left: auto; font-size: 10.5px; color: #8A97A8; }
.loading-barge-meta { display: flex; justify-content: space-between; font-size: 11px; color: #B7C0CC; font-family: 'JetBrains Mono', monospace; margin-top: 4px; }
.loading-status-badge { display: inline-flex; align-items: center; padding: 2px 9px; border-radius: 6px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .02em; }
.loading-status-active { background: rgba(251,191,36,.18); color: #FBBF24; }
.loading-status-done { background: rgba(74,222,128,.18); color: #4ADE80; }

@media (max-width: 640px) {
  .loading-overall-numbers { flex-wrap: wrap; }
  .loading-overall-pct { margin-left: 0; }
  .loading-barge-meta { flex-direction: column; gap: 2px; }
}
.review-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 320px)); gap: 14px; margin-bottom: 8px; }

@media (max-width: 768px) {
  .loading-assistant-btn { bottom: 16px; right: 16px; width: 48px; height: 48px; font-size: 9px; }
  .review-grid { grid-template-columns: 1fr; }
}

@media (max-width: 768px) {
  .settings-card { flex-direction: column; align-items: flex-start; }
  .btn-settings-action { width: 100%; justify-content: center; }
}

@media (max-width: 900px) {
  .royalty-header, .royalty-row { grid-template-columns: 55px 80px 70px 80px 90px 120px 70px; gap: 6px; font-size: 11px; }
  .log-header-row, .log-row { grid-template-columns: 1fr 90px 80px; font-size: 11px; }
}
@media (max-width: 640px) {
  .login-box { flex-direction: column; max-width: 420px; min-height: 0; }
  .login-panel-brand { flex: none; padding: 28px 24px; flex-direction: row; align-items: center; gap: 14px; }
  .login-panel-brand-pattern { display: none; }
  .login-panel-brand-content { display: flex; align-items: center; gap: 14px; }
  .login-panel-brand-name { margin-top: 0; font-size: 18px; }
  .login-panel-brand-tag { display: none; }
  .login-panel-brand-foot { display: none; }
  .login-panel-form { padding: 28px 24px; }
  .login-form-title { font-size: 19px; }
  .btn-logout-label { display: none; }
  .financial-cards-row { grid-template-columns: 1fr; }
  .royalty-table { overflow-x: auto; }
  .royalty-header, .royalty-row { grid-template-columns: 50px 70px 60px 70px 80px 100px 60px; font-size: 10px; padding: 10px 12px; }
  .log-header-row, .log-row { grid-template-columns: 1fr 70px 70px; font-size: 10px; padding: 10px 12px; }
}
`;
