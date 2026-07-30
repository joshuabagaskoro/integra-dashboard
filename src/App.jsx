import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  Package, Ship, Calendar, TrendingUp, AlertTriangle, CheckCircle2,
  Layers, ChevronDown, RotateCcw, Gauge, Upload, X,
  Search, Plus, Trash2, Lock, Unlock, LayoutGrid, FileUp, MapPin, FileText, Printer, FileDown
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

/* ----------------------------- constants ----------------------------- */

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const TODAY = new Date("2026-07-27");

/* Bump this ONLY when DEFAULT_DOMES or DEFAULT_BARGES is edited with new Excel-sourced
 * data (stock updates, new barge plans). Do NOT change it for feature/UI/logic edits that
 * don't touch the underlying data — that's the whole point of this timestamp. */
const DATA_LAST_UPDATED = "2026-07-29";

const DEFAULT_SETTINGS = {
  bargeSize: 10000,
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
const DEFAULT_DOMES = [
{id:"DOME 01_1",contractor:"IMN",stock:18,initialStock:250,ni:1.33,fe:15.44,co:0.03,sio2:47.94,mgo:12.95,al2o3:0,simg:3.7,location:"",source:"inventory"},
{id:"DOME 04_1",contractor:"IMN/AMR",stock:195,initialStock:725,ni:1.29,fe:13.34,co:0.02,sio2:48.31,mgo:14.05,al2o3:0,simg:3.44,location:"",source:"inventory"},
{id:"DOME 05_1",contractor:"AMRUL",stock:188,ni:0.96,fe:10.67,co:0.01,sio2:50.58,mgo:20.25,al2o3:0,simg:2.5,location:"",source:"inventory"},
{id:"DOME 07_1",contractor:"AMRUL",stock:123,ni:1.07,fe:14,co:0.02,sio2:48.99,mgo:13.64,al2o3:0,simg:3.59,location:"",source:"inventory"},
{id:"DOME 08_1",contractor:"AMRUL",stock:52,initialStock:600,ni:1.05,fe:14.11,co:0.03,sio2:51.7,mgo:14.4,al2o3:0,simg:3.59,location:"",source:"inventory"},
{id:"DOME 09_1",contractor:"AMRUL",stock:56,ni:1.37,fe:14.95,co:0.03,sio2:51.8,mgo:11.84,al2o3:0,simg:4.38,location:"",source:"inventory"},
{id:"DOME 10_1",contractor:"AMRUL",stock:53,ni:1.21,fe:15.86,co:0.03,sio2:51.62,mgo:10.33,al2o3:0,simg:5,location:"",source:"inventory"},
{id:"DOME 11_1",contractor:"AMRUL",stock:94,ni:0.96,fe:14.03,co:0.02,sio2:51.52,mgo:12.45,al2o3:0,simg:4.14,location:"",source:"inventory"},
{id:"DOME 12_1",contractor:"AMRUL",stock:115,ni:0.88,fe:12.62,co:0.02,sio2:51.81,mgo:15.9,al2o3:0,simg:3.26,location:"",source:"inventory"},
{id:"DOME 13_1",contractor:"AMRUL",stock:46,ni:1.01,fe:13.5,co:0.02,sio2:51.14,mgo:14.72,al2o3:0,simg:3.47,location:"",source:"inventory"},
{id:"DOME 14_1",contractor:"AMRUL",stock:158,ni:0.74,fe:13.43,co:0.02,sio2:53.93,mgo:11.22,al2o3:0,simg:4.81,location:"",source:"inventory"},
{id:"DOME 17_1",contractor:"AMRUL",stock:208,ni:1.29,fe:12.15,co:0.02,sio2:50.65,mgo:17.7,al2o3:0,simg:2.86,location:"",source:"inventory"},
{id:"DOME 18_1",contractor:"AMRUL",stock:200,ni:0.74,fe:15.92,co:0.02,sio2:45.39,mgo:14.08,al2o3:0,simg:3.22,location:"",source:"inventory"},
{id:"DOME 19_1",contractor:"AMRUL",stock:0,initialStock:292,ni:0.95,fe:18.11,co:0.03,sio2:46.07,mgo:10.08,al2o3:0,simg:4.57,location:"",source:"inventory"},
{id:"DOME 20_1",contractor:"AMRUL",stock:120,ni:1.1,fe:15.93,co:0.02,sio2:43.84,mgo:11.04,al2o3:0,simg:3.97,location:"",source:"inventory"},
{id:"DOME 21_1",contractor:"AMRUL",stock:41,ni:0.5,fe:12.5,co:0.02,sio2:47.12,mgo:18.75,al2o3:0,simg:2.51,location:"",source:"inventory"},
{id:"DOME 22_1",contractor:"AMRUL",stock:94,ni:0.87,fe:16.49,co:0.03,sio2:44.05,mgo:11.39,al2o3:0,simg:3.87,location:"",source:"inventory"},
{id:"DOME 23_1",contractor:"AMRUL",stock:0,initialStock:311,ni:1,fe:14.23,co:0.03,sio2:48.52,mgo:15.81,al2o3:0,simg:3.07,location:"",source:"inventory"},
{id:"DOME 24_1",contractor:"AMRUL",stock:23,ni:1.19,fe:12.94,co:0.02,sio2:49.82,mgo:16.67,al2o3:0,simg:2.99,location:"",source:"inventory"},
{id:"DOME 25_1",contractor:"AMRUL",stock:2073,ni:0.62,fe:12.64,co:0.03,sio2:49.85,mgo:17.79,al2o3:0,simg:2.8,location:"",source:"inventory"},
{id:"DOME 26_1",contractor:"AMRUL",stock:450,ni:0.98,fe:14.79,co:0.03,sio2:49.28,mgo:11.95,al2o3:0,simg:4.12,location:"",source:"inventory"},
{id:"DOME 27_1",contractor:"AMRUL",stock:2693,ni:0.83,fe:13.77,co:0.02,sio2:50.54,mgo:13.91,al2o3:0,simg:3.63,location:"",source:"inventory"},
{id:"DOME 29_1",contractor:"AMRUL",stock:633,ni:0.77,fe:15.84,co:0.04,sio2:49.26,mgo:10.15,al2o3:0,simg:4.85,location:"",source:"inventory"},
{id:"DOME 30_1",contractor:"AMRUL",stock:0,initialStock:530,ni:1.09,fe:14.8,co:0.03,sio2:45.86,mgo:15.99,al2o3:0,simg:2.87,location:"",source:"inventory"},
{id:"DOME 32_1",contractor:"AMRUL",stock:665,ni:0.66,fe:10.87,co:0.02,sio2:45,mgo:24.04,al2o3:0,simg:1.87,location:"",source:"inventory"},
{id:"DOME 38_1",contractor:"TANO",stock:134,ni:2.26,fe:7.45,co:0.01,sio2:40.54,mgo:30.81,al2o3:0,simg:1.32,location:"",source:"inventory"},
{id:"DOME 45_1",contractor:"TANO",stock:369,ni:1.57,fe:8.05,co:0.01,sio2:40.92,mgo:29.73,al2o3:0,simg:1.38,location:"",source:"inventory"},
{id:"DOME 46_1",contractor:"TANO",stock:182,ni:1.41,fe:15.47,co:0.03,sio2:47.31,mgo:13.72,al2o3:0,simg:3.45,location:"",source:"inventory"},
{id:"DOME 48_1",contractor:"AMRUL",stock:309,ni:0.91,fe:13.87,co:0.02,sio2:48.53,mgo:14.22,al2o3:0,simg:3.41,location:"",source:"inventory"},
{id:"DOME 49_1",contractor:"AMRUL",stock:22,ni:1.29,fe:16.77,co:0.03,sio2:49.24,mgo:8.68,al2o3:0,simg:5.67,location:"",source:"inventory"},
{id:"DOME 50_1",contractor:"AMRUL",stock:8,ni:1.13,fe:13.47,co:0.03,sio2:48.83,mgo:14.39,al2o3:0,simg:3.39,location:"",source:"inventory"},
{id:"DOME 51_1",contractor:"AMRUL",stock:61,ni:1.69,fe:17.06,co:0.03,sio2:43.11,mgo:12.85,al2o3:0,simg:3.35,location:"",source:"inventory"},
{id:"DOME 52_1",contractor:"AMRUL",stock:23,ni:1.17,fe:15.05,co:0.03,sio2:51.85,mgo:12.07,al2o3:0,simg:4.3,location:"",source:"inventory"},
{id:"DOME 57_1",contractor:"MAJID",stock:159,ni:0.92,fe:14.59,co:0.03,sio2:42.25,mgo:16.6,al2o3:0,simg:2.55,location:"",source:"inventory"},
{id:"DOME 58_1",contractor:"MAJID",stock:224,ni:0.95,fe:15.54,co:0.03,sio2:44.27,mgo:17.24,al2o3:0,simg:2.57,location:"",source:"inventory"},
{id:"MAJID 05_1",contractor:"MAJID",stock:31,ni:1.08,fe:15.16,co:0.03,sio2:42.8,mgo:18.38,al2o3:0,simg:2.33,location:"",source:"inventory"},
{id:"DOME 64",contractor:"IMN",stock:181,ni:1.26,fe:18.57,co:0.04,sio2:34.58,mgo:16.44,al2o3:0,simg:2.1,location:"",source:"inventory"},
{id:"DOME 66",contractor:"TANO",stock:0,initialStock:259,ni:1.55,fe:9.81,co:0.02,sio2:39.38,mgo:26.68,al2o3:0,simg:1.48,location:"",source:"inventory"},
{id:"DOME 67",contractor:"TANO",stock:314,ni:0.74,fe:6.74,co:0.01,sio2:40.7,mgo:31.83,al2o3:0,simg:1.28,location:"",source:"inventory"},
{id:"DOME 68",contractor:"ALI M",stock:444,ni:0.98,fe:10.51,co:0.02,sio2:28.41,mgo:15.24,al2o3:0,simg:1.86,location:"",source:"inventory"},
{id:"DOME 69",contractor:"TANO",stock:77,ni:1.16,fe:7.8,co:0.02,sio2:41.59,mgo:27.78,al2o3:0,simg:1.5,location:"",source:"inventory"},
{id:"DOME 70",contractor:"ALI M",stock:169,ni:0.77,fe:8.14,co:0.01,sio2:45.01,mgo:25.54,al2o3:0,simg:1.76,location:"",source:"inventory"},
{id:"DOME 71",contractor:"TANO",stock:0,initialStock:486,ni:1.36,fe:9.73,co:0.02,sio2:40.15,mgo:26.18,al2o3:0,simg:1.53,location:"",source:"inventory"},
{id:"DOME 72",contractor:"TANO",stock:3,initialStock:262,ni:1.89,fe:10.75,co:0.02,sio2:37.27,mgo:27.85,al2o3:0,simg:1.34,location:"",source:"inventory"},
{id:"DOME 73",contractor:"TANO",stock:149,ni:0.79,fe:7.43,co:0.01,sio2:38.35,mgo:31.59,al2o3:0,simg:1.21,location:"",source:"inventory"},
{id:"DOME 78",contractor:"TANO",stock:232,ni:1.85,fe:12.34,co:0.03,sio2:36.19,mgo:22.42,al2o3:0,simg:1.61,location:"",source:"inventory"},
{id:"DOME 92",contractor:"MAJID",stock:2731,ni:0.69,fe:17.46,co:0.03,sio2:39.66,mgo:4.3,al2o3:0,simg:9.22,location:"",source:"inventory"},
{id:"DOME 93",contractor:"MONO",stock:2197,ni:1.08,fe:15.19,co:0.02,sio2:46.93,mgo:14.39,al2o3:0,simg:3.26,location:"",source:"inventory"},
{id:"DOME 246",contractor:"MAJID",stock:87,ni:1.14,fe:15.67,co:0.03,sio2:40.37,mgo:18.44,al2o3:0,simg:2.19,location:"",source:"inventory"},
{id:"DOME 247",contractor:"MAJID",stock:94,ni:1.12,fe:17.76,co:0.03,sio2:40.31,mgo:16.57,al2o3:0,simg:2.43,location:"",source:"inventory"},
{id:"DOME 96",contractor:"MAJID",stock:0,initialStock:616,ni:1.01,fe:7.93,co:0.01,sio2:45.81,mgo:27.01,al2o3:0,simg:1.7,location:"",source:"inventory"},
{id:"DOME 98",contractor:"MAJID",stock:0,initialStock:617,ni:0.97,fe:7.24,co:0.01,sio2:44.73,mgo:27.17,al2o3:0,simg:1.65,location:"",source:"inventory"},
{id:"DOME 99",contractor:"HAIRIL",stock:241,ni:1.96,fe:18.43,co:0.04,sio2:42.73,mgo:10.61,al2o3:0,simg:4.03,location:"",source:"inventory"},
{id:"DOME 100",contractor:"HAIRIL",stock:89,initialStock:290,ni:1.96,fe:19,co:0.04,sio2:42.44,mgo:9.14,al2o3:0,simg:4.64,location:"",source:"inventory"},
{id:"DOME 101",contractor:"HAIRIL",stock:14,initialStock:247,ni:1.97,fe:20.37,co:0.04,sio2:43.46,mgo:8.77,al2o3:0,simg:4.96,location:"",source:"inventory"},
{id:"DOME 102",contractor:"MAJID",stock:24,initialStock:90,ni:1.21,fe:7.41,co:0.01,sio2:41.78,mgo:31.12,al2o3:0,simg:1.34,location:"",source:"inventory"},
{id:"DOME 103",contractor:"MAJID",stock:0,initialStock:166,ni:0.98,fe:7.74,co:0.01,sio2:44.97,mgo:26.48,al2o3:0,simg:1.7,location:"",source:"inventory"},
{id:"DOME 106",contractor:"MAJID",stock:200,initialStock:310,ni:1.1,fe:6.4,co:0.01,sio2:40.35,mgo:32.85,al2o3:0,simg:1.23,location:"",source:"inventory"},
{id:"DOME 109",contractor:"MAJID",stock:0,initialStock:433,ni:1.07,fe:6.61,co:0.01,sio2:41.4,mgo:32.99,al2o3:0,simg:1.25,location:"",source:"inventory"},
{id:"DOME 110",contractor:"MAJID",stock:103,initialStock:186,ni:1.19,fe:6.82,co:0.01,sio2:41.27,mgo:32.34,al2o3:0,simg:1.28,location:"",source:"inventory"},
{id:"DOME 111",contractor:"SURIADIN",stock:479,ni:0.51,fe:11.55,co:0.02,sio2:47.98,mgo:19.77,al2o3:0,simg:2.43,location:"",source:"inventory"},
{id:"DOME 113",contractor:"SURIADIN",stock:477,ni:1.09,fe:22.95,co:0.06,sio2:16.96,mgo:5.19,al2o3:0,simg:3.27,location:"",source:"inventory"},
{id:"DOME 114",contractor:"SURIADIN",stock:285,ni:0.47,fe:8.49,co:0.01,sio2:36.05,mgo:15.72,al2o3:0,simg:2.29,location:"",source:"inventory"},
{id:"DOME 116",contractor:"AMRUL",stock:0,initialStock:496,ni:1.65,fe:18.04,co:0.04,sio2:41.34,mgo:11.13,al2o3:0,simg:3.71,location:"",source:"inventory"},
{id:"DOME 119",contractor:"SURIADIN",stock:0,initialStock:932,ni:1.2,fe:16.64,co:0.06,sio2:42.03,mgo:15.7,al2o3:0,simg:2.68,location:"",source:"inventory"},
{id:"DOME 121",contractor:"RAMLAN",stock:691,ni:0.96,fe:29.2,co:0.08,sio2:10.92,mgo:4.52,al2o3:0,simg:2.42,location:"",source:"inventory"},
{id:"DOME 122",contractor:"RAMLAN",stock:166,ni:0.96,fe:10.25,co:0.02,sio2:29.57,mgo:20.31,al2o3:0,simg:1.46,location:"",source:"inventory"},
{id:"DOME 123",contractor:"AKAS",stock:38,initialStock:1354,ni:1.22,fe:14.82,co:0.03,sio2:31.06,mgo:12.25,al2o3:0,simg:2.54,location:"",source:"inventory"},
{id:"DOME 124",contractor:"AKAS",stock:321,ni:0.92,fe:7.65,co:0.01,sio2:30.17,mgo:18.94,al2o3:0,simg:1.59,location:"",source:"inventory"},
{id:"DOME 127",contractor:"HAIRIL",stock:0,initialStock:315,ni:1.59,fe:14.01,co:0.03,sio2:42.15,mgo:18.46,al2o3:0,simg:2.28,location:"",source:"inventory"},
{id:"DOME 128",contractor:"HAIRIL",stock:327,initialStock:1186,ni:1.88,fe:16.43,co:0.03,sio2:43.7,mgo:13.83,al2o3:0,simg:3.16,location:"",source:"inventory"},
{id:"DOME 129",contractor:"AMRUL",stock:0,initialStock:1604,ni:1.39,fe:13.97,co:0.03,sio2:45.43,mgo:17.65,al2o3:0,simg:2.57,location:"",source:"inventory"},
{id:"DOME 130",contractor:"HAIRIL",stock:0,initialStock:525,ni:1.67,fe:12.97,co:0.02,sio2:43.45,mgo:17.61,al2o3:0,simg:2.47,location:"",source:"inventory"},
{id:"DOME 131",contractor:"IMN",stock:715,ni:0.69,fe:7.51,co:0.02,sio2:20.08,mgo:8.62,al2o3:0,simg:2.33,location:"",source:"inventory"},
{id:"DOME 132",contractor:"HAIRIL",stock:102,initialStock:340,ni:1.08,fe:14.21,co:0.03,sio2:45.01,mgo:12.81,al2o3:0,simg:3.51,location:"",source:"inventory"},
{id:"DOME 133",contractor:"HAIRIL",stock:4,initialStock:302,ni:1.74,fe:19.98,co:0.04,sio2:42.78,mgo:7.99,al2o3:0,simg:5.35,location:"",source:"inventory"},
{id:"DOME 134",contractor:"HAIRIL",stock:0,initialStock:205,ni:1.78,fe:20.25,co:0.04,sio2:43.14,mgo:7.6,al2o3:0,simg:5.68,location:"",source:"inventory"},
{id:"DOME 135",contractor:"HAIRIL",stock:0,initialStock:566,ni:2,fe:15.85,co:0.03,sio2:44.99,mgo:12.55,al2o3:0,simg:3.58,location:"",source:"inventory"},
{id:"DOME 136",contractor:"HAIRIL",stock:0,initialStock:483,ni:1.71,fe:14.35,co:0.03,sio2:43.44,mgo:16.93,al2o3:0,simg:2.57,location:"",source:"inventory"},
{id:"DOME 137",contractor:"HAIRIL",stock:0,initialStock:175,ni:1.02,fe:14.82,co:0.03,sio2:23.27,mgo:5.87,al2o3:0,simg:3.96,location:"",source:"inventory"},
{id:"DOME 138",contractor:"HAIRIL",stock:0,initialStock:192,ni:1.13,fe:15.86,co:0.03,sio2:23.76,mgo:6.69,al2o3:0,simg:3.55,location:"",source:"inventory"},
{id:"DOME 139",contractor:"HAIRIL",stock:23,initialStock:145,ni:1.33,fe:30.27,co:0.12,sio2:11.89,mgo:5.98,al2o3:0,simg:1.99,location:"",source:"inventory"},
{id:"DOME 140",contractor:"HAIRIL",stock:0,initialStock:315,ni:1.5,fe:31.48,co:0.11,sio2:10.96,mgo:4.65,al2o3:0,simg:2.36,location:"",source:"inventory"},
{id:"DOME 141",contractor:"AMRUL",stock:80,ni:1.69,fe:27.03,co:0.05,sio2:34.16,mgo:4.92,al2o3:0,simg:6.94,location:"",source:"inventory"},
{id:"DOME 145",contractor:"AMRUL",stock:0,initialStock:937,ni:1.74,fe:21.28,co:0.04,sio2:38.39,mgo:9.45,al2o3:0,simg:4.06,location:"",source:"inventory"},
{id:"DOME 150_1",contractor:"AMRUL",stock:0,initialStock:140,ni:2.27,fe:18.76,co:0.04,sio2:41.33,mgo:11.14,al2o3:0,simg:3.71,location:"",source:"inventory"},
{id:"DOME 153_1",contractor:"AMRUL",stock:389,initialStock:627,ni:1.75,fe:23.35,co:0.05,sio2:37.47,mgo:8.1,al2o3:0,simg:4.63,location:"",source:"inventory"},
{id:"DOME 154_1",contractor:"AMRUL",stock:0,initialStock:1355,ni:1.71,fe:19.35,co:0.04,sio2:41.71,mgo:10.72,al2o3:0,simg:3.89,location:"",source:"inventory"},
{id:"DOME 155_1",contractor:"AMRUL",stock:0,initialStock:298,ni:1.87,fe:13.93,co:0.03,sio2:45.57,mgo:17.19,al2o3:0,simg:2.65,location:"",source:"inventory"},
{id:"DOME 160_1",contractor:"AMRUL",stock:283,ni:0.84,fe:13.3,co:0.02,sio2:41.12,mgo:11.88,al2o3:0,simg:3.46,location:"",source:"inventory"},
{id:"DOME 161_1",contractor:"AKAS",stock:2671,ni:0.69,fe:11.16,co:0.02,sio2:46.57,mgo:21.64,al2o3:0,simg:2.15,location:"",source:"inventory"},
{id:"DOME 162_1",contractor:"AKAS",stock:294,ni:0.69,fe:11.83,co:0.02,sio2:44.3,mgo:19.63,al2o3:0,simg:2.26,location:"",source:"inventory"},
{id:"DOME 163_1",contractor:"AKAS",stock:105,initialStock:703,ni:1.28,fe:14.66,co:0.03,sio2:28.95,mgo:12.29,al2o3:0,simg:2.36,location:"",source:"inventory"},
{id:"DOME 164_1",contractor:"IMN",stock:280,ni:0.83,fe:30.05,co:0.07,sio2:11.82,mgo:3.59,al2o3:0,simg:3.29,location:"",source:"inventory"},
{id:"DOME 165_1",contractor:"SURIADIN",stock:0,initialStock:350,ni:1.01,fe:31.76,co:0.09,sio2:9.81,mgo:3.74,al2o3:0,simg:2.62,location:"",source:"inventory"},
{id:"DOME 166_1",contractor:"SURIADIN",stock:1835,ni:0.92,fe:21.36,co:0.05,sio2:13.7,mgo:4.98,al2o3:0,simg:2.75,location:"",source:"inventory"},
{id:"DOME 168_1",contractor:"SURIADIN",stock:0,initialStock:1489,ni:1.3,fe:21.94,co:0.07,sio2:13.73,mgo:5.88,al2o3:0,simg:2.34,location:"",source:"inventory"},
{id:"DOME 169_1",contractor:"AMRUL",stock:382,ni:0.81,fe:14.75,co:0.02,sio2:46.02,mgo:14.43,al2o3:0,simg:3.19,location:"",source:"inventory"},
{id:"DOME 170_1",contractor:"H. EWA",stock:24,ni:1.04,fe:8.88,co:0.02,sio2:33.83,mgo:26.23,al2o3:0,simg:1.29,location:"",source:"inventory"},
{id:"DOME 171_1",contractor:"H. EWA",stock:94,ni:1.14,fe:13.26,co:0.02,sio2:27.29,mgo:15.54,al2o3:0,simg:1.76,location:"",source:"inventory"},
{id:"DOME 176_1",contractor:"NURIS",stock:462,ni:0.91,fe:21.21,co:0.04,sio2:13.05,mgo:2.93,al2o3:0,simg:4.45,location:"",source:"inventory"},
{id:"DOME 177_1",contractor:"NURIS",stock:36,ni:0.97,fe:13.23,co:0.02,sio2:25.55,mgo:10.56,al2o3:0,simg:2.42,location:"",source:"inventory"},
{id:"DOME 178_1",contractor:"NURIS",stock:0,initialStock:239,ni:1.33,fe:19.11,co:0.03,sio2:35.4,mgo:5.93,al2o3:0,simg:5.97,location:"",source:"inventory"},
{id:"DOME 180_1",contractor:"NURIS",stock:0,initialStock:159,ni:1.23,fe:18.68,co:0.02,sio2:36.16,mgo:6.37,al2o3:0,simg:5.68,location:"",source:"inventory"},
{id:"DOME 181_1",contractor:"NURIS",stock:0,initialStock:159,ni:1.13,fe:17.44,co:0.03,sio2:35.47,mgo:6.73,al2o3:0,simg:5.27,location:"",source:"inventory"},
{id:"DOME 182_1",contractor:"NURIS",stock:0,initialStock:259,ni:1.26,fe:16.85,co:0.03,sio2:34.99,mgo:8.37,al2o3:0,simg:4.18,location:"",source:"inventory"},
{id:"DOME 183_1",contractor:"NURIS",stock:0,initialStock:219,ni:1.24,fe:22.55,co:0.04,sio2:29.37,mgo:4.61,al2o3:0,simg:6.37,location:"",source:"inventory"},
{id:"DOME 184_1",contractor:"NURIS",stock:0,initialStock:120,ni:1.24,fe:14.73,co:0.03,sio2:26.27,mgo:10.99,al2o3:0,simg:2.39,location:"",source:"inventory"},
{id:"DOME 185_1",contractor:"NURIS",stock:28,ni:1.88,fe:22.38,co:0.08,sio2:18.62,mgo:9.17,al2o3:0,simg:2.03,location:"",source:"inventory"},
{id:"DOME 186_1",contractor:"NURIS",stock:193,ni:0.53,fe:6.86,co:0.01,sio2:31.27,mgo:21.43,al2o3:0,simg:1.46,location:"",source:"inventory"},
{id:"DOME 187_1",contractor:"NURIS",stock:147,ni:0.95,fe:9.02,co:0.02,sio2:29.31,mgo:18.38,al2o3:0,simg:1.59,location:"",source:"inventory"},
{id:"DOME 188_1",contractor:"NURIS",stock:78,ni:0.52,fe:7.99,co:0.01,sio2:39.22,mgo:26.18,al2o3:0,simg:1.5,location:"",source:"inventory"},
{id:"DOME 189_1",contractor:"NURIS",stock:471,ni:0.86,fe:9.32,co:0.02,sio2:30.17,mgo:19.3,al2o3:0,simg:1.56,location:"",source:"inventory"},
{id:"DOME 190_1",contractor:"NURIS",stock:15,ni:0.9,fe:8.45,co:0.01,sio2:28.74,mgo:16.79,al2o3:0,simg:1.71,location:"",source:"inventory"},
{id:"DOME 191_1",contractor:"NURIS",stock:31,ni:0.97,fe:11.44,co:0.02,sio2:18.93,mgo:10.02,al2o3:0,simg:1.89,location:"",source:"inventory"},
{id:"DOME 192_1",contractor:"IGO",stock:363,ni:1.83,fe:28.79,co:0.1,sio2:14.55,mgo:7.7,al2o3:0,simg:1.89,location:"",source:"inventory"},
{id:"DOME 193_1",contractor:"IGO",stock:395,ni:1.79,fe:30.52,co:0.11,sio2:12.96,mgo:6.65,al2o3:0,simg:1.95,location:"",source:"inventory"},
{id:"DOME 194_1",contractor:"IGO",stock:254,ni:1.5,fe:18.02,co:0.05,sio2:18.8,mgo:10.15,al2o3:0,simg:1.85,location:"",source:"inventory"},
{id:"DOME 195_1",contractor:"IGO",stock:120,ni:1.88,fe:27.18,co:0.09,sio2:15.31,mgo:7.97,al2o3:0,simg:1.92,location:"",source:"inventory"},
{id:"DOME 196_1",contractor:"IGO",stock:352,ni:1.81,fe:28.67,co:0.1,sio2:11.2,mgo:5.32,al2o3:0,simg:2.11,location:"",source:"inventory"},
{id:"DOME 197_1",contractor:"IGO",stock:71,ni:1.63,fe:36.67,co:0.13,sio2:7.59,mgo:3.4,al2o3:0,simg:2.23,location:"",source:"inventory"},
{id:"DOME 198_1",contractor:"NURIS",stock:24,ni:0.97,fe:16.01,co:0.04,sio2:12.96,mgo:5.96,al2o3:0,simg:2.17,location:"",source:"inventory"},
{id:"DOME 199_1",contractor:"NURIS",stock:202,ni:0.68,fe:11.15,co:0.02,sio2:32.17,mgo:19.6,al2o3:0,simg:1.64,location:"",source:"inventory"},
{id:"DOME 200_1",contractor:"NURIS",stock:114,ni:0.67,fe:12.12,co:0.02,sio2:37.15,mgo:23.56,al2o3:0,simg:1.58,location:"",source:"inventory"},
{id:"DOME 201",contractor:"RAHMAT",stock:1981,initialStock:3014,ni:0.97,fe:20.89,co:0.05,sio2:13.69,mgo:5.35,al2o3:0,simg:2.56,location:"",source:"inventory"},
{id:"DOME 202",contractor:"RAHMAT",stock:1252,ni:0.87,fe:13.97,co:0.03,sio2:13.11,mgo:4.84,al2o3:0,simg:2.71,location:"",source:"inventory"},
{id:"DOME 204",contractor:"RAHMAT",stock:1274,ni:0.92,fe:31.48,co:0.07,sio2:14.07,mgo:2.85,al2o3:0,simg:4.94,location:"",source:"inventory"},
{id:"DOME 205",contractor:"RAHMAT",stock:637,ni:0.56,fe:14.98,co:0.02,sio2:20.83,mgo:3.14,al2o3:0,simg:6.63,location:"",source:"inventory"},
{id:"DOME 206",contractor:"HAIRIL",stock:0,initialStock:280,ni:1.7,fe:13.48,co:0.02,sio2:43.78,mgo:18.55,al2o3:0,simg:2.36,location:"",source:"inventory"},
{id:"DOME 207",contractor:"HAIRIL",stock:0,initialStock:648,ni:1.74,fe:13.83,co:0.02,sio2:41.99,mgo:15.77,al2o3:0,simg:2.66,location:"",source:"inventory"},
{id:"DOME 208",contractor:"HAIRIL",stock:0,initialStock:385,ni:1.51,fe:11.78,co:0.02,sio2:45.08,mgo:20.78,al2o3:0,simg:2.17,location:"",source:"inventory"},
{id:"DOME 209",contractor:"HAIRIL",stock:0,initialStock:600,ni:0.88,fe:12.81,co:0.02,sio2:46.95,mgo:18.16,al2o3:0,simg:2.59,location:"",source:"inventory"},
{id:"DOME 210",contractor:"HAIRIL",stock:0,initialStock:578,ni:1.69,fe:16.28,co:0.03,sio2:42.03,mgo:12.38,al2o3:0,simg:3.39,location:"",source:"inventory"},
{id:"DOME 212",contractor:"HAIRIL",stock:250,ni:1.55,fe:20.68,co:0.03,sio2:39.28,mgo:8.91,al2o3:0,simg:4.41,location:"",source:"inventory"},
{id:"DOME 214",contractor:"HAIRIL",stock:0,initialStock:649,ni:1.64,fe:11.93,co:0.02,sio2:44.21,mgo:20.33,al2o3:0,simg:2.17,location:"",source:"inventory"},
{id:"DOME 215",contractor:"IMN",stock:38,initialStock:104,ni:0.98,fe:13.8,co:0.02,sio2:49.29,mgo:14.61,al2o3:0,simg:3.37,location:"",source:"inventory"},
{id:"DOME 217",contractor:"IMN",stock:387,ni:0.95,fe:13.86,co:0.03,sio2:29.6,mgo:12.39,al2o3:0,simg:2.39,location:"",source:"inventory"},
{id:"DOME 219",contractor:"HAIRIL",stock:0,initialStock:549,ni:1.73,fe:22.41,co:0.04,sio2:39.68,mgo:7.76,al2o3:0,simg:5.11,location:"",source:"inventory"},
{id:"DOME 220",contractor:"HAIRIL",stock:0,initialStock:280,ni:1.34,fe:18.48,co:0.05,sio2:33.28,mgo:10.9,al2o3:0,simg:3.05,location:"",source:"inventory"},
{id:"DOME 222",contractor:"HAIRIL",stock:0,initialStock:214,ni:1.68,fe:22.61,co:0.04,sio2:38.05,mgo:7.02,al2o3:0,simg:5.42,location:"",source:"inventory"},
{id:"DOME 223",contractor:"HAIRIL",stock:0,initialStock:566,ni:1.71,fe:18.47,co:0.04,sio2:42.34,mgo:11.54,al2o3:0,simg:3.67,location:"",source:"inventory"},
{id:"DOME 224",contractor:"HAIRIL",stock:0,initialStock:490,ni:1.78,fe:19.6,co:0.04,sio2:42.8,mgo:9.98,al2o3:0,simg:4.29,location:"",source:"inventory"},
{id:"DOME 225",contractor:"HAIRIL",stock:237,initialStock:2481,ni:1.07,fe:15.04,co:0.03,sio2:43.24,mgo:15.4,al2o3:0,simg:2.81,location:"",source:"inventory"},
{id:"DOME 229",contractor:"IGO",stock:597,ni:1.46,fe:22.12,co:0.04,sio2:40.3,mgo:7.95,al2o3:0,simg:5.07,location:"",source:"inventory"},
{id:"DOME 231",contractor:"HAIRIL",stock:140,initialStock:323,ni:1.98,fe:17.75,co:0.03,sio2:45.46,mgo:9.9,al2o3:0,simg:4.59,location:"",source:"inventory"},
{id:"DOME 233",contractor:"IGO",stock:1044,ni:1.26,fe:29.58,co:0.05,sio2:28.15,mgo:5.23,al2o3:0,simg:5.38,location:"",source:"inventory"},
{id:"DOME 242",contractor:"HAIRIL",stock:402,ni:1.68,fe:16.54,co:0.03,sio2:42.87,mgo:12.84,al2o3:0,simg:3.34,location:"",source:"inventory"},
{id:"DOME 248",contractor:"AMRUL",stock:44,ni:1.81,fe:17.81,co:0.04,sio2:42.72,mgo:12.75,al2o3:0,simg:3.35,location:"",source:"inventory"},
{id:"D.01/AMR-N1/IMN-ANM",contractor:"ABAD",stock:0,initialStock:299,ni:1.22,fe:12.73,co:0.02,sio2:46.54,mgo:13.69,al2o3:0,simg:3.4,location:"",source:"inventory"},
{id:"D.01/AMR-S2/IMN-ANM",contractor:"ABAD",stock:0,initialStock:544,ni:1.93,fe:18.05,co:0.05,sio2:35.43,mgo:13.3,al2o3:0,simg:2.66,location:"",source:"inventory"},
{id:"D.02/AMR-S2/IMN-ANM",contractor:"ABAD",stock:400,ni:2,fe:17.49,co:0.04,sio2:36.8,mgo:14.51,al2o3:0,simg:2.54,location:"",source:"inventory"},
{id:"D.03/AMR-S2/IMN-ANM",contractor:"ABAD",stock:0,initialStock:350,ni:2,fe:16.03,co:0.04,sio2:38.09,mgo:14.79,al2o3:0,simg:2.58,location:"",source:"inventory"},
{id:"D.05/AMR-S2/IMN-ANM",contractor:"ABAD",stock:167,initialStock:514,ni:1.53,fe:21.93,co:0.06,sio2:33.81,mgo:10.17,al2o3:0,simg:3.32,location:"",source:"inventory"},
{id:"D.11/AMR-S2/IMN-ANM",contractor:"ABAD",stock:76,initialStock:514,ni:1.74,fe:15.93,co:0.04,sio2:39.77,mgo:15.06,al2o3:0,simg:2.64,location:"",source:"inventory"},
{id:"D.12/AMR-S2/IMN-ANM",contractor:"ABAD",stock:175,initialStock:514,ni:1.75,fe:16.06,co:0.04,sio2:39.23,mgo:13.94,al2o3:0,simg:2.81,location:"",source:"inventory"},
{id:"D.13/AMR-S2/IMN-ANM",contractor:"ABAD",stock:271,ni:1.82,fe:14.56,co:0.03,sio2:40.46,mgo:15.66,al2o3:0,simg:2.58,location:"",source:"inventory"},
{id:"D.14/AMR-S2/IMN-ANM",contractor:"ABAD",stock:514,ni:1.72,fe:15.44,co:0.04,sio2:40.84,mgo:15.02,al2o3:0,simg:2.72,location:"",source:"inventory"},
{id:"D.15/AMR-S2/IMN-ANM",contractor:"ABAD",stock:87,initialStock:514,ni:1.67,fe:17.24,co:0.04,sio2:37.7,mgo:14.22,al2o3:0,simg:2.65,location:"",source:"inventory"},
{id:"D.16/AMR-S2/IMN-ANM",contractor:"ABAD",stock:514,ni:1.56,fe:18.09,co:0.05,sio2:39.08,mgo:12.67,al2o3:0,simg:3.09,location:"",source:"inventory"},
{id:"D.17/AMR-S2/IMN-ANM",contractor:"ABAD",stock:126,initialStock:514,ni:1.5,fe:15.37,co:0.04,sio2:37.99,mgo:14.9,al2o3:0,simg:2.55,location:"",source:"inventory"},
{id:"D.18/AMR-S2/IMN-ANM",contractor:"ABAD",stock:137,initialStock:514,ni:1.48,fe:12.11,co:0.03,sio2:40.4,mgo:20,al2o3:0,simg:2.02,location:"",source:"inventory"},
{id:"D.19/AMR-S2/IMN-ANM",contractor:"ABAD",stock:514,ni:1.57,fe:12.55,co:0.03,sio2:40.33,mgo:19.25,al2o3:0,simg:2.1,location:"",source:"inventory"},
{id:"D.20/AMR-S2/IMN-ANM",contractor:"ABAD",stock:149,initialStock:514,ni:1.4,fe:11.97,co:0.02,sio2:40.76,mgo:19.88,al2o3:0,simg:2.05,location:"",source:"inventory"},
{id:"D.21/AMR-S2/IMN-ANM",contractor:"ABAD",stock:177,initialStock:514,ni:1.25,fe:12.52,co:0.02,sio2:40.3,mgo:19.45,al2o3:0,simg:2.07,location:"",source:"inventory"},
{id:"D.22/AMR-S2/IMN-ANM",contractor:"ABAD",stock:83,initialStock:442,ni:1.38,fe:14.05,co:0.03,sio2:38.84,mgo:16.82,al2o3:0,simg:2.31,location:"",source:"inventory"},
{id:"D.23/AMR-S2/IMN-ANM",contractor:"ABAD",stock:357,ni:1.29,fe:13.42,co:0.03,sio2:19.76,mgo:18.17,al2o3:0,simg:1.09,location:"",source:"inventory"},
{id:"D.24/AMR-S2/IMN-ANM",contractor:"ABAD",stock:0,initialStock:397,ni:1.38,fe:15.97,co:0.04,sio2:37.14,mgo:14.91,al2o3:0,simg:2.49,location:"",source:"inventory"},
{id:"D.25/AMR-S2/IMN-ANM",contractor:"ABAD",stock:272,initialStock:1028,ni:1.53,fe:18.92,co:0.04,sio2:36.73,mgo:12.46,al2o3:0,simg:2.95,location:"",source:"inventory"},
{id:"D.26/AMR-S2/IMN-ANM",contractor:"ABAD",stock:112,initialStock:1028,ni:1.48,fe:18.47,co:0.05,sio2:36.81,mgo:12.74,al2o3:0,simg:2.89,location:"",source:"inventory"},
{id:"D.28/AMR-S2/IMN-ANM",contractor:"ABAD",stock:0,initialStock:754,ni:1.35,fe:23.31,co:0.12,sio2:35.64,mgo:8.03,al2o3:0,simg:4.44,location:"",source:"inventory"},
{id:"DM 300 A",contractor:"IMN",stock:100,ni:0.53,fe:11.52,co:0.03,sio2:53.51,mgo:17.51,al2o3:0,simg:3.06,location:"",source:"inventory"},
{id:"DM 302 A",contractor:"IMN",stock:350,ni:0.91,fe:10.42,co:0.03,sio2:35.05,mgo:26.5,al2o3:0,simg:1.32,location:"",source:"inventory"},
{id:"DM 303 A",contractor:"SURIADIN",stock:521,initialStock:800,ni:1.12,fe:9.82,co:0.03,sio2:35.58,mgo:24.51,al2o3:0,simg:1.45,location:"",source:"inventory"},
{id:"DM 304 A",contractor:"IMN",stock:500,ni:0.28,fe:5.04,co:0.01,sio2:16.31,mgo:6.14,al2o3:0,simg:2.66,location:"",source:"inventory"},
{id:"DM 305 A",contractor:"SURIADIN",stock:617,initialStock:1200,ni:1.14,fe:14.39,co:0.04,sio2:30.02,mgo:11.79,al2o3:0,simg:2.55,location:"",source:"inventory"},
{id:"DM 306 A",contractor:"SURIADIN",stock:300,ni:0.62,fe:9.07,co:0.03,sio2:16.23,mgo:7.95,al2o3:0,simg:2.04,location:"",source:"inventory"},
{id:"DM 307 A",contractor:"ANDI YUSUF",stock:618,initialStock:1200,ni:1.25,fe:13.27,co:0.04,sio2:21.38,mgo:9.04,al2o3:0,simg:2.37,location:"",source:"inventory"},
{id:"DM 308 A",contractor:"SURIADIN",stock:100,ni:0.9,fe:9.23,co:0.03,sio2:35.09,mgo:27.48,al2o3:0,simg:1.28,location:"",source:"inventory"},
{id:"DM 310 A",contractor:"AMRUL",stock:1000,ni:0.42,fe:6.21,co:0.02,sio2:27.33,mgo:11.44,al2o3:0,simg:2.39,location:"",source:"inventory"},
{id:"DM 311 A",contractor:"AMRUL",stock:700,ni:0.95,fe:15.86,co:0.04,sio2:15.56,mgo:7.34,al2o3:0,simg:2.12,location:"",source:"inventory"},
{id:"DM 315 A",contractor:"GLOBAL/RAMLAN",stock:250,ni:0.89,fe:27.99,co:0.08,sio2:11.49,mgo:4.77,al2o3:0,simg:2.41,location:"",source:"inventory"},
{id:"DM 316 A",contractor:"AMRUL",stock:3000,ni:0.98,fe:39.7,co:0.09,sio2:14.37,mgo:2.89,al2o3:0,simg:4.97,location:"",source:"inventory"},
{id:"DM 317 A",contractor:"AMRUL",stock:700,ni:1.06,fe:33.31,co:0.08,sio2:22,mgo:6.09,al2o3:0,simg:3.61,location:"",source:"inventory"},
{id:"DM 318 A",contractor:"AMRUL",stock:500,ni:1.32,fe:44.26,co:0.15,sio2:7.94,mgo:1.58,al2o3:0,simg:5.03,location:"",source:"inventory"},
{id:"DM 319 A",contractor:"AMRUL",stock:1093,initialStock:3500,ni:1.05,fe:26.15,co:0.06,sio2:27.77,mgo:6.92,al2o3:0,simg:4.01,location:"",source:"inventory"},
{id:"DM 320 A",contractor:"AMRUL",stock:321,initialStock:700,ni:1.56,fe:18.48,co:0.05,sio2:48.93,mgo:7.08,al2o3:0,simg:6.91,location:"",source:"inventory"},
{id:"DM 321 A",contractor:"AMRUL",stock:100,ni:1.53,fe:15.81,co:0.04,sio2:39.34,mgo:13.22,al2o3:0,simg:2.98,location:"",source:"inventory"},
{id:"DM 322 A",contractor:"TAUFIK ZUL",stock:1000,ni:0.72,fe:5.92,co:0.02,sio2:21.91,mgo:11.53,al2o3:0,simg:1.9,location:"",source:"inventory"},
{id:"DM 323 A",contractor:"TAUFIK ZUL",stock:0,initialStock:1301,ni:0.9,fe:8.85,co:0.03,sio2:21.5,mgo:11.65,al2o3:0,simg:1.85,location:"",source:"inventory"},
{id:"DM 324 A",contractor:"TAUFIK ZUL",stock:0,initialStock:1444,ni:1.11,fe:16.8,co:0.05,sio2:17.5,mgo:8.82,al2o3:0,simg:1.98,location:"",source:"inventory"},
{id:"DM 325 A",contractor:"IMN-SR",stock:250,ni:0.59,fe:11.72,co:0.03,sio2:17.95,mgo:2.45,al2o3:0,simg:7.33,location:"",source:"inventory"},
{id:"DM 326 A",contractor:"ANDI YUSUF",stock:0,initialStock:1282,ni:0.89,fe:10.27,co:0.03,sio2:18.01,mgo:8.16,al2o3:0,simg:2.21,location:"",source:"inventory"},
{id:"DM 327 A",contractor:"ANDI YUSUF",stock:700,ni:0.6,fe:7.89,co:0.02,sio2:18.44,mgo:8.38,al2o3:0,simg:2.2,location:"",source:"inventory"},
{id:"DM 328 A",contractor:"ANDI YUSUF",stock:350,ni:0.83,fe:28.36,co:0.08,sio2:14.58,mgo:4.78,al2o3:0,simg:3.05,location:"",source:"inventory"},
{id:"DM 329 A",contractor:"ANDI YUSUF",stock:500,ni:0.58,fe:10.19,co:0.03,sio2:12.53,mgo:4.65,al2o3:0,simg:2.69,location:"",source:"inventory"},
{id:"DM 330 A",contractor:"AMRUL",stock:21,initialStock:200,ni:1.59,fe:17.34,co:0.04,sio2:38.06,mgo:12.29,al2o3:0,simg:3.1,location:"",source:"inventory"},
{id:"DM 331 A",contractor:"IMN",stock:1057,initialStock:4000,ni:1.23,fe:23.37,co:0.08,sio2:15.04,mgo:6.82,al2o3:0,simg:2.21,location:"",source:"inventory"},
{id:"DM 332 A",contractor:"SURIADIN",stock:5000,ni:0.44,fe:7.46,co:0.02,sio2:16.31,mgo:6.41,al2o3:0,simg:2.54,location:"",source:"inventory"},
{id:"DM 333 A",contractor:"AMRUL",stock:0,initialStock:5556,ni:1.32,fe:19.69,co:0.05,sio2:28.56,mgo:9.71,al2o3:0,simg:2.94,location:"",source:"inventory"},
{id:"DM 334 A",contractor:"CPK",stock:150,ni:1.08,fe:14.98,co:0.04,sio2:47.54,mgo:14.83,al2o3:0,simg:3.21,location:"",source:"inventory"},
{id:"DM 335 A",contractor:"CPK",stock:200,ni:1.28,fe:18.58,co:0.05,sio2:40.42,mgo:15.87,al2o3:0,simg:2.55,location:"",source:"inventory"},
{id:"DM 336 A",contractor:"CPK",stock:100,ni:0.61,fe:22.94,co:0.05,sio2:42.41,mgo:3.18,al2o3:0,simg:13.34,location:"",source:"inventory"},
{id:"DOME 350",contractor:"SAID",stock:0,initialStock:1809,ni:1.19,fe:45.52,co:0.12,sio2:7.46,mgo:2,al2o3:0,simg:3.73,location:"",source:"inventory"},
{id:"DOME 351",contractor:"SAID",stock:0,initialStock:2421,ni:1.21,fe:46.13,co:0.13,sio2:7.02,mgo:1.61,al2o3:0,simg:4.36,location:"",source:"inventory"},
{id:"DOME 352",contractor:"SAID",stock:0,initialStock:1114,ni:1.22,fe:44.37,co:0.13,sio2:9.1,mgo:1.94,al2o3:0,simg:4.69,location:"",source:"inventory"},
{id:"DOME 353",contractor:"SAID",stock:0,initialStock:777,ni:1.1,fe:35.22,co:0.12,sio2:16.02,mgo:2.45,al2o3:0,simg:6.54,location:"",source:"inventory"},
{id:"DOME 354",contractor:"SAID",stock:0,initialStock:3213,ni:1.18,fe:45.98,co:0.13,sio2:7.36,mgo:1.83,al2o3:0,simg:4.02,location:"",source:"inventory"},
{id:"DOME 355",contractor:"SAID",stock:0,initialStock:3173,ni:1.21,fe:42.1,co:0.12,sio2:10.98,mgo:3.76,al2o3:0,simg:2.92,location:"",source:"inventory"},
{id:"DOME 356",contractor:"CPK",stock:0,initialStock:119,ni:1.54,fe:14.11,co:0.04,sio2:35.75,mgo:26.69,al2o3:0,simg:1.34,location:"",source:"inventory"},
{id:"DOME 357",contractor:"CPK",stock:100,ni:1.94,fe:13.5,co:0.03,sio2:36.22,mgo:26.91,al2o3:0,simg:1.35,location:"",source:"inventory"},
{id:"DOME 358",contractor:"CPK",stock:34,initialStock:400,ni:1.31,fe:17.77,co:0.04,sio2:46.15,mgo:11.3,al2o3:0,simg:4.08,location:"",source:"inventory"},
{id:"DOME 359",contractor:"CPK/SR",stock:250,ni:1.07,fe:9.99,co:0.03,sio2:40.15,mgo:29.84,al2o3:0,simg:1.35,location:"",source:"inventory"},
{id:"DOME 360",contractor:"AMRUL",stock:300,ni:0.84,fe:12.62,co:0.03,sio2:48.04,mgo:18.33,al2o3:0,simg:2.62,location:"",source:"inventory"},
{id:"DOME 361",contractor:"AMRUL",stock:250,ni:0.89,fe:12.93,co:0.03,sio2:49.9,mgo:17.29,al2o3:0,simg:2.89,location:"",source:"inventory"},
{id:"DOME 362",contractor:"AMRUL",stock:400,ni:1.13,fe:14.5,co:0.03,sio2:50.05,mgo:12.93,al2o3:0,simg:3.87,location:"",source:"inventory"},
{id:"DOME 363",contractor:"AMRUL",stock:0,initialStock:256,ni:1.22,fe:13.8,co:0.03,sio2:50.55,mgo:14.73,al2o3:0,simg:3.43,location:"",source:"inventory"},
{id:"DOME 364",contractor:"AMRUL",stock:50,initialStock:150,ni:0.96,fe:23.82,co:0.08,sio2:43.37,mgo:5.58,al2o3:0,simg:7.77,location:"",source:"inventory"},
{id:"DOME 365",contractor:"AMRUL",stock:700,ni:1.11,fe:13.87,co:0.03,sio2:49.87,mgo:14.91,al2o3:0,simg:3.34,location:"",source:"inventory"},
{id:"DOME 366",contractor:"AMRUL",stock:0,initialStock:621,ni:1.3,fe:15.07,co:0.04,sio2:49.39,mgo:13.11,al2o3:0,simg:3.77,location:"",source:"inventory"},
{id:"DOME 367",contractor:"AMRUL",stock:0,initialStock:656,ni:1.48,fe:15.09,co:0.04,sio2:50.04,mgo:12.87,al2o3:0,simg:3.89,location:"",source:"inventory"},
{id:"DOME 368",contractor:"AMRUL",stock:350,ni:1.08,fe:13.72,co:0.03,sio2:50.39,mgo:14.58,al2o3:0,simg:3.46,location:"",source:"inventory"},
{id:"DOME 369",contractor:"AMRUL",stock:0,initialStock:218,ni:1.22,fe:15.13,co:0.04,sio2:48.57,mgo:12.5,al2o3:0,simg:3.89,location:"",source:"inventory"},
{id:"DOME 370",contractor:"YUSUF",stock:10,initialStock:150,ni:1.26,fe:15.18,co:0.04,sio2:48.34,mgo:13.28,al2o3:0,simg:3.64,location:"",source:"inventory"},
{id:"DOME 371",contractor:"YUSUF",stock:0,initialStock:159,ni:1.15,fe:14.49,co:0.04,sio2:50.05,mgo:14.35,al2o3:0,simg:3.49,location:"",source:"inventory"},
{id:"DOME 372",contractor:"AMRUL",stock:30,initialStock:150,ni:1.22,fe:14.96,co:0.04,sio2:49.99,mgo:13.51,al2o3:0,simg:3.7,location:"",source:"inventory"},
{id:"DOME 373",contractor:"AMRUL",stock:0,initialStock:279,ni:1.28,fe:14.55,co:0.04,sio2:49.59,mgo:14.41,al2o3:0,simg:3.44,location:"",source:"inventory"},
{id:"DOME 374",contractor:"AMRUL",stock:150,ni:1.27,fe:14.47,co:0.04,sio2:48.79,mgo:14.25,al2o3:0,simg:3.42,location:"",source:"inventory"},
{id:"DOME 375",contractor:"AMRUL",stock:0,initialStock:278,ni:1.3,fe:13.67,co:0.03,sio2:49.87,mgo:14.95,al2o3:0,simg:3.34,location:"",source:"inventory"},
{id:"DOME 376",contractor:"AMRUL",stock:51,initialStock:250,ni:1.39,fe:14.83,co:0.04,sio2:50.55,mgo:13.93,al2o3:0,simg:3.63,location:"",source:"inventory"},
{id:"DOME 377",contractor:"AMRUL",stock:250,ni:1.21,fe:13.53,co:0.03,sio2:48.98,mgo:16.86,al2o3:0,simg:2.91,location:"",source:"inventory"},
{id:"DOME 378",contractor:"AMRUL",stock:200,ni:1.28,fe:13.74,co:0.03,sio2:51.23,mgo:15.04,al2o3:0,simg:3.41,location:"",source:"inventory"},
{id:"DOME 379",contractor:"AMRUL",stock:170,ni:1.12,fe:15.53,co:0.04,sio2:49.9,mgo:12.5,al2o3:0,simg:3.99,location:"",source:"inventory"},
{id:"DOME 380",contractor:"AMRUL",stock:281,initialStock:500,ni:1.02,fe:12.87,co:0.03,sio2:48.72,mgo:17.82,al2o3:0,simg:2.73,location:"",source:"inventory"},
{id:"DOME 381",contractor:"AMRUL",stock:100,ni:1.06,fe:16.75,co:0.05,sio2:44.51,mgo:14.56,al2o3:0,simg:3.06,location:"",source:"inventory"},
{id:"DOME 382",contractor:"CPK",stock:22,initialStock:100,ni:1.63,fe:10.04,co:0.03,sio2:39.96,mgo:26.69,al2o3:0,simg:1.5,location:"",source:"inventory"},
{id:"DOME 383",contractor:"HAIRIL",stock:48,initialStock:100,ni:1.65,fe:17.53,co:0.05,sio2:37.61,mgo:15.6,al2o3:0,simg:2.41,location:"",source:"inventory"},
{id:"DOME 384",contractor:"HAIRIL",stock:0,initialStock:210,ni:1.42,fe:16.36,co:0.04,sio2:45.61,mgo:14.03,al2o3:0,simg:3.25,location:"",source:"inventory"},
{id:"DOME 385",contractor:"HAIRIL",stock:0,initialStock:266,ni:1.71,fe:21.3,co:0.05,sio2:39.91,mgo:9.6,al2o3:0,simg:4.16,location:"",source:"inventory"},
{id:"DOME 386",contractor:"HAIRIL",stock:28,initialStock:150,ni:1.54,fe:16.14,co:0.04,sio2:37.61,mgo:15.99,al2o3:0,simg:2.35,location:"",source:"inventory"},
{id:"DOME 387",contractor:"HAIRIL",stock:13,initialStock:100,ni:1.53,fe:16.26,co:0.04,sio2:39.53,mgo:13.61,al2o3:0,simg:2.9,location:"",source:"inventory"},
{id:"DOME 388",contractor:"IMN",stock:150,ni:0.97,fe:6.5,co:0.02,sio2:33.17,mgo:21.96,al2o3:0,simg:1.51,location:"",source:"inventory"},
{id:"DOME 389",contractor:"IMN",stock:100,ni:0.79,fe:6.23,co:0.02,sio2:36.45,mgo:24.11,al2o3:0,simg:1.51,location:"",source:"inventory"},
{id:"DOME 390",contractor:"IMN",stock:80,ni:0.92,fe:5.74,co:0.02,sio2:31.85,mgo:20.89,al2o3:0,simg:1.52,location:"",source:"inventory"},
{id:"DOME 393 (D.29/AMR-S.2)",contractor:"AMR - NSS",stock:77,initialStock:1050,ni:1.42,fe:21.36,co:0.06,sio2:31.95,mgo:10.6,al2o3:0,simg:3.01,location:"",source:"inventory"},
{id:"DOME 391  (D.30/AMR-S.2)",contractor:"AMR - NSS",stock:0,initialStock:1094,ni:1.53,fe:19.54,co:0.05,sio2:35.09,mgo:11.36,al2o3:0,simg:3.09,location:"",source:"inventory"},
{id:"DOME 392  (D.31/AMR-S.2)",contractor:"AMR - NSS",stock:177,initialStock:1050,ni:1.63,fe:16.17,co:0.04,sio2:37.76,mgo:14.01,al2o3:0,simg:2.7,location:"",source:"inventory"},
{id:"DOME.32/AMR-S2",contractor:"AMR - NSS",stock:1050,ni:1.65,fe:15.52,co:0.04,sio2:38.32,mgo:15.65,al2o3:0,simg:2.45,location:"",source:"inventory"},
{id:"D-400-1MN",contractor:"AMR",stock:150,ni:0.96,fe:12.55,co:0,sio2:43.35,mgo:15.32,al2o3:0,simg:2.83,location:"",source:"inventory"},
{id:"D-401-1MN",contractor:"AMR",stock:150,ni:1.14,fe:12.9,co:0,sio2:41.01,mgo:13.68,al2o3:0,simg:3,location:"",source:"inventory"},
{id:"D-402-1MN",contractor:"CPK",stock:100,ni:1.04,fe:10.9,co:0,sio2:40.12,mgo:19.67,al2o3:0,simg:2.04,location:"",source:"inventory"},
{id:"DOME 403/IMN/CPK",contractor:"CPK",stock:55,initialStock:600,ni:1.39,fe:49.31,co:0,sio2:17.42,mgo:4.6,al2o3:0,simg:3.79,location:"",source:"inventory"},
{id:"D-404-1MN",contractor:"IMN",stock:200,ni:0.5,fe:10.12,co:0,sio2:39.72,mgo:18.53,al2o3:0,simg:2.14,location:"",source:"inventory"},
{id:"D-405-1MN",contractor:"IMN",stock:300,ni:0.36,fe:8.19,co:0,sio2:41.1,mgo:23.05,al2o3:0,simg:1.78,location:"",source:"inventory"},
{id:"D-406-1MN",contractor:"IMN",stock:200,ni:0.38,fe:8.67,co:0,sio2:40.45,mgo:22.14,al2o3:0,simg:1.83,location:"",source:"inventory"},
{id:"D-407-1MN",contractor:"IMN",stock:200,ni:0.37,fe:8.24,co:0,sio2:40.81,mgo:24.19,al2o3:0,simg:1.69,location:"",source:"inventory"},
{id:"D-408-1MN",contractor:"IMN",stock:200,ni:0.43,fe:9.3,co:0,sio2:41.01,mgo:20.89,al2o3:0,simg:1.96,location:"",source:"inventory"},
{id:"DOME 409 IMN AMRUL",contractor:"AMR",stock:0,initialStock:305,ni:1.5,fe:17.16,co:0,sio2:25.75,mgo:11.52,al2o3:0,simg:2.24,location:"",source:"inventory"},
{id:"D-411-1MN",contractor:"RAMLAN",stock:150,ni:1.15,fe:13.01,co:0,sio2:35.9,mgo:14.68,al2o3:0,simg:2.45,location:"",source:"inventory"},
{id:"DOME 412",contractor:"AMR",stock:0,initialStock:3006,ni:1.21,fe:15.73,co:0,sio2:37.14,mgo:13.23,al2o3:0,simg:2.81,location:"",source:"inventory"},
{id:"D.INV L2/BLOK S/IMN-1",contractor:"AMR",stock:1000,ni:1.21,fe:16.86,co:0,sio2:33.87,mgo:11.12,al2o3:0,simg:3.05,location:"",source:"inventory"},
{id:"D.INV L3/BLOK S/IMN-1",contractor:"AMR",stock:1000,ni:1.15,fe:17.26,co:0,sio2:34.33,mgo:12.66,al2o3:0,simg:2.71,location:"",source:"inventory"},
{id:"D-420 AMR-IMN",contractor:"IMN",stock:1000,ni:1.29,fe:20.07,co:0,sio2:29.55,mgo:8.16,al2o3:0,simg:3.62,location:"",source:"inventory"},
{id:"D-421 AMR-IMN",contractor:"IMN",stock:1000,ni:0.79,fe:19.1,co:0,sio2:37.13,mgo:6.26,al2o3:0,simg:5.93,location:"",source:"inventory"},
{id:"D-422 AMR-IMN",contractor:"IMN",stock:1000,ni:0.75,fe:19.02,co:0,sio2:38.16,mgo:7.13,al2o3:0,simg:5.35,location:"",source:"inventory"},
{id:"D-423 AMR-IMN",contractor:"IMN",stock:1000,ni:0.82,fe:19.52,co:0,sio2:36.32,mgo:7.2,al2o3:0,simg:5.04,location:"",source:"inventory"},
{id:"D-424 AMR-IMN",contractor:"IMN",stock:1000,ni:0.74,fe:20.19,co:0,sio2:35.74,mgo:6.61,al2o3:0,simg:5.41,location:"",source:"inventory"},
{id:"D-425 AMR-IMN",contractor:"IMN",stock:250,ni:1.11,fe:20.51,co:0,sio2:29.69,mgo:7.59,al2o3:0,simg:3.91,location:"",source:"inventory"},
{id:"D-426 AMR-IMN",contractor:"IMN",stock:300,ni:1.3,fe:14.4,co:0,sio2:30.44,mgo:16.74,al2o3:0,simg:1.82,location:"",source:"inventory"},
{id:"D-427 AMR-IMN",contractor:"IMN",stock:700,ni:0.84,fe:17.05,co:0,sio2:12.07,mgo:5.49,al2o3:0,simg:2.2,location:"",source:"inventory"},
{id:"D-428 AMR-IMN",contractor:"IMN",stock:150,ni:1.27,fe:40.94,co:0,sio2:19.91,mgo:6.63,al2o3:0,simg:3,location:"",source:"inventory"},
{id:"ID.001/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:332,initialStock:1181,ni:1.47,fe:23.58,co:0.06,sio2:28.65,mgo:8.72,al2o3:0,simg:3.29,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.002/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:679,initialStock:1180,ni:1.1,fe:18.66,co:0.04,sio2:35.2,mgo:11.59,al2o3:0,simg:3.04,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.003/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:0,initialStock:1220,ni:1.83,fe:23.06,co:0.06,sio2:30.49,mgo:9.23,al2o3:0,simg:3.3,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.004/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:245,initialStock:1181,ni:1.42,fe:20.52,co:0.06,sio2:31.58,mgo:9.42,al2o3:0,simg:3.35,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.005/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:309,initialStock:1181,ni:1.25,fe:22.94,co:0.06,sio2:29.37,mgo:7.64,al2o3:0,simg:3.84,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.006/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.47,fe:18.69,co:0.05,sio2:34.09,mgo:12.73,al2o3:0,simg:2.68,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.007/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:528,initialStock:1181,ni:1.32,fe:24.2,co:0.06,sio2:28.9,mgo:8.03,al2o3:0,simg:3.6,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.008/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.21,fe:24.73,co:0.06,sio2:27.76,mgo:7.37,al2o3:0,simg:3.77,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.009/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.18,fe:24.78,co:0.06,sio2:27.41,mgo:7.37,al2o3:0,simg:3.72,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.010/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1180,ni:1.22,fe:27.16,co:0.07,sio2:26.66,mgo:6.74,al2o3:0,simg:3.96,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.011/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:295,initialStock:1232,ni:1.57,fe:23.65,co:0.06,sio2:30.25,mgo:9.02,al2o3:0,simg:3.35,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.012/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1198,ni:1.26,fe:26.7,co:0.07,sio2:26.71,mgo:7.06,al2o3:0,simg:3.78,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.013/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.59,fe:20.88,co:0.05,sio2:31.61,mgo:10.71,al2o3:0,simg:2.95,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.014/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.71,fe:14.36,co:0.03,sio2:37.01,mgo:14.09,al2o3:0,simg:2.63,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.015/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1164,ni:1.6,fe:17.81,co:0.04,sio2:33.56,mgo:12.29,al2o3:0,simg:2.73,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.016/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.49,fe:20.81,co:0.05,sio2:31.81,mgo:11.44,al2o3:0,simg:2.78,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.017/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1180,ni:1.67,fe:22.96,co:0.05,sio2:32.67,mgo:12.12,al2o3:0,simg:2.7,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.018/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.59,fe:20.93,co:0.05,sio2:31.29,mgo:11.23,al2o3:0,simg:2.79,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.019/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.59,fe:20.93,co:0.05,sio2:31.29,mgo:11.23,al2o3:0,simg:2.79,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.020/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.72,fe:24.7,co:0.07,sio2:29.07,mgo:11.83,al2o3:0,simg:2.46,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.021/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1198,ni:1.37,fe:23.7,co:0.06,sio2:29.21,mgo:8.94,al2o3:0,simg:3.27,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.022/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.6375,fe:16.4625,co:0.035,sio2:35.957499999999996,mgo:13.212499999999999,al2o3:3.6325,simg:2.7274999999999996,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.023/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.7025000000000001,fe:16.4675,co:0.037500000000000006,sio2:34.9825,mgo:12.612499999999999,al2o3:4.0675,simg:2.7925,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.024/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.835,fe:16.3125,co:0.04,sio2:36.595,mgo:13.372499999999999,al2o3:3.4325,simg:2.7375000000000003,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.025/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.635,fe:14.2375,co:0.0275,sio2:36.45,mgo:15.16,al2o3:3.1100000000000003,simg:2.41,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.026/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.5325,fe:15.827499999999999,co:0.035,sio2:37.4975,mgo:13.114999999999998,al2o3:3.5475000000000003,simg:2.8925,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.027/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.57,fe:13.55,co:0.022500000000000003,sio2:37.9625,mgo:15.380000000000003,al2o3:3.185,simg:2.4675000000000002,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.028/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:2.0425,fe:13.3825,co:0.0275,sio2:39.552499999999995,mgo:14.002500000000001,al2o3:2.6149999999999998,simg:2.8275,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.029/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.7875,fe:18.4575,co:0.0425,sio2:34.715,mgo:12.32,al2o3:3.9025000000000003,simg:2.82,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.030/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.6475,fe:15.665,co:0.0375,sio2:38.16,mgo:12.600000000000001,al2o3:3.3975,simg:3.05,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.031/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:1.74,fe:22.597500000000004,co:0.060000000000000005,sio2:32.7425,mgo:11.515,al2o3:4.4875,simg:2.935,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.032/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.033/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.034/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.035/BLOK.S/IMN 01/2026",contractor:"IMN-1",stock:1181,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-1-SELATAN",source:"production"},
{id:"ID.001/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:437,ni:0.98,fe:16.6,co:0.04,sio2:38.89,mgo:10.03,al2o3:6.27,simg:3.88,location:"IMN-2-UTARA",source:"production"},
{id:"ID.002/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:109,ni:1.05,fe:17.41,co:0.05,sio2:39.35,mgo:9.88,al2o3:5.57,simg:3.98,location:"IMN-2-UTARA",source:"production"},
{id:"ID.003/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:87,ni:1.06,fe:17.3,co:0.05,sio2:37.46,mgo:11.87,al2o3:5.67,simg:3.16,location:"IMN-2-UTARA",source:"production"},
{id:"ID.004/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:1.02,fe:17.09,co:0.04,sio2:47.32,mgo:9.95,al2o3:6.09,simg:4.76,location:"IMN-2-UTARA",source:"production"},
{id:"ID.005/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:306,ni:1.11,fe:13.45,co:0.04,sio2:53.46,mgo:12.98,al2o3:2.9,simg:4.12,location:"IMN-2-UTARA",source:"production"},
{id:"ID.006/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:1.03,fe:16.12,co:0.04,sio2:37.73,mgo:13,al2o3:4.84,simg:2.9,location:"IMN-2-UTARA",source:"production"},
{id:"ID.007/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:459,ni:1.19,fe:15.95,co:0.04,sio2:41.16,mgo:10.25,al2o3:4.54,simg:4.02,location:"IMN-2-UTARA",source:"production"},
{id:"ID.008/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:240,ni:1.34,fe:16.05,co:0.03,sio2:42.36,mgo:10.69,al2o3:3.44,simg:3.96,location:"IMN-2-UTARA",source:"production"},
{id:"ID.009/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:437,ni:0.89,fe:13.41,co:0.03,sio2:44.37,mgo:12.34,al2o3:3.45,simg:3.6,location:"IMN-2-UTARA",source:"production"},
{id:"ID.010/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:371,ni:0.9,fe:17.48,co:0.04,sio2:38.53,mgo:9.02,al2o3:7.43,simg:4.27,location:"IMN-2-UTARA",source:"production"},
{id:"ID.011/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:0.93,fe:16.42,co:0.04,sio2:39.87,mgo:10.05,al2o3:6.07,simg:3.97,location:"IMN-2-UTARA",source:"production"},
{id:"ID.012/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:240,ni:1.03,fe:14.78,co:0.03,sio2:39.32,mgo:14.77,al2o3:3.51,simg:2.66,location:"IMN-2-UTARA",source:"production"},
{id:"ID.013/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:175,ni:0.91,fe:14.01,co:0.03,sio2:42.32,mgo:13.66,al2o3:3.13,simg:3.1,location:"IMN-2-UTARA",source:"production"},
{id:"ID.014/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:502,ni:0.97,fe:15.27,co:0.03,sio2:41.55,mgo:11.77,al2o3:4.45,simg:3.53,location:"IMN-2-UTARA",source:"production"},
{id:"ID.015/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:393,ni:0.92,fe:11.67,co:0.02,sio2:40.99,mgo:15.13,al2o3:4.14,simg:2.71,location:"IMN-2-UTARA",source:"production"},
{id:"ID.016/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:743,ni:1.19,fe:12.65,co:0.02,sio2:42.45,mgo:14.99,al2o3:3.15,simg:2.83,location:"IMN-2-UTARA",source:"production"},
{id:"ID.017/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:393,ni:1.03,fe:13.38,co:0.02,sio2:43.33,mgo:14.12,al2o3:3.04,simg:3.07,location:"IMN-2-UTARA",source:"production"},
{id:"ID.018/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:590,ni:0.96,fe:13.12,co:0.02,sio2:43.5,mgo:13.07,al2o3:3.98,simg:3.33,location:"IMN-2-UTARA",source:"production"},
{id:"ID.019/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:699,ni:0.93,fe:13.61,co:0.03,sio2:44.23,mgo:14.5,al2o3:3.04,simg:3.05,location:"IMN-2-UTARA",source:"production"},
{id:"ID.020/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:502,ni:0.92,fe:13.49,co:0.03,sio2:40.51,mgo:12.87,al2o3:3.48,simg:3.15,location:"IMN-2-UTARA",source:"production"},
{id:"ID.021/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:349,ni:0.78,fe:12.12,co:0.02,sio2:41.68,mgo:13.83,al2o3:2.6,simg:3.01,location:"IMN-2-UTARA",source:"production"},
{id:"ID.001/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:743,ni:1.05,fe:15.2,co:0.03,sio2:37.21,mgo:12.06,al2o3:4.17,simg:3.08,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.002/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:175,ni:1.09,fe:15.94,co:0.03,sio2:36.29,mgo:11.64,al2o3:3.98,simg:3.12,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.003/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:655,ni:0.99,fe:16.75,co:0.04,sio2:18.79,mgo:8.04,al2o3:4.04,simg:2.34,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.022/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:546,ni:0.85,fe:11.83,co:0.02,sio2:42.24,mgo:16.14,al2o3:2.82,simg:2.62,location:"IMN-2-UTARA",source:"production"},
{id:"ID.004/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:808,ni:1.16,fe:15.84,co:0.03,sio2:28.85,mgo:11.28,al2o3:5.23,simg:2.56,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.005/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:655,ni:1.02,fe:16.67,co:0.03,sio2:27.86,mgo:10.12,al2o3:5.03,simg:2.75,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.006/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:480,ni:1.21,fe:23.91,co:0.06,sio2:23.24,mgo:8.6,al2o3:7.49,simg:2.7,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.023/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:328,ni:0.87,fe:12.74,co:0.02,sio2:41.28,mgo:14.07,al2o3:2.96,simg:2.93,location:"IMN-2-UTARA",source:"production"},
{id:"ID.007/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:612,ni:1.17,fe:18.18,co:0.04,sio2:27.99,mgo:11.9,al2o3:3.84,simg:2.35,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.008/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1114,ni:1.33,fe:25.54,co:0.07,sio2:23.35,mgo:9.49,al2o3:5.22,simg:2.46,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.009/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:808,ni:1.14,fe:21.24,co:0.05,sio2:26.44,mgo:10.18,al2o3:6.14,simg:2.6,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.010/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:743,ni:1.04,fe:24.76,co:0.06,sio2:25.5,mgo:6.56,al2o3:9.45,simg:3.89,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.011/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:612,ni:1.08,fe:24.1,co:0.06,sio2:24.6,mgo:6.97,al2o3:8.68,simg:3.53,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.012/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1179,ni:1.02,fe:25.59,co:0.06,sio2:27.02,mgo:7.04,al2o3:9.05,simg:3.84,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.013/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1551,ni:1,fe:27.21,co:0.07,sio2:25.67,mgo:5.91,al2o3:10.7,simg:4.34,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.014/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:524,ni:0.98,fe:28.84,co:0.08,sio2:25.97,mgo:5.23,al2o3:12.3,simg:4.97,location:"IMN-2-SELATAN",source:"production"},
{id:"SOUTH-CPK 01",contractor:"IMN-2",stock:677,ni:1.01,fe:14.39,co:0.03,sio2:39.38,mgo:14.1,al2o3:3.5,simg:2.79,location:"IMN-2-UTARA",source:"production"},
{id:"SOUTH-CPK 02",contractor:"IMN-2",stock:284,ni:0.98,fe:20.49,co:0.05,sio2:36.32,mgo:8.63,al2o3:6.15,simg:4.21,location:"IMN-2-UTARA",source:"production"},
{id:"SOUTH-CPK 03",contractor:"IMN-2",stock:328,ni:0.98,fe:20.49,co:0.05,sio2:36.32,mgo:8.63,al2o3:6.15,simg:4.21,location:"IMN-2-UTARA",source:"production"},
{id:"ID.024/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:131,ni:0.82,fe:18.9,co:0.04,sio2:25.22,mgo:9.57,al2o3:6.58,simg:2.64,location:"IMN-2-UTARA",source:"production"},
{id:"ID.025/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:0.89,fe:18.77,co:0.04,sio2:27.15,mgo:8.31,al2o3:7.48,simg:3.27,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.026/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:502,ni:0.81,fe:16.66,co:0.03,sio2:26.87,mgo:9.72,al2o3:5.93,simg:2.77,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.027/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:284,ni:0.82,fe:16.52,co:0.03,sio2:25.35,mgo:9.82,al2o3:5.51,simg:2.58,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.028/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:3058,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.015/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:743,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.016/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1070,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.029/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:524,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.030/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:633,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.017/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1070,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.018/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1572,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.019/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1572,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.031/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:306,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.032/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:349,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.033/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:743,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.034/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:240,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-SELATAN",source:"production"},
{id:"ID.035/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:109,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.036/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.037/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-2-UTARA",source:"production"},
{id:"ID.038/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:437,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.039/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:917,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.040/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:218,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.020/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:677,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.021/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:284,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.022/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:328,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.041/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:284,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.023/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:328,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.024/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:371,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.025/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1158,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.026/BLOK.S/IMN 02/2026",contractor:"IMN-2",stock:1332,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.042/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:1310,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.043/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:983,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.044/BLOK.U/IMN 02/2026",contractor:"IMN-2",stock:765,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"",source:"production"},
{id:"ID.001/BLOK.U/IMN 03/2026",contractor:"IMN-3",stock:1082,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-3-UTARA",source:"production"},
{id:"ID.002/BLOK.U/IMN 03/2026",contractor:"IMN-3",stock:541,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-3-UTARA",source:"production"},
{id:"ID.003/BLOK.U/IMN 03/2026",contractor:"IMN-3",stock:659,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-3-UTARA",source:"production"},
{id:"ID.004/BLOK.U/IMN 03/2026",contractor:"IMN-3",stock:1082,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-3-UTARA",source:"production"},
{id:"ID.005/BLOK.U/IMN 03/2026",contractor:"IMN-3",stock:676,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-3-UTARA",source:"production"},
{id:"ID.006/BLOK.U/IMN 03/2026",contractor:"IMN-3",stock:710,ni:0,fe:0,co:0,sio2:0,mgo:0,al2o3:0,simg:0,location:"IMN-3-UTARA",source:"production"}
];

/* ----------------------------- actual barge data (finalized) -----------------------------
 * Real barges shipped, imported from BARGE_PLAN_01..08.xlsx (Date / Barge Name / Tugboat
 * Name / Dome ID / WMT). Sorted by date, numbered 1-8. Marked finalized so their tonnage is
 * treated as already accounted for. Ni grades pulled from DEFAULT_DOMES at import time.
 * "DOME 391" mapped to "DOME 391  (D.30/AMR-S.2)" (only close match in dome database).
 */
const DEFAULT_BARGES = [
{no:1,shipDate:"2026-06-20",bargeName:"BG GOLDEN WAY 3319",tugboatName:"TB BUANA EXPRESS 19",sources:[{id:"DOME 127",amt:315.0,grade:1.59}, {id:"DOME 130",amt:525.0,grade:1.67}, {id:"DOME 133",amt:298.0,grade:1.74}, {id:"DOME 134",amt:122.0,grade:1.78}, {id:"DOME 137",amt:175.0,grade:1.02}, {id:"DOME 138",amt:192.0,grade:1.13}, {id:"DOME 139",amt:122.0,grade:1.33}, {id:"DOME 140",amt:315.0,grade:1.5}, {id:"DOME 201",amt:1033.0,grade:0.97}, {id:"DOME 206",amt:280.0,grade:1.7}, {id:"DOME 207",amt:648.0,grade:1.74}, {id:"DOME 208",amt:385.0,grade:1.51}, {id:"DOME 210",amt:578.0,grade:1.69}, {id:"DOME 220",amt:280.0,grade:1.34}, {id:"DOME 224",amt:490.0,grade:1.78}, {id:"DOME 225",amt:2244.0,grade:1.07}, {id:"DOME 354",amt:2016.0,grade:1.18}, {id:"DOME 383",amt:52.0,grade:1.65}, {id:"DOME 384",amt:210.0,grade:1.42}, {id:"DOME 386",amt:122.0,grade:1.54}, {id:"DOME 387",amt:87.0,grade:1.53}],totalWMT:10489.0,grade:1.3385,status:"exact",finalized:true},
{no:2,shipDate:"2026-06-22",bargeName:"BG BUANA EXPRESS 15",tugboatName:"TB GOLDEN WAY 3315",sources:[{id:"DOME 01_1",amt:232.0,grade:1.33}, {id:"DOME 96",amt:616.0,grade:1.01}, {id:"DOME 101",amt:233.0,grade:1.97}, {id:"DOME 102",amt:66.0,grade:1.21}, {id:"DOME 103",amt:166.0,grade:0.98}, {id:"DOME 109",amt:433.0,grade:1.07}, {id:"DOME 110",amt:83.0,grade:1.19}, {id:"DOME 128",amt:749.0,grade:1.88}, {id:"DOME 134",amt:83.0,grade:1.78}, {id:"DOME 136",amt:483.0,grade:1.71}, {id:"DOME 165_1",amt:350.0,grade:1.01}, {id:"DOME 209",amt:600.0,grade:0.88}, {id:"DOME 214",amt:649.0,grade:1.64}, {id:"DOME 215",amt:66.0,grade:0.98}, {id:"DOME 219",amt:549.0,grade:1.73}, {id:"DOME 222",amt:50.0,grade:1.68}, {id:"DOME 223",amt:566.0,grade:1.71}, {id:"DM 305 A",amt:583.0,grade:1.14}, {id:"DM 307 A",amt:582.0,grade:1.25}, {id:"DOME 355",amt:2716.0,grade:1.21}, {id:"DOME 358",amt:366.0,grade:1.31}, {id:"DOME 385",amt:266.0,grade:1.71}],totalWMT:10487.0,grade:1.3542,status:"exact",finalized:true},
{no:3,shipDate:"2026-07-01",bargeName:"BG GOLDEN WAY 3306",tugboatName:"TB BUANA EXPRESS 6",sources:[{id:"DOME 04_1",amt:530.0,grade:1.29}, {id:"DOME 08_1",amt:548.0,grade:1.05}, {id:"DOME 19_1",amt:292.0,grade:0.95}, {id:"DOME 23_1",amt:311.0,grade:1.0}, {id:"DOME 30_1",amt:530.0,grade:1.09}, {id:"DOME 100",amt:201.0,grade:1.96}, {id:"DOME 106",amt:110.0,grade:1.1}, {id:"DOME 128",amt:110.0,grade:1.88}, {id:"DOME 132",amt:238.0,grade:1.08}, {id:"DOME 135",amt:566.0,grade:2.0}, {id:"DOME 222",amt:164.0,grade:1.68}, {id:"DOME 231",amt:183.0,grade:1.98}, {id:"D.05/AMR-S2/IMN-ANM",amt:347.0,grade:1.53}, {id:"D.11/AMR-S2/IMN-ANM",amt:438.0,grade:1.74}, {id:"D.20/AMR-S2/IMN-ANM",amt:365.0,grade:1.4}, {id:"D.25/AMR-S2/IMN-ANM",amt:37.0,grade:1.53}, {id:"DM 324 A",amt:1444.0,grade:1.11}, {id:"DOME 350",amt:1809.0,grade:1.19}, {id:"DOME 351",amt:219.0,grade:1.21}, {id:"DOME 354",amt:420.0,grade:1.18}, {id:"DOME 355",amt:457.0,grade:1.21}, {id:"DOME 363",amt:256.0,grade:1.22}, {id:"DOME 366",amt:621.0,grade:1.3}, {id:"DOME 380",amt:219.0,grade:1.02}],totalWMT:10415.0,grade:1.2911,status:"deficit",finalized:true},
{no:4,shipDate:"2026-07-03",bargeName:"BG GOLDEN WAY 3317",tugboatName:"TB BUANA EXPRESS 17",sources:[{id:"DOME 71",amt:486.0,grade:1.36}, {id:"DOME 129",amt:253.0,grade:1.39}, {id:"D.01/AMR-S2/IMN-ANM",amt:544.0,grade:1.93}, {id:"D.03/AMR-S2/IMN-ANM",amt:350.0,grade:2.0}, {id:"D.15/AMR-S2/IMN-ANM",amt:427.0,grade:1.67}, {id:"D.17/AMR-S2/IMN-ANM",amt:388.0,grade:1.5}, {id:"D.25/AMR-S2/IMN-ANM",amt:719.0,grade:1.53}, {id:"D.26/AMR-S2/IMN-ANM",amt:757.0,grade:1.48}, {id:"DM 323 A",amt:1301.0,grade:0.9}, {id:"DM 326 A",amt:1282.0,grade:0.89}, {id:"DM 333 A",amt:2369.0,grade:1.32}, {id:"DOME 353",amt:777.0,grade:1.1}, {id:"DOME 354",amt:777.0,grade:1.18}, {id:"DOME 382",amt:78.0,grade:1.63}],totalWMT:10508.0,grade:1.2957,status:"deficit",finalized:true},
{no:5,shipDate:"2026-07-09",bargeName:"BG BMP 3308",tugboatName:"TB JEFFTAR 28",sources:[{id:"DOME 66",amt:259.0,grade:1.55}, {id:"DOME 72",amt:259.0,grade:1.89}, {id:"DOME 98",amt:617.0,grade:0.97}, {id:"DOME 129",amt:875.0,grade:1.39}, {id:"DM 319 A",amt:1114.0,grade:1.05}, {id:"DM 330 A",amt:179.0,grade:1.59}, {id:"DM 331 A",amt:2943.0,grade:1.23}, {id:"DM 333 A",amt:1611.0,grade:1.32}, {id:"DOME 352",amt:1114.0,grade:1.22}, {id:"DOME 356",amt:119.0,grade:1.54}, {id:"DOME 367",amt:438.0,grade:1.48}, {id:"DOME 391  (D.30/AMR-S.2)",amt:1094.0,grade:1.53}],totalWMT:10622.0,grade:1.2964,status:"deficit",finalized:true},
{no:6,shipDate:"2026-07-09",bargeName:"BG LINTAS TAMA 3303",tugboatName:"TB LEO POWER 2201",sources:[{id:"DOME 116",amt:496.0,grade:1.65}, {id:"DOME 119",amt:932.0,grade:1.2}, {id:"DOME 129",amt:476.0,grade:1.39}, {id:"DOME 153_1",amt:238.0,grade:1.75}, {id:"DOME 155_1",amt:298.0,grade:1.87}, {id:"DOME 168_1",amt:1369.0,grade:1.3}, {id:"D.18/AMR-S2/IMN-ANM",amt:377.0,grade:1.48}, {id:"D.21/AMR-S2/IMN-ANM",amt:337.0,grade:1.25}, {id:"D.24/AMR-S2/IMN-ANM",amt:397.0,grade:1.38}, {id:"D.26/AMR-S2/IMN-ANM",amt:159.0,grade:1.48}, {id:"D.28/AMR-S2/IMN-ANM",amt:575.0,grade:1.35}, {id:"DM 319 A",amt:595.0,grade:1.05}, {id:"DOME 351",amt:2202.0,grade:1.21}, {id:"DOME 367",amt:218.0,grade:1.48}, {id:"DOME 375",amt:278.0,grade:1.3}, {id:"DOME 393 (D.29/AMR-S.2)",amt:694.0,grade:1.42}, {id:"DOME 392  (D.31/AMR-S.2)",amt:873.0,grade:1.63}],totalWMT:10514.0,grade:1.3574,status:"exact",finalized:true},
{no:7,shipDate:"2026-07-19",bargeName:"BG MILKY WAY 124",tugboatName:"TB JELAJAH 124",sources:[{id:"DOME 123",amt:1316.0,grade:1.22}, {id:"DOME 145",amt:937.0,grade:1.74}, {id:"DOME 150_1",amt:140.0,grade:2.27}, {id:"DOME 154_1",amt:1355.0,grade:1.71}, {id:"DOME 163_1",amt:598.0,grade:1.28}, {id:"DOME 168_1",amt:120.0,grade:1.3}, {id:"DOME 178_1",amt:239.0,grade:1.33}, {id:"DOME 180_1",amt:159.0,grade:1.23}, {id:"DOME 181_1",amt:159.0,grade:1.13}, {id:"DOME 182_1",amt:259.0,grade:1.26}, {id:"DOME 183_1",amt:219.0,grade:1.24}, {id:"DOME 184_1",amt:120.0,grade:1.24}, {id:"D.01/AMR-N1/IMN-ANM",amt:299.0,grade:1.22}, {id:"D.12/AMR-S2/IMN-ANM",amt:339.0,grade:1.75}, {id:"D.22/AMR-S2/IMN-ANM",amt:359.0,grade:1.38}, {id:"D.28/AMR-S2/IMN-ANM",amt:179.0,grade:1.35}, {id:"DM 303 A",amt:279.0,grade:1.12}, {id:"DM 319 A",amt:219.0,grade:1.05}, {id:"DM 320 A",amt:379.0,grade:1.56}, {id:"DM 333 A",amt:1576.0,grade:1.32}, {id:"DOME 364",amt:100.0,grade:0.96}, {id:"DOME 370",amt:140.0,grade:1.26}, {id:"DOME 371",amt:159.0,grade:1.15}, {id:"DOME 372",amt:120.0,grade:1.22}, {id:"DOME 373",amt:279.0,grade:1.28}, {id:"DOME 376",amt:199.0,grade:1.39}, {id:"DOME 393 (D.29/AMR-S.2)",amt:279.0,grade:1.42}],totalWMT:10526.0,grade:1.4034,status:"excess",finalized:true},
{no:8,shipDate:"2026-07-21",bargeName:"BG EDWARD 330 5",tugboatName:"TB EDWARD 2000 2",sources:[{id:"DM 319 A",amt:479.0,grade:1.05}, {id:"DOME 369",amt:218.0,grade:1.22}, {id:"DOME 403/IMN/CPK",amt:545.0,grade:1.39}, {id:"DOME 409 IMN AMRUL",amt:305.0,grade:1.5}, {id:"DOME 412",amt:3006.0,grade:1.21}, {id:"ID.003/BLOK.S/IMN 01/2026",amt:1220.0,grade:1.83}, {id:"ID.001/BLOK.S/IMN 01/2026",amt:849.0,grade:1.47}, {id:"ID.004/BLOK.S/IMN 01/2026",amt:936.0,grade:1.42}, {id:"ID.011/BLOK.S/IMN 01/2026",amt:937.0,grade:1.57}, {id:"ID.005/BLOK.S/IMN 01/2026",amt:872.0,grade:1.25}, {id:"ID.007/BLOK.S/IMN 01/2026",amt:653.0,grade:1.32}, {id:"ID.002/BLOK.S/IMN 01/2026",amt:501.0,grade:1.1}],totalWMT:10521.0,grade:1.3692,status:"exact",finalized:true},
];

/* ----------------------------- helpers ----------------------------- */

const fmt = (n, d = 0) => Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });

// Ensures every dome carries an immutable initialStock baseline (set once, never decremented).
// Current .stock decreases as barges finalize; initialStock - stock = stock already barged out.
function withInitialStock(list) {
  return list.map((d) => (d.initialStock !== undefined ? d : { ...d, initialStock: d.stock }));
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
function fillOneBarge(pool, bargeSize, target, allowedContractors) {
  const EPS = 1e-6;
  const CEILING_PAD = 0.2; // max allowed excess over target, in Ni percentage points
  let filled = 0, ni = 0; // ni = running Ni mass, i.e. sum(amt * grade/100)
  const sources = [];
  let guard = 0;
  const t = target / 100;
  const maxT = (target + CEILING_PAD) / 100;

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
    const highs = cands.filter((p) => p.grade > target).sort((a, b) => a.grade - b.grade); // closest-to-target first
    const lows = cands.filter((p) => p.grade < target).sort((a, b) => a.grade - b.grade); // lowest first

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
function generateFromPool(pool, count, bargeSize, target, tolerance) {
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
    const b = fillOneBarge(pool, bargeSize, target, allowed);
    b.status = statusFor(b, bargeSize, target, tolerance);
    b.pair = allowedList;
    out.push(b);
  }
  return out;
}

function poolFromDomes(domes, subtractBarges) {
  const pool = domes.map((d) => ({ id: d.id, contractor: d.contractor, grade: d.ni, remaining: d.stock, source: d.source || "inventory" }));
  subtractBarges.forEach((b) => b.sources.forEach((s) => {
    const p = pool.find((x) => x.id === s.id);
    if (p) p.remaining = Math.max(0, p.remaining - s.amt);
  }));
  return pool;
}

/* ----------------------------- import parsing ----------------------------- */

function parseDomeCSV(csv, source, forcedContractor) {
  const rows = Papa.parse(csv, { skipEmptyLines: true }).data;
  if (!rows.length) return [];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const [domeId, cid, stock, ni, fe, co, sio2, mgo, al2o3, simg, loc] = rows[i];
    if (!domeId) continue;
    out.push({
      id: domeId, contractor: forcedContractor || cid || "UNKNOWN",
      stock: parseFloat(stock) || 0, ni: parseFloat(ni) || 0, fe: parseFloat(fe) || 0,
      co: parseFloat(co) || 0, sio2: parseFloat(sio2) || 0, mgo: parseFloat(mgo) || 0,
      al2o3: parseFloat(al2o3) || 0, simg: parseFloat(simg) || 0, location: loc || "",
      source,
    });
  }
  return out;
}

// Reads a simple 2-column barge composition file: Dome ID, WMT (xlsx or csv)
function parseBargeComposition(rows) {
  // rows: array of arrays (first row = header)
  if (!rows.length) return [];
  const header = rows[0].map((h) => String(h || "").toLowerCase().trim());
  const domeIdx = header.findIndex((h) => h.includes("dome"));
  const wmtIdx = header.findIndex((h) => h.includes("wmt") || h.includes("stock") || h.includes("weight"));
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[domeIdx]) continue;
    const amt = parseFloat(row[wmtIdx]) || 0;
    if (amt <= 0) continue;
    out.push({ domeId: String(row[domeIdx]).trim(), amt });
  }
  return out;
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

function BargeRow({ barge, domesById, pool, onUpdate, onFinalize, onImport, onOpenInvoice, onExportBarge }) {
  const [open, setOpen] = useState(false);
  const [addDome, setAddDome] = useState("");
  const [addAmt, setAddAmt] = useState("");

  const availableDomes = pool.filter((p) => p.remaining > 1).sort((a, b) => b.remaining - a.remaining);

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
        const rows = Papa.parse(evt.target.result, { skipEmptyLines: true }).data;
        const comp = parseBargeComposition(rows);
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
    const validSources = [];
    const unknownIds = [];
    comp.forEach((c) => {
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
    if (validSources.length) onImport(barge.no, validSources);
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
              <label className="import-barge-label">
                <FileUp size={13} /> Import final plan for this barge (.xlsx / .csv)
                <input type="file" accept=".xlsx,.csv" onChange={handleFileImport} />
              </label>
            </>
          )}

          <div className="barge-row-actions">
            <button className={`btn-toggle ${barge.finalized ? "btn-toggle-on" : ""}`} onClick={() => onFinalize(barge.no)}>
              {barge.finalized ? <><Unlock size={13} /> Reopen</> : <><Lock size={13} /> Finalize</>}
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

function TimelineTab({ barges, settings }) {
  const monthCounts = useMemo(() => {
    const arr = MONTHS.map(() => ({ final: 0, draft: 0 }));
    barges.forEach((b) => {
      const m = new Date(b.shipDate).getMonth();
      if (m < 0 || m > 11) return;
      if (b.finalized) arr[m].final += 1; else arr[m].draft += 1;
    });
    return arr;
  }, [barges]);

  const maxCount = Math.max(...monthCounts.map((m) => m.final + m.draft), 1);
  const finalizedTotal = barges.filter((b) => b.finalized).length;
  const currentMonthIdx = TODAY.getMonth();
  const monthsLeft = Math.max(1, 12 - currentMonthIdx);
  const remaining = settings.planTarget - finalizedTotal;

  return (
    <div className="stack">
      <section className="glass summary-strip">
        <Kpi label="Finalized" value={`${finalizedTotal}`} unit={`/ ${settings.planTarget}`} accent="good" />
        <Kpi label="Created" value={`${barges.length}`} unit={`/ ${settings.planTarget}`} />
        <Kpi label="Months left in 2026" value={`${monthsLeft}`} unit="mo" />
        <Kpi label="Required pace" value={fmt(remaining / monthsLeft, 1)} unit="barges/mo" />
      </section>

      <section className="glass panel">
        <div className="panel-head"><TrendingUp size={16} /><span>Shipping schedule by month</span></div>
        <div className="month-chart">
          {monthCounts.map((m, i) => (
            <div className="month-col" key={i}>
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
              {b.finalized && <Lock size={12} className="lock-icon" />}
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

/* ----------------------------- main app ----------------------------- */


/* ============================================================
 * Invoice generator — data URIs for letterhead / signature / stamp
 * ============================================================ */
const INVOICE_LOGO_HEADER = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBAUEBAYFBQUGBgYHCQ4JCQgICRINDQoOFRIWFhUSFBQXGiEcFxgfGRQUHScdHyIjJSUlFhwpLCgkKyEkJST/2wBDAQYGBgkICREJCREkGBQYJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCT/wAARCAF/AfQDASIAAhEBAxEB/8QAHQABAAMAAwEBAQAAAAAAAAAAAAYHCAEEBQMCCf/EAFwQAAEDAwEDBgcIDAkMAQQDAAEAAgMEBREGBxIhCBMxQVGTFRdUVWGR0RQWIjZxdIGyMjVCUnOUsbPBwtLhNDdEZHJ1g5KhIyQlJiczRVNigqLDZRhDY7Sk8PH/xAAbAQEAAgMBAQAAAAAAAAAAAAAABQYBBAcCA//EADYRAQABAQMHCwMEAwEAAAAAAAABAgMEEQUWNFFSkdEGEhMUMlNxgaGxwRUh4TEzQWEkgvDx/9oADAMBAAIRAxEAPwDVCIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiFAREQEREBEQICIiAiIgIiICIiAiIgIiICIiAiIgIidKAiIgIiICIiAiIgIiICIiCP631LNpSyG4w07Kh3Osj3HuLRh2eOR8igB233DpFnpO+d7FJtsXxPPzqL9Ko0lVPLWULxYXjmWVWEYRq/tcchZMu14u3PtaMZxnX/Sy/HhcfM9J3zvYnjwuHmek753sVZ9C5UT9YvneekcEz9DuPd+s8Vl+PC4eZ6TvnexPHjcPM9J3zvYqzRPrF82/SOB9DuPd+s8VmePC4eZ6TvnexPHhcPM9J3zvYqzXKfWb53npHA+h3Hu/WeK2bHtfrrteaGgfaqaNtTM2IvbK4loJxnGFaQWbdHH/AFss3zyL6y0krNkK9Wt4s6qrWrGYlVOUFzsbta0U2NOETHyIiKcQAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiIINti+J5+dRfpVHcFpm+2Ch1HQ+4bjG58G+1+GvLTkdHEKPeKPSg/klR+Mv9qreVckW16t+ks5jDDD7/8Ai05Hy1YXO79FaROOMz9vL+1DYTCvnxSaU8kqPxl/tTxSaU8kqPxl/tUbm7etcb54JXOe6aqt0cVCkIAr58UelPJKj8Zf7U8UelPJKj8Zf7UzdvWuN88DOe6aqt0cVD4RXydkelPI6j8Zf7U8UelPJKj8Zf7VjN29a43zwM57pqq3RxU/o4gass3zyL6y0kopQ7MNNW6tgraelnbNA8SMJneQHDo4ZUrHQrBke4Wl0s6qbTD7z/Ct5byjZX20prsonCIw+4iIpdCiIiB1IiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIUBERAREQEREBERAREQEREBERARRTaTe6/T+nPdtumEM/Pxs3i0O4HOeBVWeNTVvnNn4uz2KKvmV7G62nR2kTj+v2/9S9xyLb3yz6WzmMMcPvjwX8ioDxqau85s/F4/YnjU1d5zZ+Ls9i1c47tqndHFuZsXvap3zwX+ioDxqat85s/F2exDtU1aP8AibPxeP2JnHdtU7o4mbF62qd88F/os/8AjV1b5zZ+Ls9i5G1TVvnNv4uz2JnHdtVW6OJmxetqnfPBf6Kk9N7RtTXHUFtpKm4NfBPUsje3mWDLSePEBXYpK43+zvdM1WcT9taLv+T7S5VRRazH3+/2ERFutAREQEREBERAREQAmURAREQEREBERAREQEREBERAREQEREBERAREQEROlAREQEREBERAREQEREBERAREQEREBERAREQQbbF8Tz86i/KVRpV5bYSPeeeI/hUX6VRxx2j1qj8odL8o+V/5M6H5z8PyuQuOHaPWnDtHrUGsJ1rlBjtHrXPDtHrQfnCBcnB6x60GO0etYHs6P+Nlm+eRfWC0ks26PI99lm4j+GRdf/UtJK48mv2a/H4UjlV+9R4fIiIrIqwiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAidKICIiAiIgIiICJ8iICIiAiIgJxREBERA6EREBERAREQEREBERB8aqipq6LmqunhqI8h25KwOGeo4K6fvbsh6bPbvxZnsXpIvFVnRVOMw902ldMYUzMPO97dk8z278WZ7Fx727J5nt34sz2L0kWOhs9mNz101ptTved727J5nt34sz2J727If+D278WZ7F6KJ0NnsxuY6a02p3vN97dk8z278WZ7E97dk8z278WZ7F6SJ0NnsxuOmtNqd7oRaftEEjJYrVQRyMO817YGAtPa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uOOpSbafsst+0q2U8M9S+jrqQudT1TGb27vY3mubwy04HWCCEHV2E3Cmq9llkbDMx7qaJ0EwB4xva92Qezhg/IVH+UdqF3i4bHbaqOanq69lJUvgkDhgNc4sJHa5oBH0KK23ks3VlQW1uq4YqR32YpYH77x9LgPXlWlR7HNKUWiajR8dLKaCpdzssjpMyul4YlB6A4YGMDHDGEFaclSG0OZfZS2Lwu18bWl2OcEG79z143s5x2DPUu9yp71bve/a7OKuJ1w92ic07XZe2MRuG8R1DLgOPSvJk5LN3p68voNWQMgB+BIad7Jmj/tdjPyEKWN5NthGk6u2Prp57xUubIbrMzec1zTkANzwYeIIzk56eAQdPkqjd0bd88f9In80xdvlREeLqmP/AMnD9SRdHQvJ5u2lrmyrn1rVwwslZK6mtzXRCYtOQHkuII6sY6F6ev8AYhete3aoqKnXNWy3vm52C3vpy+Kn4YwBvgZ6eOOtB4fJPI8Bag+eRfmlfHUqR0/yer/pWUvsu0SroGve10rIKUtbLjtHOYPDI4jrV3Y4Y7UGK9a3Jl52tXGS/vcacXfmJ94/7unZLu7voAYPyrYpq7VZ7Q2pE9JS2yCMFsgc1kLIwOGD0Yx0YVXbTeT1Q60u018tNwFsuFRg1DJIy+KZwGN/gQWuwOOMg9mV4umOTE6CeH3zahdXUMLt4UNK17GP9BLjwH9EA+kIKctdXBctrFJWU0gdBPf2yxvxjea6pyD9IIW3VRGsOTM27akmudhvFNa6Wd4k9zupyeYdgZ3N0jhwyBwwrR0FpCp0ZZ30NXf7jfJpJeddPWOyWnAG63JJDeHQSUEmREQEREBERARE6EBERAREQEREBERAREQEREBERAREQEREBERAREQEREBOtEQECIgIiICIiAidKICIiAiIgIiIGUREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBETCAiIgIiICIiAiIgIiICIiAiIgIiICIiAEREBERAREQEREBERAREQEREBERAREQEREBE606kBERAROnqRARE4ICImQgIiICIiAmETKAmERAREQEREBERAROhEBERAREygIiICJlEBETKAiIgIiICJlMoCInQgIiICIiAiIgIiICImUBEymUBETKAiIgIiICIiAiJlAREQETKIP/9k=";

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

  const handlePrint = () => window.print();

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

/* ============================================================
 * BargeExportModal — print-ready export for a SINGLE barge: its
 * header info (no, date, barge/tugboat name, qty, avg Ni, Draft/
 * Final status) plus the full list of domes that make it up.
 * "Print / Save as PDF" uses the browser's print dialog against
 * the shared .print-area mechanism.
 * ============================================================ */
function BargeExportModal({ barge, domesById, onClose }) {
  const todayISO = new Date().toISOString().slice(0, 10);
  const sources = [...(barge.sources || [])].sort((a, b) => b.amt - a.amt);

  const handlePrint = () => window.print();

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
  const [validationAlert, setValidationAlert] = useState(null); // { bargeNo, violations: [{domeId, current, requested}] }
  const [invoiceBarge, setInvoiceBarge] = useState(null); // barge object currently being invoiced, or null
  const [exportBarge, setExportBarge] = useState(null); // barge object currently being exported, or null
  const domesByIdTop = useMemo(() => { const m = {}; domes.forEach((d) => (m[d.id] = d)); return m; }, [domes]);
  const [showImport, setShowImport] = useState(false);
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

  const resetDefaults = () => { setDomes(withInitialStock(DEFAULT_DOMES)); setSettings(DEFAULT_SETTINGS); setBarges(DEFAULT_BARGES); };

  // finalize / reopen actually mutates the master stock here, since it needs setDomes
  const toggleFinalize = (no) => {
    const barge = barges.find((b) => b.no === no);
    if (!barge) return;
    const sign = barge.finalized ? 1 : -1; // reopening adds back; finalizing subtracts
    
    // VALIDATION: Before finalizing, check if any dome would go negative
    if (sign === -1) { // only validate on finalize, not reopen
      const violations = [];
      barge.sources.forEach((s) => {
        const dome = domes.find((d) => d.id === s.id);
        if (dome && (dome.stock - s.amt) < 0) {
          violations.push({ domeId: s.id, current: dome.stock, requested: s.amt });
        }
      });
      if (violations.length > 0) {
        setValidationAlert({ bargeNo: no, violations });
        return; // block finalization
      }
    }
    
    setDomes((prev) => prev.map((d) => {
      const used = barge.sources.filter((s) => s.id === d.id).reduce((s, x) => s + x.amt, 0);
      return used > 0 ? { ...d, stock: Math.max(0, d.stock + sign * used) } : d;
    }));
    setBarges((prev) => prev.map((b) => b.no === no ? { ...b, finalized: !b.finalized } : b));
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
          const ids = new Set(rows.map((r) => r.id));
          merged = merged.filter((d) => !ids.has(d.id)).concat(rows);
        } catch (err) { console.error(err); }
        processed++;
        if (processed === files.length) {
          setDomes(withInitialStock(merged));
          setImportStatus(`✓ Imported ${merged.length} domes`);
          setTimeout(() => { setShowImport(false); setImportStatus(""); }, 1500);
        }
      };
      reader.readAsText(file);
    });
  };

  return (
    <div className="app">
      <style>{CSS}</style>
      <div className="bg-glow bg-glow-a" />
      <div className="bg-glow bg-glow-b" />

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
              <label className="import-input-label">
                <input type="file" accept=".csv" multiple onChange={handleImport} />
                <span>Select CSV file(s)</span>
              </label>
              {importStatus && <div className="import-status">{importStatus}</div>}
            </div>
          </div>
        </div>
      )}

      {validationAlert && (
        <div className="validation-modal">
          <div className="validation-panel glass">
            <div className="validation-head">
              <AlertTriangle size={20} style={{ color: "#F87171" }} />
              <span>Inventory Mismatch — Cannot Finalize</span>
              <button className="validation-close" onClick={() => setValidationAlert(null)}><X size={18} /></button>
            </div>
            <div className="validation-body">
              <p>Barge #{validationAlert.bargeNo} would cause the following domes to drop below zero. Please verify dome stock data with your data provider:</p>
              <div className="violations-list">
                {validationAlert.violations.map((v) => (
                  <div key={v.domeId} className="violation-item">
                    <span className="violation-dome">{v.domeId}</span>
                    <span className="violation-detail">Current: {fmt(v.current)} WMT | Requested: {fmt(v.requested)} WMT | Deficit: {fmt(v.requested - v.current)} WMT</span>
                  </div>
                ))}
              </div>
              <div className="validation-actions">
                <button className="btn-primary" onClick={() => setValidationAlert(null)}>Dismiss &amp; Update Stock</button>
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
              <Calendar size={11} /> Data updated {fmtShortDate(DATA_LAST_UPDATED)}
            </div>
          </div>
        </div>
        <div className="nav-desktop-wrapper">
          <nav className="nav-desktop">
            <NavButton icon={LayoutGrid} label="Overview" active={tab === "overview"} onClick={() => setTab("overview")} />
            <NavButton icon={Layers} label="Stock" active={tab === "stock"} onClick={() => setTab("stock")} />
            <NavButton icon={Ship} label="Barging Plan" active={tab === "plan"} onClick={() => setTab("plan")} />
            <NavButton icon={Calendar} label="Timeline" active={tab === "timeline"} onClick={() => setTab("timeline")} />
          </nav>
          <button className="btn-import" onClick={() => setShowImport(!showImport)}><Upload size={14} /> Import</button>
          <button className="btn-ghost" onClick={resetDefaults}><RotateCcw size={13} /></button>
        </div>
      </header>

      <main className="content">
        {tab === "overview" && <OverviewTab domes={domes} barges={barges} settings={settings} />}
        {tab === "stock" && <StockTab domes={domes} />}
        {tab === "plan" && <PlanTabWired domes={domes} settings={settings} barges={barges} setBarges={setBarges} toggleFinalize={toggleFinalize} onOpenInvoice={setInvoiceBarge} onExportBarge={setExportBarge} />}
        {tab === "timeline" && <TimelineTab domes={domes} settings={settings} barges={barges} />}
      </main>

      <nav className="nav-mobile">
        <NavButton icon={LayoutGrid} label="Overview" active={tab === "overview"} onClick={() => setTab("overview")} mobile />
        <NavButton icon={Layers} label="Stock" active={tab === "stock"} onClick={() => setTab("stock")} mobile />
        <NavButton icon={Ship} label="Plan" active={tab === "plan"} onClick={() => setTab("plan")} mobile />
        <NavButton icon={Calendar} label="Timeline" active={tab === "timeline"} onClick={() => setTab("timeline")} mobile />
      </nav>
    </div>
  );
}

/* PlanTabWired: thin wrapper so finalize can reach the top-level setDomes */
function PlanTabWired({ domes, settings, barges, setBarges, toggleFinalize, onOpenInvoice, onExportBarge }) {
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
      })
      // Rule 3 (optional) — Si/Mg ratio filter. Only applied when the person actually
      // sets an operator + value; left untouched (no filtering) otherwise.
      .filter((p) => {
        if (!simgActive) return true;
        const d = domesById[p.id];
        const simg = d ? d.simg : undefined;
        if (simg === undefined) return true;
        return genSimgOp === "lte" ? simg <= simgVal : simg >= simgVal;
      });
    const target = Number(genTargetNi) || settings.targetGrade;
    const qty = Number(genQty) > 0 ? Number(genQty) : settings.bargeSize;
    const generated = generateFromPool(genPool, n, qty, target, settings.tolerance);
    const withNo = generated.map((g, i) => ({ no: start + i, shipDate: defaultShipDate(start + i, settings.planTarget), bargeName: "", tugboatName: "", finalized: false, ...g }));
    setBarges((prev) => [...prev, ...withNo]);
  };

  const onUpdate = (no, sources, patch) => {
    setBarges((prev) => prev.map((b) => {
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
    }));
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
        <button className="btn-ghost add-blank-btn" onClick={addBlankBarge}><Plus size={14} /> Add blank barge</button>

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
              <label>Si/Mg ratio (optional)</label>
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
        <div className="note">Generation rule: Targeted Ni is a floor and target+0.2% is a ceiling — barges are never blended outside that range. Lowest-Ni domes are used first as diluent. A barge can draw from any number of existing-inventory contractors, but at most 2 production contractors (IMN-1–4). Domes below 1% or above 1.5% Ni are hard to source, so each is capped at 500 WMT per barge (spread across multiple barges instead). Si/Mg is only applied when you set an operator and value. If there isn't enough on-spec ore left to reach a full barge within all active constraints, the barge is left incomplete rather than filled out of spec.</div>
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
            onUpdate={onUpdate} onImport={onUpdate} onFinalize={toggleFinalize} onRemove={removeBarge} onOpenInvoice={onOpenInvoice} onExportBarge={onExportBarge} />
        ))}
      </section>
    </div>
  );
}

/* ----------------------------- styles ----------------------------- */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@500;600&display=swap');

* { box-sizing: border-box; }
.app {
  min-height: 100vh;
  background: radial-gradient(circle at 20% 0%, #101826 0%, #070A10 55%, #050709 100%);
  color: #EAF0F6;
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  position: relative;
  overflow-x: hidden;
  padding-bottom: 84px;
}
.bg-glow { position: fixed; border-radius: 999px; filter: blur(90px); opacity: .18; pointer-events: none; z-index: 0; }
.bg-glow-a { width: 420px; height: 420px; background: #E35F0C; top: -120px; right: -100px; }
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
.nav-mobile { display: none; }

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
.panel-head { display: flex; align-items: center; gap: 8px; font-family: 'Space Grotesk', sans-serif; font-weight: 600; font-size: 14px; margin-bottom: 16px; color: #EAF0F6; }
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
.btn-export { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 9px;
  background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.18); color: #D6DDE6; font-size: 12px; font-weight: 700;
  cursor: pointer; }
.btn-export:hover { background: rgba(255,255,255,.12); }
.btn-toggle { display: flex; align-items: center; gap: 6px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.15);
  color: #B7C0CC; font-size: 12px; font-weight: 600; padding: 7px 12px; border-radius: 9px; cursor: pointer; }
.btn-toggle-on { background: rgba(74,222,128,.14); border-color: rgba(74,222,128,.4); color: #4ADE80; }

.month-chart { display: flex; align-items: flex-end; gap: 6px; height: 140px; margin-bottom: 8px; }
.month-col { flex: 1; display: flex; flex-direction: column; align-items: center; height: 100%; justify-content: flex-end; }
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

/* Validation alert modal */
.validation-modal { position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 50; background: rgba(0,0,0,.6);
  display: flex; align-items: center; justify-content: center; padding: 20px; }
.validation-panel { max-width: 480px; max-height: 80vh; overflow-y: auto; padding: 0; border-radius: 20px;
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

@media (max-width: 900px) {
  .kpi-row { grid-template-columns: repeat(2, 1fr); }
  .summary-strip { grid-template-columns: repeat(2, 1fr); }
  .contractor-grid { grid-template-columns: repeat(2, 1fr); }
  .targets-grid { grid-template-columns: repeat(2, 1fr); }
  .gen-form-grid { grid-template-columns: repeat(2, 1fr); }
}

@media (max-width: 640px) {
  .nav-desktop, .btn-import { display: none; }
  .nav-mobile {
    display: flex; position: fixed; bottom: 0; left: 0; right: 0; z-index: 10;
    background: rgba(10,14,20,.85); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
    border-top: 1px solid rgba(255,255,255,.08); padding: 8px 10px calc(8px + env(safe-area-inset-bottom));
    justify-content: space-around;
  }
  .navbtn-mobile { flex-direction: column; gap: 3px; font-size: 10px; padding: 6px 10px; flex: 1; }
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
.inv-footer { margin-top: 18pt; text-align: right; font-size: 7pt; color: #000; line-height: 1.5; }

/* Export to PDF button (barging plan filter row) */
.btn-export-pdf { display: inline-flex; align-items: center; gap: 6px; margin-left: auto; padding: 7px 13px; border-radius: 9px;
  background: rgba(227,95,12,.12); border: 1px solid rgba(227,95,12,.35); color: #E35F0C; font-size: 12px; font-weight: 700; cursor: pointer; }
.btn-export-pdf:hover { background: rgba(227,95,12,.22); }

/* Barging Plan export report sheet */
.plan-sheet { background: #ffffff; color: #000000; width: 780px; max-width: 100%; padding: 26pt 22pt;
  font-family: Arial, sans-serif; font-size: 10pt; line-height: 1.4; box-shadow: 0 8px 30px rgba(0,0,0,.4); border-radius: 4px; box-sizing: border-box; }
.plan-header { display: flex; align-items: center; gap: 16pt; margin-bottom: 14pt; }
.plan-logo { height: 50pt; }
.plan-title { font-size: 15pt; font-weight: 700; }
.plan-subtitle { font-size: 9pt; color: #444; margin-top: 2pt; }
.plan-summary { display: flex; flex-wrap: wrap; gap: 18pt; padding: 10pt 0; border-top: 1pt solid #000; border-bottom: 1pt solid #000; margin-bottom: 12pt; }
.plan-summary-item { display: flex; flex-direction: column; gap: 2pt; }
.plan-summary-label { font-size: 8pt; text-transform: uppercase; letter-spacing: .03em; color: #555; }
.plan-summary-value { font-size: 11pt; font-weight: 700; }
.plan-table { width: 100%; border-collapse: collapse; margin-bottom: 14pt; table-layout: fixed; }
.plan-table th, .plan-table td { border: 1pt solid #000; padding: 5pt 6pt; font-size: 9pt; vertical-align: middle; }
.plan-table th { font-weight: 700; text-align: center; background: #999999; }
.plan-center { text-align: center; }
.plan-right { text-align: right; }
.plan-muted { color: #888; font-style: italic; }
.plan-status-badge { display: inline-block; padding: 2pt 8pt; border-radius: 4pt; font-size: 8.5pt; font-weight: 700; border: 1pt solid #000; }
.plan-status-final { background: #d7f4de; }
.plan-status-draft { background: #fdf0d0; }
.plan-footer { margin-top: 10pt; text-align: right; font-size: 7pt; color: #000; line-height: 1.5; }
.plan-footer-bold { font-weight: 700; }

@media print {
  /* Remove the rest of the dashboard from the page flow entirely (display:none, not
   * visibility:hidden) — visibility:hidden keeps the full-height layout intact even
   * though it's invisible, which makes the print engine paginate against the entire
   * dashboard's height instead of just the small modal, producing blank leading pages
   * before the real content shows up. display:none collapses that height to zero. */
  .bg-glow, .topbar, .content, .nav-mobile { display: none !important; }
  .no-print { display: none !important; }

  /* Neutralize the modal's on-screen chrome (dark backdrop, centering, fixed
   * positioning, scroll clipping) so the print-area flows naturally from the top
   * of the page and paginates across as many pages as it actually needs. */
  .invoice-modal { position: static !important; background: none !important; padding: 0 !important; display: block !important; }
  .invoice-panel { position: static !important; box-shadow: none !important; max-height: none !important;
    overflow: visible !important; width: auto !important; max-width: none !important; border-radius: 0 !important; }
  .invoice-body { display: block !important; }
  .invoice-preview-wrap { padding: 0 !important; background: none !important; display: block !important; min-height: 0 !important; }
  .print-area { box-shadow: none !important; }
}
`;
