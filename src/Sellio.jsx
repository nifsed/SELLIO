import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ============================================================================
   SELLIO  -  Analytics & Optimization for Marketplace Brands
   Single-file React. Zero-cost: client-side parsing, IndexedDB persistence.
   Channel 1: Shopee Ads (Data Keseluruhan Iklan CSV) + Product Performance (xlsx).
   ========================================================================== */

/* ---------- IndexedDB (no deps) ------------------------------------------ */
const DB_NAME = "sellio";
const DB_VER = 14;
function openDB() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB_NAME, DB_VER);
    r.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains("snapshots"))
        db.createObjectStore("snapshots", { keyPath: "id" });
      if (!db.objectStoreNames.contains("products"))
        db.createObjectStore("products", { keyPath: "id" });
      if (!db.objectStoreNames.contains("cogs"))
        db.createObjectStore("cogs", { keyPath: "key" });
      if (!db.objectStoreNames.contains("settings"))
        db.createObjectStore("settings", { keyPath: "key" });
      if (!db.objectStoreNames.contains("stock"))
        db.createObjectStore("stock", { keyPath: "code" });
      if (!db.objectStoreNames.contains("income"))
        db.createObjectStore("income", { keyPath: "id" }); // Shopee Income snapshots
      if (!db.objectStoreNames.contains("income_tiktok"))
        db.createObjectStore("income_tiktok", { keyPath: "id" }); // TikTok Income snapshots
      if (!db.objectStoreNames.contains("orders"))
        db.createObjectStore("orders", { keyPath: "id" }); // Shopee Order export
      if (!db.objectStoreNames.contains("cogs_items"))
        db.createObjectStore("cogs_items", { keyPath: "sku" }); // COGS per SKU
      if (!db.objectStoreNames.contains("meta_ads"))
        db.createObjectStore("meta_ads", { keyPath: "id" }); // Meta Ads snapshots
      if (!db.objectStoreNames.contains("tiktok_ads"))
        db.createObjectStore("tiktok_ads", { keyPath: "id" }); // TikTok Ads (GMV Max) snapshots
      if (!db.objectStoreNames.contains("tiktok_campaign"))
        db.createObjectStore("tiktok_campaign", { keyPath: "id" }); // TikTok per-campaign/product
      if (!db.objectStoreNames.contains("pnl_inputs"))
        db.createObjectStore("pnl_inputs", { keyPath: "key" }); // P&L manual inputs
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function idbPut(store, val) {
  const db = await openDB();
  return new Promise((res, rej) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).put(val);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
}
async function idbAll(store) {
  const db = await openDB();
  return new Promise((res, rej) => {
    const tx = db.transaction(store, "readonly");
    const req = tx.objectStore(store).getAll();
    req.onsuccess = () => res(req.result);
    req.onerror = () => rej(req.error);
  });
}
async function idbDel(store, key) {
  const db = await openDB();
  return new Promise((res, rej) => {
    const tx = db.transaction(store, "readwrite");
    tx.objectStore(store).delete(key);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
}

/* ---------- Parsing helpers ---------------------------------------------- */
const num = (s) => {
  if (s == null) return 0;
  const t = String(s).trim();
  if (t === "" || t === "-") return 0;
  return parseFloat(t.replace(/%/g, "").replace(/\./g, (m, i, str) =>
    // keep decimal point only if it's a ROAS-style "x.yy"; Shopee uses "." as decimal here
    "."
  ).replace(/[^0-9.\-]/g, "")) || 0;
};
const pct = (s) => {
  if (s == null) return 0;
  const t = String(s).trim().replace("%", "").replace(",", ".");
  return parseFloat(t) || 0;
};
// integer rupiah fields use no thousands separators in this export
const intnum = (s) => {
  if (s == null) return 0;
  const t = String(s).trim();
  if (t === "" || t === "-") return 0;
  return parseInt(t.replace(/[^0-9\-]/g, ""), 10) || 0;
};

function parseCSV(text) {
  // robust-enough CSV split (handles quoted commas)
  const rows = [];
  let cur = [], field = "", inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') inQ = false;
      else field += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ",") { cur.push(field); field = ""; }
      else if (c === "\n") { cur.push(field); rows.push(cur); cur = []; field = ""; }
      else if (c === "\r") { /* skip */ }
      else field += c;
    }
  }
  if (field.length || cur.length) { cur.push(field); rows.push(cur); }
  return rows;
}

function parseShopee(text) {
  const rows = parseCSV(text);
  const meta = {};
  let headerIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const r0 = (rows[i][0] || "").trim();
    if (r0 === "Urutan") { headerIdx = i; break; }
    if (rows[i].length >= 2 && r0) meta[r0] = rows[i][1];
  }
  if (headerIdx === -1) throw new Error("Header 'Urutan' tidak ditemukan  -  pastikan ini export Data Keseluruhan Iklan Shopee.");
  const header = rows[headerIdx].map((h) => h.trim());
  const col = (name) => header.indexOf(name);
  const idx = {
    name: col("Nama Iklan"),
    status: col("Status"),
    type: col("Jenis Iklan"),
    code: col("Kode Produk"),
    bidding: col("Mode Bidding"),
    start: col("Tanggal Mulai"),
    impr: col("Dilihat"),
    clicks: col("Jumlah Klik"),
    ctr: col("Persentase Klik"),
    conv: col("Konversi"),
    cvr: col("Tingkat konversi"),
    cpa: col("Biaya per Konversi"),
    sold: col("Produk Terjual"),
    gmv: col("Omzet Penjualan"),
    spend: col("Biaya"),
    roas: col("Efektifitas Iklan"),
    acos: col("Persentase Biaya Iklan terhadap Penjualan dari Iklan (ACOS)"),
  };
  const ads = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[idx.name] || !r[idx.name].trim()) continue;
    ads.push({
      name: r[idx.name].trim(),
      status: (r[idx.status] || "").trim(),
      type: (r[idx.type] || "").trim() || " - ",
      code: (r[idx.code] || "").trim(),
      bidding: (r[idx.bidding] || "").trim() || " - ",
      start: (r[idx.start] || "").trim(),
      impr: intnum(r[idx.impr]),
      clicks: intnum(r[idx.clicks]),
      ctr: pct(r[idx.ctr]),
      conv: intnum(r[idx.conv]),
      cvr: pct(r[idx.cvr]),
      cpa: num(r[idx.cpa]),
      sold: intnum(r[idx.sold]),
      gmv: intnum(r[idx.gmv]),
      spend: intnum(r[idx.spend]),
      roas: num(r[idx.roas]),
      acos: pct(r[idx.acos]),
    });
  }
  // derive period end date from "Periode" meta "DD/MM/YYYY - DD/MM/YYYY"
  let periodEnd = null, periodStart = null;
  if (meta["Periode"]) {
    const m = meta["Periode"].split("-").map((s) => s.trim());
    periodStart = m[0]; periodEnd = m[1] || m[0];
  }
  return { meta, ads, periodStart, periodEnd, store: meta["Nama Toko"] || "Toko" };
}

/* parse DD/MM/YYYY -> Date */
function dmy(s) {
  if (!s) return null;
  const [d, m, y] = s.split("/").map((x) => parseInt(x, 10));
  if (!d || !m || !y) return null;
  return new Date(y, m - 1, d);
}

/* ---------- Product performance parser (Shopee xlsx) --------------------- */
// "Parent SKU Detail" export. Reads only parent rows (Kode Variasi == '-').
// Indonesian number format: "34.896.000" = 34896000 ; "1,14%" = 1.14
const idNum = (s) => {
  if (s == null) return 0;
  const t = String(s).trim();
  if (t === "" || t === "-") return 0;
  // remove thousand dots, treat comma as decimal
  return parseFloat(t.replace(/\./g, "").replace(",", ".").replace(/[^0-9.\-]/g, "")) || 0;
};
const idPct = (s) => {
  if (s == null) return 0;
  const t = String(s).trim().replace("%", "").replace(",", ".");
  if (t === "" || t === "-") return 0;
  return parseFloat(t) || 0;
};

function parseProductXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  // best-performing sheet holds the full product list
  const sheetName = wb.SheetNames.find((n) => /performa terbaik/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  if (!rows.length) throw new Error("Sheet kosong.");
  const header = rows[0].map((h) => String(h).trim());
  const col = (name) => header.indexOf(name);
  const dataStart = 1;
  const idx = {
    code: col("Kode Produk"),
    name: col("Produk"),
    status: col("Status Produk Saat Ini"),
    varCode: col("Kode Variasi"),
    sales: col("Total Penjualan (Pesanan Dibuat) (IDR)"),
    views: col("Jumlah Produk Dilihat"),
    clicks: col("Produk Diklik"),
    ctr: col("Persentase Klik"),
    cvr: col("Tingkat Konversi Pesanan (Pesanan Dibuat)"),
    unitsSold: col("Produk (Pesanan Dibuat)"),
    orders: col("Pesanan Dibuat"),
    aov: col("Penjualan per Pesanan (Pesanan Dibuat) (IDR)"),
    bounceRate: col("Tingkat Pengunjung Melihat Tanpa Membeli"),
    atcRate: col("Tingkat Konversi Produk Dimasukkan ke Keranjang"),
    repeatRate: col("Tingkat Pesanan Berulang (Pesanan Dibuat)"),
  };
  if (idx.code === -1 || idx.sales === -1)
    throw new Error("Kolom inti tidak ditemukan  -  pastikan ini export 'Detail SKU Induk' dari Shopee.");

  const products = [];
  for (let i = dataStart; i < rows.length; i++) {
    const r = rows[i];
    const code = String(r[idx.code] ?? "").trim();
    if (!code || code === "-") continue;
    // parent rows only: variation code is '-' (avoid double counting sizes)
    const varCode = String(r[idx.varCode] ?? "").trim();
    if (varCode && varCode !== "-") continue;
    products.push({
      code,
      name: String(r[idx.name] ?? "").trim(),
      status: String(r[idx.status] ?? "").trim(),
      sales: idNum(r[idx.sales]),
      views: idNum(r[idx.views]),
      clicks: idNum(r[idx.clicks]),
      ctr: idPct(r[idx.ctr]),
      cvr: idPct(r[idx.cvr]),
      unitsSold: idNum(r[idx.unitsSold]),
      orders: idNum(r[idx.orders]),
      aov: idNum(r[idx.aov]),
      bounceRate: idPct(r[idx.bounceRate]),
      atcRate: idPct(r[idx.atcRate]),
      repeatRate: idPct(r[idx.repeatRate]),
    });
  }
  return { products, sheetName };
}

/* derive period from filename: parentskudetail_YYYYMMDD_YYYYMMDD.xlsx */
function periodFromFilename(name) {
  const m = name.match(/(\d{8})[_-](\d{8})/);
  if (!m) return { start: null, end: null };
  const fmt = (s) => `${s.slice(6, 8)}/${s.slice(4, 6)}/${s.slice(0, 4)}`;
  return { start: fmt(m[1]), end: fmt(m[2]) };
}

/* ---------- BCG classification (share × CVR) ----------------------------- */
// X = net revenue share (fee-adjusted if available), Y = blended CVR (product + ad signal)
// adMap: { [code]: adObj } from active snapshot
// feeRate: actual fee % from income file (or 0 if not available)
function classifyBCG(products, adMap, feeRate) {
  adMap = adMap || {};
  feeRate = feeRate || 0;

  // Enrich each product with ad data + compute net revenue
  const enriched = products.map(p => {
    const ad = adMap[p.code] || null;
    // Net revenue = sales * (1 - fee rate) — more accurate basis for share
    const netSales = feeRate > 0 ? p.sales * (1 - feeRate/100) : p.sales;
    // Blended CVR: weight product CVR 60% + ad CVR 40% if available
    const adCvr = ad && ad.conv && ad.clicks ? (ad.conv / ad.clicks * 100) : 0;
    const blendedCvr = adCvr > 0 ? p.cvr * 0.6 + adCvr * 0.4 : p.cvr;
    // ROAS signal
    const roas = ad ? ad.roas : null;
    const adGmv = ad ? ad.gmv : 0;
    const adSpend = ad ? ad.spend : 0;
    return { ...p, netSales, blendedCvr, roas, adGmv, adSpend, ad };
  });

  const totalNet = enriched.reduce((t, p) => t + p.netSales, 0) || 1;
  const withShare = enriched.map(p => ({ ...p, share: (p.netSales / totalNet) * 100 }));

  const shares = withShare.map(p => p.share).sort((a, b) => a - b);
  const cvrs   = withShare.map(p => p.blendedCvr).filter(x => x > 0).sort((a, b) => a - b);
  const med    = arr => arr.length ? arr[Math.floor(arr.length / 2)] : 0;
  const meanShare = 100 / Math.max(withShare.length, 1);
  const shareCut  = Math.max(med(shares), meanShare * 1.5, 2);
  const cvrCut    = Math.max(med(cvrs), 0.8);

  return withShare.map(p => {
    const hiShare = p.share >= shareCut;
    const hiCvr   = p.blendedCvr >= cvrCut;
    let quadrant;
    if (hiShare && hiCvr)   quadrant = "star";
    else if (hiShare)        quadrant = "cashcow";
    else if (hiCvr)          quadrant = "question";
    else                     quadrant = "dog";
    return { ...p, quadrant, hiShare, hiCvr };
  }).sort((a, b) => b.sales - a.sales);
}
const QUAD = {
  star: { label: "Star", color: "#3fb950", desc: "Share tinggi · CVR tinggi" },
  cashcow: { label: "Cash Cow", color: "#58a6ff", desc: "Share tinggi · CVR rendah" },
  question: { label: "Question Mark", color: "#d29922", desc: "Share rendah · CVR tinggi" },
  dog: { label: "Dog", color: "#f85149", desc: "Share rendah · CVR rendah" },
};

/* ---------- DIO + stock-aware ad recommendation -------------------------- */
// periodDays: derived from product file filename (e.g. 30 days)
function calcDio(units, unitsSold, periodDays) {
  if (!units || !unitsSold || !periodDays) return null;
  const dailyRate = unitsSold / periodDays;
  if (!dailyRate) return null;
  return Math.round(units / dailyRate); // days until stockout at current rate
}

// DIO thresholds
const DIO_OVERSTOCK = 90;  // > 90 days = overstock
const DIO_SAFE     = 30;   // 30-90 days = safe zone
const DIO_LOW      = 14;   // 14-30 days = watch
const DIO_CRITICAL = 14;   // < 14 days = critical

function dioLabel(dio) {
  if (dio === null) return null;
  if (dio > DIO_OVERSTOCK) return { tag: "OVERSTOCK",  color: "#58a6ff", level: "overstock" };
  if (dio > DIO_SAFE)      return { tag: "AMAN",       color: "#3fb950", level: "safe" };
  if (dio > DIO_LOW)       return { tag: "WATCH",      color: "#d29922", level: "low" };
  return                          { tag: "KRITIS",     color: "#f85149", level: "critical" };
}

// Decision tree: stock condition × ad performance → recommendation
/* ---------- Seasonal multipliers ----------------------------------------- */
// Default Indonesia fashion marketplace seasonality (editable in UI)
const DEFAULT_SEASONAL = {
  1: 1.0, 2: 0.9, 3: 3.5, 4: 0.6, 5: 1.0,
  6: 1.1, 7: 1.0, 8: 1.0, 9: 1.0, 10: 1.0, 11: 1.2, 12: 1.8,
};

// Forecast demand for next N days from today's date, applying seasonal multiplier
function forecastDemand(dailyRate, periodDays, days = 30) {
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const month = nextMonth.getMonth() + 1; // 1-12
  const mult = DEFAULT_SEASONAL[month] ?? 1.0;
  return { units: Math.round(dailyRate * days * mult), mult, month };
}

// stockAdRec  -  table pill only (short label). Full reasoning is in buildDecision.
function stockAdRec(dio, roas, targetRoas, quadrant, ad, thresholds) {
  const dioLvl = dioLabel(dio)?.level;
  if (!dioLvl) return null;
  const tr = targetRoas;
  const roasStrong  = roas >= tr * 1.5;
  const roasOk      = roas >= tr;
  const roasDead    = roas < tr * 0.75;
  const ctrLow      = ad ? ad.ctr < (thresholds?.minCtr || 2) : false;
  const cvrLow      = ad ? ad.cvr < (thresholds?.minCvr || 1) : false;

  if (dioLvl === "critical") {
    if (roasDead || (ctrLow && cvrLow)) return { action: "PAUSE + RESTOCK", level: "bad" };
    if (roasStrong) return { action: "TURUN BUDGET", level: "watch" };
    return { action: "KURANGI BUDGET", level: "watch" };
  }
  if (dioLvl === "low") {
    if (roasDead) return { action: "PAUSE + PO", level: "bad" };
    if (roasStrong) return { action: "NAIKKAN ROAS TARGET", level: "watch" };
    return { action: "KURANGI BUDGET", level: "watch" };
  }
  if (dioLvl === "overstock") {
    if (roasDead) return { action: "STOP IKLAN", level: "bad" };
    return { action: "REM BUDGET", level: "watch" };
  }
  // safe
  if (roasStrong && (quadrant === "star" || quadrant === "question")) return { action: "SCALE", level: "good" };
  if (roasStrong && quadrant === "cashcow") return { action: "FIX CVR DULU", level: "watch" };
  if (roasOk) return { action: "MAINTAIN", level: "good" };
  if (roasDead) return { action: "STOP & AUDIT", level: "bad" };
  return { action: "OPTIMALKAN", level: "watch" };
}

// periodDays from product snapshot filename or fallback 30
function periodDaysFromSnap(snap) {
  if (!snap) return 30;
  const m = (snap.fileName || "").match(/(\d{8})[_-](\d{8})/);
  if (!m) return 30;
  const a = m[1], b = m[2];
  const da = new Date(+a.slice(0,4), +a.slice(4,6)-1, +a.slice(6,8));
  const db2 = new Date(+b.slice(0,4), +b.slice(4,6)-1, +b.slice(6,8));
  const diff = Math.round((db2 - da) / 86400000);
  return diff > 0 ? diff : 30;
}

/* diagnose a single product based on quadrant + metrics + thresholds */
function diagnoseProduct(p, th) {
  const out = [];
  const q = QUAD[p.quadrant];
  if (p.quadrant === "star") {
    out.push({ level: "good", title: "Star  -  mesin revenue", msg: `Menyumbang ${p.share.toFixed(1)}% revenue (${rp(p.sales)}) dengan CVR ${(p.blendedCvr||p.cvr).toFixed(2)}%. Produk terpenting sekarang. Prioritas: jaga stok jangan stockout, scale paid ads (ROAS sudah layak), jadikan template foto/copy untuk SKU lain.` });
    if (p.share > 25) out.push({ level: "watch", title: "Konsentrasi revenue tinggi", msg: `${p.share.toFixed(0)}% revenue dari satu produk  -  risiko sistemik kalau drop. Mulai develop SKU lain jadi Star kedua.` });
  } else if (p.quadrant === "cashcow") {
    out.push({ level: "watch", title: "Cash Cow  -  konversi bocor", msg: `Share ${p.share.toFixed(1)}% tinggi tapi CVR ${(p.blendedCvr||p.cvr).toFixed(2)}% di bawah median. Traffic & exposure sudah ada  -  upside terbesar ada di fix konversi: harga, foto, deskripsi, review. Naikin CVR di sini = revenue naik tanpa nambah spend.` });
    if (p.bounceRate > 20) out.push({ level: "watch", title: "Bounce tinggi", msg: `${p.bounceRate.toFixed(0)}% pengunjung lihat tanpa beli. Cek 3 hal pertama yang dilihat pembeli: harga, foto utama, ongkir.` });
  } else if (p.quadrant === "question") {
    out.push({ level: "watch", title: "Question Mark  -  bagus tapi sepi", msg: `CVR ${(p.blendedCvr||p.cvr).toFixed(2)}% bagus tapi share cuma ${p.share.toFixed(1)}%  -  produk meyakinkan yang beli, tapi kurang exposure. Kandidat kuat untuk di-push iklan: konversi sudah terbukti, tinggal kasih traffic.` });
  } else {
    out.push({ level: "bad", title: "Dog  -  evaluasi", msg: `Share ${p.share.toFixed(1)}% & CVR ${(p.blendedCvr||p.cvr).toFixed(2)}% dua-duanya rendah. Jangan habiskan resource iklan di sini. Pilihan: perbaiki listing dulu (foto/harga), bundling dengan Star, atau clearance kalau stok mati.` });
  }
  // cross-cut signals regardless of quadrant
  if (p.views > 5000 && p.ctr < th.minCtr) {
    out.push({ level: "watch", title: "CTR listing rendah", msg: `${nfmt(p.views)} kali dilihat tapi CTR ${p.ctr.toFixed(2)}%  -  thumbnail/judul/harga di hasil pencarian kurang menarik klik.` });
  }
  return out;
}

/* ---------- Threshold defaults ------------------------------------------- */
// Editable benchmark thresholds (persisted). targetRoas can be auto-derived from margin.
const DEFAULT_TH = {
  targetRoas: 8,   // x   -  desired ROAS floor (Brand OS uses 8 as default)
  minCtr: 2,       // %   -  below = weak creative / placement
  minCvr: 1,       // %   -  below = landing/price/stock issue
  maxCpa: 50000,   // Rp  -  above = acquisition too expensive
  autoRoasFromMargin: true, // if true & margin known, targetRoas = max(default, BE ROAS)
  minImprForCtr: 1000,
  minClicksForCvr: 200,
};

/* ---------- Diagnosis engine (threshold-driven) -------------------------- */
// Returns array of {level:'good'|'watch'|'bad', tag, msg} given metrics, margin, thresholds.
function diagnose(ad, marginPct, th = DEFAULT_TH) {
  const isGMVMax = /GMV Max/i.test(ad.bidding);
  const out = [];

  // Resolve effective target ROAS: explicit threshold, optionally raised to break-even
  let targetRoas = th.targetRoas;
  let beRoas = null;
  if (marginPct != null) {
    beRoas = 100 / marginPct;
    if (th.autoRoasFromMargin) targetRoas = Math.max(th.targetRoas, beRoas);
  }

  // 1) ROAS vs target  (primary profitability signal)
  if (ad.roas >= targetRoas * 1.5 && ad.spend > 300000) {
    out.push({ level: "good", tag: "SCALE", msg: `ROAS ${ad.roas.toFixed(2)}x jauh di atas target ${targetRoas.toFixed(1)}x dengan spend ${rpShort(ad.spend)}. Sinyal kuat untuk scale${isGMVMax ? "  -  naikkan budget bertahap atau longgarkan target ROAS GMV Max (jangan ubah materi yang menang)" : "  -  naikkan budget bertahap"}.` });
  } else if (ad.roas >= targetRoas) {
    out.push({ level: "good", tag: "ON TARGET", msg: `ROAS ${ad.roas.toFixed(2)}x memenuhi target ${targetRoas.toFixed(1)}x. Pertahankan, pantau drift harian.` });
  } else {
    out.push({ level: "bad", tag: "DI BAWAH TARGET", msg: `ROAS ${ad.roas.toFixed(2)}x di bawah target ${targetRoas.toFixed(1)}x${beRoas != null ? ` (break-even ${beRoas.toFixed(1)}x)` : ""}. ${isGMVMax ? "Ketatkan target ROAS GMV Max 1 langkah, atau pause kalau konsisten merah" : "Turunkan budget 20–30% dan perbaiki materi/bid"}. Jangan tambah budget untuk cari volume di iklan rugi.` });
  }

  // 2) Profitability vs margin (ACOS)
  if (marginPct != null) {
    if (ad.acos > marginPct) {
      out.push({ level: "bad", tag: "MAKAN MARGIN", msg: `ACOS ${ad.acos.toFixed(1)}% melebihi margin ${marginPct.toFixed(0)}%  -  tiap order dari iklan ini rugi sebelum biaya lain. Prioritas dikoreksi.` });
    } else if (ad.acos > marginPct * 0.8) {
      out.push({ level: "watch", tag: "MARGIN TIPIS", msg: `ACOS ${ad.acos.toFixed(1)}% mendekati margin ${marginPct.toFixed(0)}%  -  ruang profit menipis, hati-hati kalau mau scale.` });
    }
  }

  // 3) CTR  -  top of funnel / creative
  if (ad.impr >= th.minImprForCtr) {
    if (ad.ctr < th.minCtr) {
      out.push({ level: "watch", tag: "CTR RENDAH", msg: `CTR ${ad.ctr.toFixed(2)}% di bawah ambang ${th.minCtr}%  -  materi/penempatan kurang menarik perhatian. Ganti visual atau hook; uji 1 variasi baru sebelum dorong budget.` });
    } else if (ad.ctr >= th.minCtr * 2) {
      out.push({ level: "good", tag: "CTR KUAT", msg: `CTR ${ad.ctr.toFixed(2)}%  -  materi efektif menarik perhatian. Jadikan template untuk SKU lain.` });
    }
  }

  // 4) CVR  -  bottom of funnel / landing
  if (ad.clicks >= th.minClicksForCvr) {
    if (ad.cvr < th.minCvr) {
      out.push({ level: "watch", tag: "CVR RENDAH", msg: `CVR ${ad.cvr.toFixed(2)}% di bawah ambang ${th.minCvr}%  -  traffic masuk tapi nggak closing. Ini bukan masalah iklan: cek harga, stok, foto produk, review. Hindari scale sampai CVR membaik.` });
    } else if (ad.cvr >= th.minCvr * 1.6) {
      out.push({ level: "good", tag: "CVR KUAT", msg: `CVR ${ad.cvr.toFixed(2)}%  -  halaman & harga meyakinkan pembeli. Sinyal aman untuk scale.` });
    }
  }

  // 5) CPA  -  acquisition cost
  if (ad.cpa > 0 && ad.conv > 0 && ad.cpa > th.maxCpa) {
    out.push({ level: "watch", tag: "CPA MAHAL", msg: `CPA ${rp(ad.cpa)} di atas ambang ${rp(th.maxCpa)}. ${isGMVMax ? "GMV Max menentukan harga konversi otomatis  -  kalau CPA tinggi tapi ROAS oke berarti AOV besar, tidak masalah. Cek konteksnya." : "Pertimbangkan long-tail keyword lebih spesifik atau jadwal di jam sepi."}` });
  }

  // 6) Manual bidding hint
  if (!isGMVMax && ad.conv >= 10) {
    out.push({ level: "watch", tag: "MANUAL", msg: `Mode bidding manual dengan ${ad.conv} konversi  -  histori cukup untuk pertimbangkan migrasi ke GMV Max.` });
  }

  // NOTE: stock-based advice (e.g. "11 hari stok tersisa") needs an inventory source.
  // Shopee "Data Keseluruhan Iklan" export has no stock column. Placeholder for future import.

  return out;
}
function worstLevel(dx) {
  if (dx.some((d) => d.level === "bad")) return "bad";
  if (dx.some((d) => d.level === "watch")) return "watch";
  return "good";
}

/* ---------- Formatting ---------------------------------------------------- */
const rp = (n) => "Rp" + Math.round(n).toLocaleString("id-ID");
const rpShort = (n) => {
  const a = Math.abs(n);
  if (a >= 1e9) return "Rp" + (n / 1e9).toFixed(1) + "M";
  if (a >= 1e6) return "Rp" + (n / 1e6).toFixed(1) + "jt";
  if (a >= 1e3) return "Rp" + (n / 1e3).toFixed(0) + "rb";
  return rp(n);
};
const nfmt = (n) => Math.round(n).toLocaleString("id-ID");


/* ============================================================================
   INCOME (PENGHASILAN) PARSER  -  Shopee "Laporan Penghasilan" XLSX
   Sheet "Summary"  → waterfall totals
   Sheet "Income"   → per-order detail (row 6 = headers, row 7+ = data)
   ========================================================================== */
function parseIncomeXlsx(buf) {
  const wb = XLSX.read(buf, { type: "array" });

  // --- Summary sheet ---
  // Read cell-by-cell to avoid sheet_to_json collapsing sparse columns.
  // Shopee Summary has label in col A, value in col D (skips B and C).
  const sumSh = wb.Sheets["Summary"];

  function getCellVal(sh, addr) {
    const c = sh[addr];
    if (!c) return undefined;
    return c.v; // raw value (number or string)
  }
  function sheetToRows(sh) {
    // Convert sheet to array-of-arrays preserving all columns including empty ones
    const range = XLSX.utils.decode_range(sh["!ref"] || "A1:Z100");
    const rows = [];
    for (let R = range.s.r; R <= range.e.r; R++) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = sh[addr];
        row.push(cell ? cell.v : null);
      }
      rows.push(row);
    }
    return rows;
  }
  const sumRows = sheetToRows(sumSh);

  function findVal(rows, label, partial) {
    for (const r of rows) {
      for (let i = 0; i < r.length; i++) {
        const cell = String(r[i] ?? "").trim();
        if (!cell) continue;
        const match = partial ? cell.includes(label) : cell === label;
        if (match) {
          // Strategy: scan RIGHT-TO-LEFT for last numeric in row (handles col D value pattern)
          // Then also try left-to-right for col C value pattern
          // Take whichever is non-zero, prefer rightmost
          let rightVal = null;
          for (let j = r.length - 1; j > i; j--) {
            const v = r[j];
            if (v === null || v === undefined) continue;
            const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
            if (!isNaN(n) && n !== 0) { rightVal = n; break; }
          }
          if (rightVal !== null) return rightVal;
          // Fallback: left-to-right first numeric
          for (let j = i + 1; j < r.length; j++) {
            const v = r[j];
            if (v === null || v === undefined) continue;
            const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
            if (!isNaN(n)) return n;
          }
        }
      }
    }
    return 0;
  }

  // Period
  let periodStart = "", periodEnd = "", seller = "";
  for (const r of sumRows) {
    if (String(r[0] ?? "").trim() === "Username (Penjual)") seller = String(r[1] || "").trim();
    if (String(r[0] ?? "").trim() === "Dari") periodStart = String(r[1] || "").trim();
    if (String(r[0] ?? "").trim() === "ke") periodEnd = String(r[1] || "").trim();
  }

  const summary = {
    hargaAsli:      findVal(sumRows, "Harga Asli Produk"),
    diskonProduk:   findVal(sumRows, "Total Diskon Produk"),
    refund:         findVal(sumRows, "Jumlah Pengembalian Dana ke Pembeli"),
    voucherSeller:  findVal(sumRows, "Voucher disponsor oleh Penjual"),
    totalPendapatan:findVal(sumRows, "Total Pendapatan", true),
    biayaOngkir:    findVal(sumRows, "Total Biaya Pengiriman"),
    ongkirBuyer:    findVal(sumRows, "Ongkir Dibayar Pembeli"),
    gratisOngkirShopee: findVal(sumRows, "Gratis Ongkir dari Shopee"),
    ongkirDiteruskan:   findVal(sumRows, "Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim"),
    promoOngkirSeller:  findVal(sumRows, "Promo Gratis Ongkir dari Penjual"),
    totalBiayaAdmin:findVal(sumRows, "Biaya Admin & Layanan"),
    komisiAMS:      findVal(sumRows, "Biaya Komisi AMS"),
    biayaAdmin:     findVal(sumRows, "Biaya Administrasi (termasuk PPN 11%)"),
    biayaLayanan:   findVal(sumRows, "Biaya Layanan"),
    biayaProses:    findVal(sumRows, "Biaya Proses Pesanan"),
    totalPengeluaran: findVal(sumRows, "Total Pengeluaran", true),
    totalDilepas:   findVal(sumRows, "Total yang Dilepas", true),
  };

  // Derive net GMV
  summary.netGmv = summary.hargaAsli + summary.diskonProduk + summary.refund + summary.voucherSeller;
  summary.totalFee = summary.komisiAMS + summary.biayaAdmin + summary.biayaLayanan + summary.biayaProses;
  summary.feeRateNetGmv = summary.netGmv ? Math.abs(summary.totalFee) / summary.netGmv * 100 : 0;
  summary.feeRateGross = summary.hargaAsli ? Math.abs(summary.totalFee) / summary.hargaAsli * 100 : 0;
  summary.takeRate = summary.netGmv ? (1 - summary.totalDilepas / summary.netGmv) * 100 : 0;

  // --- Income sheet (per-order) ---
  const incSh = wb.Sheets["Income"];
  let orders = [];
  if (incSh) {
    const rows = XLSX.utils.sheet_to_json(incSh, { header: 1, defval: "" });
    // Find header row (contains "No. Pesanan")
    let hdrIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].some(c => String(c).trim() === "No. Pesanan")) { hdrIdx = i; break; }
    }
    if (hdrIdx >= 0) {
      const hdrs = rows[hdrIdx].map(c => String(c).trim());
      const col = (name) => hdrs.indexOf(name);
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[col("No. Pesanan")] || r[col("No. Pesanan")] === "No. Pesanan") continue;
        const n = (idx) => { const v = parseFloat(String(r[idx] ?? "").replace(/[^0-9.\-]/g, "")); return isNaN(v) ? 0 : v; };
        orders.push({
          noPesanan:      String(r[col("No. Pesanan")] || ""),
          tglDibuat:      String(r[col("Waktu Pesanan Dibuat")] || ""),
          tglDilepas:     String(r[col("Tanggal Dana Dilepaskan")] || ""),
          metodePembayaran: String(r[col("Metode pembayaran pembeli")] || ""),
          hargaAsli:      n(col("Harga Asli Produk")),
          diskonProduk:   n(col("Total Diskon Produk")),
          refund:         n(col("Jumlah Pengembalian Dana ke Pembeli")),
          pengembalianDana: n(col("Pengembalian Dana ke Pembeli")),
          komisiAMS:      n(col("Biaya Komisi AMS")),
          biayaAdmin:     n(col("Biaya Administrasi (termasuk PPN 11%)")),
          biayaLayanan:   n(col("Biaya Layanan")),
          biayaProses:    n(col("Biaya Proses Pesanan")),
          totalDilepas:   n(col("Total Penghasilan")),
        });
      }
    }
  }

  // --- Refund metrics from per-order data ---
  const refundOrders = orders.filter(o => o.pengembalianDana < 0);
  summary.refundOrderCount = refundOrders.length;
  summary.refundValue = refundOrders.reduce((t, o) => t + Math.abs(o.pengembalianDana), 0);
  summary.totalOrderCount = orders.length;
  summary.refundRateOrder = orders.length ? (refundOrders.length / orders.length) * 100 : 0;
  summary.refundRateValue = summary.hargaAsli ? (summary.refundValue / summary.hargaAsli) * 100 : 0;

  return { seller, periodStart, periodEnd, summary, orders };
}


/* ============================================================================
   TIKTOK SHOP INCOME PARSER  -  "Detail pesanan" single sheet
   Columns: ID Pesanan, Jenis transaksi, Waktu pemesanan, Waktu pembayaran pesanan,
   Mata uang, Jumlah penyelesaian pembayaran, Total Pendapatan,
   Subtotal setelah diskon penjual, Subtotal sebelum diskon, Diskon penjual,
   Total Biaya, Biaya komisi platform, Komisi Afiliasi, Komisi dinamis,
   Biaya pemrosesan pesanan, Ongkir, Ongkir yang ditanggung platform, ...
   ========================================================================== */
function parseTikTokIncomeXlsx(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const sh = wb.Sheets["Detail pesanan"];
  if (!sh) throw new Error("Sheet 'Detail pesanan' tidak ditemukan — pastikan file TikTok Shop yang benar.");

  function readSheetRows(sheet) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:CB1000");
    const out = [];
    for (let R = range.s.r; R <= range.e.r; R++) {
      const row = [];
      for (let CC = range.s.c; CC <= range.e.c; CC++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: CC });
        const cell = sheet[addr];
        row.push(cell ? cell.v : "");
      }
      out.push(row);
    }
    return out;
  }

  const n = (v) => {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    const s = String(v).replace(/[^0-9.\-]/g, "");
    const x = parseFloat(s);
    return isNaN(x) ? 0 : x;
  };

  // ── 1. Read Laporan sheet for authoritative summary ──
  let laporan = {};
  let periodStart = "", periodEnd = "";
  const lsh = wb.Sheets["Laporan"];
  if (lsh) {
    const lrows = readSheetRows(lsh);
    lrows.forEach(row => {
      const label = String(row[0] || "").trim();
      const label2 = String(row[1] || "").trim();
      const val1 = n(row[1]);
      const val2 = n(row[2]);
      if (label) laporan[label] = val1 || val2;
      if (label2 && label2 !== label) laporan[label2] = val2 || val1;
    });
    // Parse period
    const periodeRow = lrows.find(r => String(r[0]).trim() === "Periode");
    if (periodeRow) {
      const parts = String(periodeRow[1] || "").split("-");
      if (parts.length === 2) {
        periodStart = parts[0].trim().replace(/\//g, "-");
        periodEnd   = parts[1].trim().replace(/\//g, "-");
      }
    }
  }

  // ── 2. Read Detail pesanan for per-order breakdown ──
  const rows = readSheetRows(sh);
  if (!rows.length) throw new Error("File kosong.");
  const hdrs = rows[0].map(c => String(c ?? "").trim());
  const col = (name) => hdrs.indexOf(name);

  const orders = [], adPayments = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[col("ID Pesanan/Penyesuaian")]) continue;
    const jenis = String(r[col("Jenis transaksi")] || "").trim();
    const tgl = String(r[col("Waktu pemesanan")] || "").slice(0, 10);
    if (tgl && !periodStart) {
      if (!periodStart || tgl < periodStart) periodStart = tgl;
      if (!periodEnd   || tgl > periodEnd)   periodEnd   = tgl;
    }
    if (jenis === "Pesanan") {
      orders.push({
        id:             String(r[col("ID Pesanan/Penyesuaian")] || ""),
        tglPesanan:     tgl,
        tglBayar:       String(r[col("Waktu pembayaran pesanan")] || "").slice(0,10),
        settlement:     n(r[col("Jumlah penyelesaian pembayaran")]),
        totalPendapatan:n(r[col("Total Pendapatan")]),
        subtotalBruto:  n(r[col("Subtotal sebelum diskon")]),
        subtotalRefund: n(r[col("Subtotal pengembalian dana sebelum diskon penjual")]),
        diskonPenjual:  n(r[col("Diskon penjual")]),
        totalBiaya:     n(r[col("Total Biaya")]),
        komisiPlatform: n(r[col("Biaya komisi platform")]),
        komisiAfiliasi: n(r[col("Komisi Afiliasi")]),
        komisiDinamis:  n(r[col("Komisi dinamis")]),
        biayaProses:    n(r[col("Biaya pemrosesan pesanan")]),
        ongkir:         n(r[col("Ongkir")]),
        ongkirPlatform: n(r[col("Ongkir yang ditanggung platform")]),
        sumberPesanan:  String(r[col("Sumber pesanan")] || "TikTok Shop").trim() || "TikTok Shop",
      });
    } else if (jenis === "Pembayaran GMV untuk Iklan TikTok") {
      adPayments.push(Math.abs(n(r[col("Jumlah penyelesaian pembayaran")])));
    }
  }

  // ── 3. Use Laporan as authoritative, per-order as fallback ──
  const hasLaporan = laporan["Subtotal sebelum diskon"] > 0;

  const gmvBruto       = hasLaporan ? laporan["Subtotal sebelum diskon"]         : orders.reduce((t,o) => t + o.subtotalBruto, 0);
  const diskonPenjual  = hasLaporan ? laporan["Diskon penjual"]                  : orders.reduce((t,o) => t + o.diskonPenjual, 0);
  const netGmv         = hasLaporan ? laporan["Subtotal setelah diskon penjual"] : gmvBruto + diskonPenjual;
  const totalBiayaRaw  = hasLaporan ? Math.abs(laporan["Total Biaya"])           : Math.abs(orders.reduce((t,o) => t + o.totalBiaya, 0));
  const komisiPlatform = hasLaporan ? Math.abs(laporan["Biaya komisi platform"]) : Math.abs(orders.reduce((t,o) => t + o.komisiPlatform, 0));
  const komisiAfiliasi = hasLaporan ? Math.abs(laporan["Komisi Afiliasi"] || 0)  : Math.abs(orders.reduce((t,o) => t + o.komisiAfiliasi, 0));
  const komisiDinamis  = hasLaporan ? Math.abs(laporan["Komisi dinamis"]  || 0)  : Math.abs(orders.reduce((t,o) => t + o.komisiDinamis, 0));
  const biayaProses    = hasLaporan ? Math.abs(laporan["Biaya pemrosesan pesanan"] || 0) : Math.abs(orders.reduce((t,o) => t + o.biayaProses, 0));
  const totalDilepas   = hasLaporan ? laporan["Jumlah penyelesaian pembayaran"]  : orders.reduce((t,o) => t + o.settlement, 0);
  const adSpend        = adPayments.reduce((t,v) => t + v, 0);

  // Refund
  const refundValue    = hasLaporan
    ? Math.abs(laporan["Subtotal pengembalian dana sebelum diskon penjual"] || 0)
    : orders.filter(o => o.subtotalRefund < 0).reduce((t,o) => t + Math.abs(o.subtotalRefund), 0);
  const refundOrders   = orders.filter(o => o.subtotalRefund < 0);

  const summary = {
    channel: "tiktok",
    gmvBruto,
    diskonPenjual: -Math.abs(diskonPenjual),
    netGmv,
    totalBiaya:     -totalBiayaRaw,
    komisiPlatform: -komisiPlatform,
    komisiAfiliasi: -komisiAfiliasi,
    komisiDinamis:  -komisiDinamis,
    biayaProses:    -biayaProses,
    totalDilepas, adSpend,
    refundOrderCount: refundOrders.length,
    refundValue,
    totalOrderCount: orders.length,
    refundRateOrder: orders.length ? (refundOrders.length / orders.length) * 100 : 0,
    refundRateValue: gmvBruto ? (refundValue / gmvBruto) * 100 : 0,
    totalFee:       -totalBiayaRaw,
    feeRateNetGmv:  netGmv   ? totalBiayaRaw / netGmv   * 100 : 0,
    feeRateGross:   gmvBruto ? totalBiayaRaw / gmvBruto * 100 : 0,
    takeRate:       netGmv   ? (1 - totalDilepas / netGmv) * 100 : 0,
    hargaAsli: gmvBruto,
    diskonProduk: -Math.abs(diskonPenjual),
    refund: -refundValue,
    voucherSeller: 0,
    komisiAMS:    -komisiPlatform,
    biayaAdmin:   -(komisiAfiliasi + komisiDinamis),
    biayaLayanan: 0,
    netGmv,
    dataSource: hasLaporan ? "laporan" : "detail_pesanan",
  };

  return { channel: "tiktok", seller: "TikTok Shop", periodStart, periodEnd, summary, orders };
}



/* ============================================================================
   SHOPEE ORDER EXPORT PARSER  -  "orders" sheet
   Columns: No. Pesanan, Status Pesanan, SKU Induk, Nama Produk,
   Jumlah, Kota/Kabupaten, Provinsi, Waktu Pesanan Dibuat, dll
   ========================================================================== */
function parseOrderXlsx(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const sh = wb.Sheets["orders"] || wb.Sheets[wb.SheetNames[0]];
  if (!sh) throw new Error("Sheet tidak ditemukan.");

  const range = XLSX.utils.decode_range(sh["!ref"] || "A1:AX2000");
  const rows = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const row = [];
    for (let CC = range.s.c; CC <= range.e.c; CC++) {
      const cell = sh[XLSX.utils.encode_cell({ r: R, c: CC })];
      row.push(cell ? cell.v : "");
    }
    rows.push(row);
  }
  if (!rows.length) throw new Error("File kosong.");

  const hdrs = rows[0].map(c => String(c ?? "").trim());
  const col = name => hdrs.indexOf(name);

  const SELESAI = ["Selesai", "Completed"];
  const ACTIVE = str => str && !str.startsWith("Batal") && !str.startsWith("Cancel");

  const orders = [];
  let periodStart = "", periodEnd = "";

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[col("No. Pesanan")]) continue;
    const status = String(r[col("Status Pesanan")] ?? "").trim();
    if (!ACTIVE(status)) continue; // exclude cancelled

    const tgl = String(r[col("Waktu Pesanan Dibuat")] ?? "").slice(0, 10);
    if (tgl) {
      if (!periodStart || tgl < periodStart) periodStart = tgl;
      if (!periodEnd   || tgl > periodEnd)   periodEnd   = tgl;
    }
    const qty = parseFloat(String(r[col("Jumlah")] ?? "").replace(/[^0-9.]/g,"")) || 1;
    orders.push({
      noPesanan: String(r[col("No. Pesanan")] || ""),
      status,
      selesai: SELESAI.includes(status),
      skuInduk: String(r[col("SKU Induk")] || ""),
      namaProduk: String(r[col("Nama Produk")] || ""),
      jumlah: qty,
      kota: String(r[col("Kota/Kabupaten")] || "").trim(),
      provinsi: String(r[col("Provinsi")] || "").trim(),
      tglDibuat: tgl,
    });
  }

  return { periodStart, periodEnd, orders };
}


/* ============================================================================
   COGS TEMPLATE PARSER + GENERATOR
   Template columns: SKU, Nama Internal, HPP Produksi, HPP Packaging,
                     HPP Lainnya, Total HPP (auto), Harga Jual
   ========================================================================== */
function generateCOGSTemplate() {
  const rows = [
    ["SELLIO — Template COGS / HPP", "", "", "", "", "", ""],
    ["Isi kolom biru (HPP Produksi, HPP Packaging, HPP Lainnya, Harga Jual). Total HPP & Gross Margin% otomatis.", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["Nama Produk", "HPP Produksi (Rp)", "HPP Packaging (Rp)", "HPP Lainnya (Rp)", "Total HPP (auto)", "Harga Jual (Rp)", "Gross Margin %"],
    ["Contoh: Nama Produk Kamu", 85000, 6200, 0, { f: "B5+C5+D5" }, 189000, { f: 'IF(F5>0,(F5-E5)/F5,"-")' }],
  ];
  for (let i = 0; i < 50; i++) {
    const r = i + 6;
    rows.push(["", "", "", "", { f: `IF(B${r}+C${r}+D${r}>0,B${r}+C${r}+D${r},"-")` }, "", { f: `IF(F${r}>0,(F${r}-E${r})/F${r},"-")` }]);
  }
  rows.push(["Biru = input kamu | Abu = formula otomatis | HPP Packaging default Rp 6.200 (sesuaikan kalau beda)", "", "", "", "", "", ""]);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 42 }, { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "COGS");
  XLSX.writeFile(wb, "sellio_cogs_template.xlsx");
}

function parseCOGSXlsx(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const sh = wb.Sheets["COGS"] || wb.Sheets[wb.SheetNames[0]];
  if (!sh) throw new Error("Sheet COGS tidak ditemukan.");
  const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
  if (rows.length < 2) throw new Error("File kosong.");
  // Find actual header row — require multiple separate cells with column header keywords
  let headerIdx = 0;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const cells = rows[i].map(c => String(c ?? "").trim().toLowerCase());
    const nonEmpty = cells.filter(c => c.length > 0);
    if (nonEmpty.length < 3) continue; // header must have 3+ separate columns
    // Each cell individually should be a short label, not a sentence
    const isHeaderRow = nonEmpty.every(c => c.length < 40) &&
      (cells.some(c => c.includes("produksi")) || cells.some(c => c.includes("packaging")));
    if (isHeaderRow) { headerIdx = i; break; }
  }
  const hdrs = rows[headerIdx].map(c => String(c).trim().toLowerCase());
  const col = name => hdrs.findIndex(h => h.includes(name.toLowerCase()));
  // Data starts after header row
  const dataStart = headerIdx + 1;
  const colSku     = col("sku");
  const colNama    = col("nama");
  const colProd    = col("produksi");
  const colPack    = col("packaging");
  const colLain    = col("lainnya");
  const colTotal   = col("total");
  const colHarga   = col("harga");
  const n = v => { const x = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.-]/g,"")); return isNaN(x) ? 0 : x; };
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    // SKU bisa kosong — fallback ke nama produk sebagai identifier
    const sku  = colSku  >= 0 ? String(r[colSku]  ?? "").trim() : "";
    const nama = colNama >= 0 ? String(r[colNama] ?? "").trim() : "";
    const identifier = sku || nama;
    if (!identifier) continue; // skip baris kosong
    const hpp_prod = n(r[colProd]);
    const hpp_pack = n(r[colPack]);
    const hpp_lain = colLain >= 0 ? n(r[colLain]) : 0;
    const total = hpp_prod + hpp_pack + hpp_lain || (colTotal >= 0 ? n(r[colTotal]) : 0);
    if (!total && !n(r[colHarga])) continue; // skip baris yang benar-benar kosong
    const harga = colHarga >= 0 ? n(r[colHarga]) : 0;
    items.push({ sku: identifier, nama: nama || sku, hpp_prod, hpp_pack, hpp_lain, total, harga });
  }
  return items;
}


/* ============================================================================
   TIKTOK CAMPAIGN PER PRODUK PARSER
   File: creative_data_for_product_campaigns_...xlsx
   Columns: Nama kampanye, ID Campaign, ID produk, Jenis materi iklan,
            Biaya, Pesanan SKU, Biaya per pesanan, Pendapatan kotor, ROI
   ========================================================================== */
function parseTikTokCampaignXlsx(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const sh = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
  if (!rows.length) throw new Error("File kosong.");

  const hdrs = rows[0].map(c => String(c ?? "").trim());
  const col  = name => hdrs.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));

  const colCampaign   = col("nama kampanye");
  const colCampaignId = col("id campaign");
  const colProductId  = col("id produk");
  const colSpend      = col("biaya");
  const colOrders     = col("pesanan sku");
  const colCpa        = col("biaya per pesanan");
  const colGmv        = col("pendapatan kotor");
  const colRoi        = col("roi");
  const colStatus     = col("status");
  const colImpr       = col("impresi");
  const colClicks     = col("jumlah klik");
  const colCtr        = col("tingkat klik");
  const colCvr        = col("rasio konversi");

  const n  = v => { const x = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.-]/g,"")); return isNaN(x) ? 0 : x; };
  const pct = v => { const s = String(v ?? "").replace("%","").trim(); const x = parseFloat(s); return isNaN(x) ? 0 : x; };

  const campaignMap = {};
  let totalSpend = 0, totalOrders = 0, totalGmv = 0;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[colCampaign]) continue;
    const camp   = String(r[colCampaign]).trim();
    const campId = String(r[colCampaignId] || camp).trim();
    const prodId = String(r[colProductId] || "").trim();
    const spend  = n(r[colSpend]);
    const orders = n(r[colOrders]);
    const gmv    = n(r[colGmv]);
    const roi    = n(r[colRoi]);
    const cpa    = n(r[colCpa]);
    const status = String(r[colStatus] || "").trim();
    const impr   = colImpr   >= 0 ? n(r[colImpr])     : 0;
    const clicks = colClicks >= 0 ? n(r[colClicks])    : 0;
    const ctr    = colCtr    >= 0 ? pct(r[colCtr])     : 0;
    const cvr    = colCvr    >= 0 ? pct(r[colCvr])     : 0;

    if (!campaignMap[campId]) campaignMap[campId] = {
      name: camp, id: campId,
      spend: 0, orders: 0, gmv: 0, impr: 0, clicks: 0,
      roiSum: 0, roiCount: 0, ctrSum: 0, ctrCount: 0, cvrSum: 0, cvrCount: 0,
      products: []
    };
    const c = campaignMap[campId];
    c.spend  += spend;  c.orders += orders; c.gmv += gmv;
    c.impr   += impr;   c.clicks += clicks;
    if (roi  > 0) { c.roiSum += roi; c.roiCount++; }
    if (ctr  > 0) { c.ctrSum += ctr; c.ctrCount++; }
    if (cvr  > 0) { c.cvrSum += cvr; c.cvrCount++; }
    c.products.push({ prodId, spend, orders, gmv, roi, cpa, status, impr, clicks, ctr, cvr });

    totalSpend += spend; totalOrders += orders; totalGmv += gmv;
  }

  const campaigns = Object.values(campaignMap).map(c => ({
    ...c,
    roi: c.spend > 0 ? c.gmv / c.spend : 0,
    ctr: c.ctrCount > 0 ? c.ctrSum / c.ctrCount : (c.clicks > 0 && c.impr > 0 ? c.clicks / c.impr * 100 : 0),
    cvr: c.cvrCount > 0 ? c.cvrSum / c.cvrCount : 0,
  })).sort((a, b) => b.gmv - a.gmv);

  return {
    id: "tiktok_campaign|" + Date.now(),
    campaigns,
    summary: {
      totalSpend, totalOrders, totalGmv,
      blendedRoi: totalSpend > 0 ? totalGmv / totalSpend : 0,
      campaignCount: campaigns.length,
      totalImpr: campaigns.reduce((t,c) => t + c.impr, 0),
    },
    importedAt: Date.now(),
  };
}

   /* ============================================================================
   META ADS CSV PARSER  -  Facebook/Instagram Ads Manager export
   Columns: Campaign name, Impressions, Reach, Results, Result type,
   Amount spent (IDR), CTR (all), CPC (all), Cost per result,
   Results ROAS, Adds to cart, Reporting starts/ends
   ========================================================================== */
function parseMetaAdsCsv(text) {
  const lines = text.trim().split("\n");
  if (!lines.length) throw new Error("File kosong.");
  const headers = lines[0].split(",").map(h => h.trim().replace(/"/g,""));
  const col = name => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));

  const colCampaign  = col("campaign name");
  const colImpr      = col("impressions");
  const colReach     = col("reach");
  const colResults   = col("results");
  const colResType   = col("result type");
  const colSpend     = col("amount spent");
  const colCtr       = col("ctr");
  const colCpc       = col("cpc");
  const colCpr       = col("cost per result");
  const colRoas      = col("results roas");
  const colAtc       = col("adds to cart");
  const colStart     = col("reporting starts");
  const colEnd       = col("reporting ends");

  const n = v => { const x = parseFloat(String(v ?? "").replace(/[^0-9.-]/g,"")); return isNaN(x) ? 0 : x; };

  const campaigns = [];
  let periodStart = "", periodEnd = "";

  for (let i = 1; i < lines.length; i++) {
    const raw = lines[i];
    if (!raw.trim()) continue;
    // CSV parse — handle quoted fields
    const r = [];
    let cur = "", inQ = false;
    for (let j = 0; j < raw.length; j++) {
      if (raw[j] === '"') { inQ = !inQ; continue; }
      if (raw[j] === "," && !inQ) { r.push(cur.trim()); cur = ""; continue; }
      cur += raw[j];
    }
    r.push(cur.trim());

    const name = r[colCampaign] || "";
    if (!name || name.toLowerCase() === "campaign name") continue;

    const start = colStart >= 0 ? String(r[colStart] || "").trim() : "";
    const end   = colEnd   >= 0 ? String(r[colEnd]   || "").trim() : "";
    if (start && (!periodStart || start < periodStart)) periodStart = start;
    if (end   && (!periodEnd   || end   > periodEnd))   periodEnd   = end;

    campaigns.push({
      name,
      resultType:  colResType >= 0 ? String(r[colResType] || "").trim() : "",
      impressions: n(r[colImpr]),
      reach:       n(r[colReach]),
      results:     n(r[colResults]),
      spend:       n(r[colSpend]),
      ctr:         n(r[colCtr]),
      cpc:         n(r[colCpc]),
      cpr:         n(r[colCpr]),
      roas:        n(r[colRoas]),
      atc:         colAtc >= 0 ? n(r[colAtc]) : 0,
    });
  }

  if (!campaigns.length) throw new Error("Tidak ada data campaign ditemukan.");

  // Aggregate totals
  const totalSpend  = campaigns.reduce((t,c) => t+c.spend, 0);
  const totalImpr   = campaigns.reduce((t,c) => t+c.impressions, 0);
  const totalResult = campaigns.reduce((t,c) => t+c.results, 0);
  const blendedRoas = totalSpend > 0
    ? campaigns.reduce((t,c) => t + c.roas * c.spend, 0) / totalSpend : 0;
  const blendedCtr  = totalImpr  > 0
    ? campaigns.reduce((t,c) => t + c.ctr  * c.impressions, 0) / totalImpr : 0;
  const blendedCpr  = totalResult > 0 ? totalSpend / totalResult : 0;

  return {
    periodStart, periodEnd,
    campaigns,
    summary: { totalSpend, totalImpr, totalResult, blendedRoas, blendedCtr, blendedCpr }
  };
}


/* ============================================================================
   META ADS DIAGNOSIS ENGINE
   Classifies campaigns by objective, applies appropriate benchmarks per type
   ========================================================================== */
function classifyMetaObjective(resultType) {
  const rt = (resultType || "").toLowerCase();
  if (rt.includes("purchase") || rt.includes("conversion") || rt.includes("sale"))
    return "conversion";
  if (rt.includes("link click") || rt.includes("landing") || rt.includes("traffic") || rt.includes("profile visit"))
    return "consideration";
  if (rt.includes("engagement") || rt.includes("video view") || rt.includes("post"))
    return "engagement";
  if (rt.includes("reach") || rt.includes("brand awareness") || rt.includes("impression"))
    return "awareness";
  return "other";
}

const META_OBJ = {
  conversion:   { label: "Konversi",     color: "#3fb950", icon: "🛒", desc: "Purchase / CPAS" },
  consideration:{ label: "Consideration",color: "#58a6ff", icon: "🔗", desc: "Traffic / Klik" },
  engagement:   { label: "Engagement",   color: "#d29922", icon: "💬", desc: "Post Boost" },
  awareness:    { label: "Awareness",    color: "#8b949e", icon: "📢", desc: "Reach / Brand" },
  other:        { label: "Lainnya",      color: "#8b949e", icon: "—",  desc: "" },
};

// Benchmarks per objective — Indonesian fashion brand context
const META_BENCH = {
  conversion:    { roasMin: 3.0, roasGood: 6.0, cprMax: 50000, ctrMin: 0.8 },
  consideration: { ctrMin: 1.5, ctrGood: 3.0, cprMax: 500, roasMin: 0 },
  engagement:    { cprMax: 300, ctrMin: 0.5 },
  awareness:     { cprMax: 2000, ctrMin: 0.1 },
  other:         { cprMax: 10000 },
};

function diagnoseMetaCampaign(c) {
  const obj = classifyMetaObjective(c.resultType);
  const bench = META_BENCH[obj] || META_BENCH.other;
  const items = [];

  if (obj === "conversion") {
    if (c.roas >= (bench.roasGood || 6)) {
      items.push({ level: "good", tag: "SCALE", msg: `ROAS ${c.roas.toFixed(2)}x jauh di atas threshold ${bench.roasGood}x. Campaign ini proven — naikkan budget bertahap 20–30% tiap 3–4 hari, pantau frekuensi audience agar tidak fatigue.` });
    } else if (c.roas >= (bench.roasMin || 3)) {
      items.push({ level: "watch", tag: "ON TARGET", msg: `ROAS ${c.roas.toFixed(2)}x di atas minimum ${bench.roasMin}x tapi belum optimal. Coba refresh creative atau narrow audience ke lookalike buyer yang lebih ketat.` });
    } else if (c.roas > 0 && c.roas < bench.roasMin) {
      items.push({ level: "bad", tag: "ROAS RENDAH", msg: `ROAS ${c.roas.toFixed(2)}x di bawah minimum ${bench.roasMin}x untuk campaign konversi. Audit: landing page, produk yang dipromote, audience match. Pertimbangkan pause dan redesign creative.` });
    }
    if (c.cpr > bench.cprMax) {
      items.push({ level: "watch", tag: "CPR TINGGI", msg: `Cost per purchase Rp ${c.cpr.toLocaleString("id-ID")} — di atas threshold Rp ${bench.cprMax.toLocaleString("id-ID")}. Cek apakah audience terlalu broad atau produk yang dipromote harganya tidak kompetitif.` });
    }
    if (c.ctr < bench.ctrMin) {
      items.push({ level: "watch", tag: "CTR RENDAH", msg: `CTR ${c.ctr.toFixed(2)}% di bawah benchmark ${bench.ctrMin}% untuk conversion campaign. Creative tidak cukup menarik untuk mendorong klik — coba video vs static, atau ganti hook teks pertama 3 detik.` });
    }
  }

  if (obj === "consideration") {
    if (c.ctr >= (bench.ctrGood || 3)) {
      items.push({ level: "good", tag: "CTR BAGUS", msg: `CTR ${c.ctr.toFixed(2)}% di atas benchmark ${bench.ctrGood}% — creative resonan dengan audience. Pastikan landing page load time < 3 detik dan ada CTA yang jelas untuk maximize konversi post-click.` });
    } else if (c.ctr >= bench.ctrMin) {
      items.push({ level: "watch", tag: "CTR MODERAT", msg: `CTR ${c.ctr.toFixed(2)}% — acceptable tapi masih ada ruang. A/B test thumbnail atau headline. Untuk traffic campaign, CTR > ${bench.ctrGood}% adalah target yang realistis untuk fashion.` });
    } else {
      items.push({ level: "bad", tag: "CTR RENDAH", msg: `CTR ${c.ctr.toFixed(2)}% di bawah benchmark ${bench.ctrMin}% untuk traffic campaign. Creative tidak cukup menarik. Ganti visual utama — fashion sangat visual-driven, gambar produk on-model vs flat lay bisa beda CTR signifikan.` });
    }
    if (c.cpr <= bench.cprMax) {
      items.push({ level: "good", tag: "COST EFISIEN", msg: `Cost per klik Rp ${c.cpr.toLocaleString("id-ID")} — efisien untuk campaign traffic fashion Indonesia. Pertahankan audience dan creative yang sekarang.` });
    } else {
      items.push({ level: "watch", tag: "COST TINGGI", msg: `Cost per klik Rp ${c.cpr.toLocaleString("id-ID")} melebihi benchmark Rp ${bench.cprMax}. Cek apakah audience overlap dengan campaign lain (auction overlap bikin CPM naik) atau coba broad audience dengan creative kuat.` });
    }
    if (c.roas === 0) {
      items.push({ level: "neutral", tag: "FUNNEL ATAS", msg: "Traffic campaign tidak diukur dari ROAS langsung — tujuannya build awareness dan isi funnel retargeting. Ukur efektivitasnya dari kualitas traffic (bounce rate, time on site) dan apakah audience ini convert di campaign retargeting berikutnya." });
    }
  }

  if (obj === "engagement") {
    if (c.cpr <= bench.cprMax) {
      items.push({ level: "good", tag: "COST EFISIEN", msg: `Cost per engagement Rp ${c.cpr.toLocaleString("id-ID")} — efisien untuk boost post. Engagement tinggi juga bantu social proof organik (likes/comments visible ke calon pembeli baru).` });
    } else {
      items.push({ level: "watch", tag: "COST TINGGI", msg: `Cost per engagement Rp ${c.cpr.toLocaleString("id-ID")} di atas Rp ${bench.cprMax}. Boost post lebih efektif kalau post organik-nya sudah punya engagement awal — consider boost hanya post yang sudah proven secara organik.` });
    }
    items.push({ level: "neutral", tag: "TUJUAN", msg: "Engagement campaign tidak direct-response — fungsinya social proof dan brand recall. Jangan ukur dari ROAS. Efektivitasnya terlihat dari peningkatan profile visits dan organic reach setelah campaign." });
  }

  if (obj === "awareness") {
    if (c.impressions > 0) {
      const cpm = c.spend / c.impressions * 1000;
      if (cpm < 15000) {
        items.push({ level: "good", tag: "CPM EFISIEN", msg: `CPM Rp ${cpm.toLocaleString("id-ID", {maximumFractionDigits:0})} — efisien untuk awareness campaign. Reach ${c.reach.toLocaleString("id-ID")} unique audience dengan cost yang wajar.` });
      } else {
        items.push({ level: "watch", tag: "CPM TINGGI", msg: `CPM Rp ${cpm.toLocaleString("id-ID", {maximumFractionDigits:0})} — untuk awareness campaign, coba perluas audience atau gunakan Advantage+ audience untuk biarkan Meta optimasi sendiri.` });
      }
    }
    items.push({ level: "neutral", tag: "KONTEKS", msg: "Awareness campaign diukur dari CPM dan reach, bukan ROAS atau klik. Fungsinya top-of-funnel: pastikan orang-orang yang belum kenal brand kamu mulai terekspos. Efektivitas jangka panjangnya terlihat dari peningkatan branded search dan organic traffic." });
  }

  if (!items.length) {
    items.push({ level: "neutral", tag: "INFO", msg: "Tidak cukup data untuk analisa mendalam. Pastikan campaign sudah berjalan minimal 7 hari dengan budget yang cukup untuk keluar dari fase learning Meta." });
  }

  return { obj, items };
}


/* ============================================================================
   TIKTOK GMV MAX PARSER  -  Campaign Overview daily export XLSX
   Columns: Per Hari, Biaya, Pesanan SKU, Biaya per pesanan,
            Penghasilan bruto, ROI, Mata Uang
   ========================================================================== */
function parseTikTokAdsXlsx(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const sh = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
  if (!rows.length) throw new Error("File kosong.");

  const hdrs = rows[0].map(c => String(c ?? "").trim().toLowerCase());
  const col  = name => hdrs.findIndex(h => h.includes(name.toLowerCase()));
  const colDate    = col("per hari");
  const colSpend   = col("biaya");
  const colOrders  = col("pesanan sku");
  const colCpa     = col("biaya per pesanan");
  const colGmv     = col("penghasilan bruto");
  const colRoi     = col("roi");

  const n = v => { const x = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.-]/g,"")); return isNaN(x) ? 0 : x; };

  const daily = [];
  let periodStart = "", periodEnd = "";

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const rawDate = r[colDate];
    if (!rawDate) continue;
    const tgl = String(rawDate).slice(0, 10).replace(/ .*/,"");
    if (tgl && tgl !== "-") {
      if (!periodStart || tgl < periodStart) periodStart = tgl;
      if (!periodEnd   || tgl > periodEnd)   periodEnd   = tgl;
    }
    daily.push({
      tgl,
      spend:  n(r[colSpend]),
      orders: n(r[colOrders]),
      cpa:    n(r[colCpa]),
      gmv:    n(r[colGmv]),
      roi:    n(r[colRoi]),
    });
  }

  if (!daily.length) throw new Error("Tidak ada data harian.");

  // Aggregate
  const totalSpend  = daily.reduce((t,d) => t+d.spend, 0);
  const totalOrders = daily.reduce((t,d) => t+d.orders, 0);
  const totalGmv    = daily.reduce((t,d) => t+d.gmv, 0);
  const blendedRoi  = totalSpend > 0 ? totalGmv/totalSpend : 0;
  const avgCpa      = totalOrders > 0 ? totalSpend/totalOrders : 0;
  const avgRoi      = daily.filter(d=>d.roi>0).reduce((t,d)=>t+d.roi,0) / Math.max(daily.filter(d=>d.roi>0).length, 1);

  // Weekly aggregates
  const weeks = {};
  daily.forEach(d => {
    if (!d.tgl || d.tgl === "-") return;
    const dt = new Date(d.tgl);
    const weekNum = Math.ceil(dt.getDate()/7);
    const wk = `W${weekNum}`;
    if (!weeks[wk]) weeks[wk] = { spend:0, gmv:0, orders:0, days:0 };
    weeks[wk].spend += d.spend; weeks[wk].gmv += d.gmv;
    weeks[wk].orders += d.orders; weeks[wk].days++;
  });

  return {
    periodStart, periodEnd, daily,
    summary: { totalSpend, totalOrders, totalGmv, blendedRoi, avgCpa, avgRoi, weeks }
  };
}

/* ============================================================================
   APP
   ========================================================================== */
export default function Sellio() {
  const [snapshots, setSnapshots] = useState([]);
  const [products, setProducts] = useState([]);
  const [cogsMap, setCogsMap] = useState({});
  const [cogsItems, setCogsItems] = useState([]); // [{sku, nama, hpp_prod, hpp_pack, hpp_lain, total, harga}]
  const [metaSnaps, setMetaSnaps] = useState([]);   // Meta Ads snapshots
  const [tiktokAdsSnaps, setTiktokAdsSnaps] = useState([]); // TikTok Ads (GMV Max) snapshots
  const [tiktokCampaignSnap, setTiktokCampaignSnap] = useState(null); // TikTok per-campaign/product
  const [pnlInputs, setPnlInputs] = useState({}); // P&L manual inputs
  const [stockMap, setStockMap] = useState({}); // code -> units
  const [thresholds, setThresholds] = useState(DEFAULT_TH);
  const [tab, setTab] = useState("overview");
  const [activeId, setActiveId] = useState(null);
  const [toast, setToast] = useState(null);
  const [confirmDlg, setConfirmDlg] = useState(null); // { msg, onOk }
  const [importMenuOpen, setImportMenuOpen] = useState(false);
  const [panduanOpen, setPanduanOpen] = useState(false);
  const [incomeSnaps, setIncomeSnaps] = useState([]);       // Shopee income
  const [tiktokSnaps, setTiktokSnaps] = useState([]);       // TikTok income
  const [orderSnaps, setOrderSnaps] = useState([]);           // Shopee Order export
  const adFileRef = useRef();
  const prodFileRef = useRef();
  const incomeFileRef = useRef();
  const tiktokFileRef = useRef();
  const orderFileRef = useRef();
  const cogsFileRef = useRef();
  const metaFileRef = useRef();
  const tiktokAdsFileRef = useRef();
  const tiktokCampaignFileRef = useRef();

  useEffect(() => {
    (async () => {
      const snaps = await idbAll("snapshots");
      snaps.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setSnapshots(snaps);
      if (snaps[0]) setActiveId(snaps[0].id);
      const prods = await idbAll("products");
      prods.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setProducts(prods);
      const cogs = await idbAll("cogs");
      const m = {};
      cogs.forEach((c) => (m[c.key] = c));
      setCogsMap(m);
      const stocks = await idbAll("stock");
      const sm = {};
      stocks.forEach((s) => (sm[s.code] = s.units));
      setStockMap(sm);
      const st = await idbAll("settings");
      const thRec = st.find((x) => x.key === "thresholds");
      if (thRec) setThresholds({ ...DEFAULT_TH, ...thRec.value });
      const incomes = await idbAll("income");
      incomes.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setIncomeSnaps(incomes);
      const tiktoks = await idbAll("income_tiktok");
      tiktoks.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setTiktokSnaps(tiktoks);
      const orders = await idbAll("orders");
      orders.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setOrderSnaps(orders);
      const ci = await idbAll("cogs_items");
      setCogsItems(ci);
      const metas = await idbAll("meta_ads");
      metas.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setMetaSnaps(metas);
      const tiktokAds = await idbAll("tiktok_ads");
      tiktokAds.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setTiktokAdsSnaps(tiktokAds);
      const ttCamps = await idbAll("tiktok_campaign");
      if (ttCamps.length > 0) setTiktokCampaignSnap(ttCamps[ttCamps.length - 1]);
      const pnlRecs = await idbAll("pnl_inputs");
      const pnlMap = {};
      pnlRecs.forEach(r => { pnlMap[r.key] = r.value; });
      setPnlInputs(pnlMap);
    })();
  }, []);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2600); };

  function confirm(msg, onOk) { setConfirmDlg({ msg, onOk }); }

  async function deleteSnapshot(id) {
    await idbDel("snapshots", id);
    const snaps = await idbAll("snapshots");
    snaps.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
    setSnapshots(snaps);
    if (activeId === id) setActiveId(snaps[0]?.id || null);
    flash("Snapshot dihapus.");
  }

  async function deleteAllData() {
    // confirm handled by caller via confirmDlg
    const unused = false; if (unused) return;
    const db = await openDB();
    await Promise.all(["snapshots","products","cogs","stock","settings","income","income_tiktok","orders","cogs_items","meta_ads","tiktok_ads","pnl_inputs"].map(store =>
      new Promise((res, rej) => {
        const tx = db.transaction(store, "readwrite");
        tx.objectStore(store).clear();
        tx.oncomplete = res; tx.onerror = rej;
      })
    ));
    setSnapshots([]); setProducts([]); setCogsMap({}); setStockMap({}); setIncomeSnaps([]); setTiktokSnaps([]); setOrderSnaps([]); setCogsItems([]); setMetaSnaps([]); setTiktokAdsSnaps([]); setTiktokCampaignSnap(null); setPnlInputs({});
    setActiveId(null); setThresholds(DEFAULT_TH);
    flash("Semua data direset.");
  }

  async function handleAdFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    const text = await f.text();
    try {
      const parsed = parseShopee(text);
      const endDate = dmy(parsed.periodEnd);
      const id = "shopee|" + (parsed.periodEnd || f.name);
      const snap = {
        id, channel: "shopee", store: parsed.store,
        periodStart: parsed.periodStart, periodEnd: parsed.periodEnd,
        endTs: endDate ? endDate.getTime() : Date.now(),
        ads: parsed.ads, importedAt: Date.now(), fileName: f.name,
      };
      await idbPut("snapshots", snap);
      const snaps = await idbAll("snapshots");
      snaps.sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : -1));
      setSnapshots(snaps);
      setActiveId(id);
      flash(`Iklan: ${parsed.ads.length} baris · ${parsed.periodStart} → ${parsed.periodEnd}`);
    } catch (err) {
      flash("Gagal baca file iklan: " + err.message);
    }
    e.target.value = "";
  }

  async function handleProdFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const parsed = parseProductXlsx(buf);
      const { start, end } = periodFromFilename(f.name);
      const endDate = dmy(end);
      const id = "shopee-prod|" + (end || f.name);
      const snap = {
        id, periodStart: start, periodEnd: end,
        endTs: endDate ? endDate.getTime() : Date.now(),
        products: parsed.products, importedAt: Date.now(), fileName: f.name,
      };
      // Replace: clear old product data before inserting new one
      const db = await openDB();
      await new Promise((res, rej) => {
        const tx = db.transaction("products", "readwrite");
        tx.objectStore("products").clear();
        tx.oncomplete = res; tx.onerror = rej;
      });
      await idbPut("products", snap);
      setProducts([snap]);
      setTab("product");
      flash(`Produk: ${parsed.products.length} SKU · ${start || "?"} → ${end || "?"}  ·  data lama diganti.`);
    } catch (err) {
      flash("Gagal baca file produk: " + err.message);
    }
    e.target.value = "";
  }

  async function handleTikTokAdsFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const parsed = parseTikTokAdsXlsx(buf);
      const id = "tiktokads|" + (parsed.periodEnd || f.name);
      const snap = { id, periodStart: parsed.periodStart, periodEnd: parsed.periodEnd, daily: parsed.daily, summary: parsed.summary, importedAt: Date.now(), fileName: f.name };
      const db = await openDB();
      await new Promise((res,rej) => { const tx=db.transaction("tiktok_ads","readwrite"); tx.objectStore("tiktok_ads").clear(); tx.oncomplete=res; tx.onerror=rej; });
      await idbPut("tiktok_ads", snap);
      setTiktokAdsSnaps([snap]);
      setTab("performance");
      flash(`TikTok Ads: ${parsed.daily.length} hari · ${parsed.periodStart} → ${parsed.periodEnd}`);
    } catch(err) {
      flash("Gagal baca file TikTok Ads: " + err.message);
    }
    e.target.value = "";
  }

  async function handleTikTokCampaignFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const parsed = parseTikTokCampaignXlsx(buf);
      await new Promise((res,rej) => { openDB().then(db => { const tx=db.transaction("tiktok_campaign","readwrite"); tx.objectStore("tiktok_campaign").clear(); tx.oncomplete=res; tx.onerror=rej; }); });
      await idbPut("tiktok_campaign", parsed);
      setTiktokCampaignSnap(parsed);
      setTab("performance");
      flash(`TikTok Campaign: ${parsed.summary.campaignCount} kampanye · ${parsed.summary.totalOrders} orders`);
    } catch(err) {
      flash("Gagal baca file TikTok Campaign: " + err.message);
    }
    e.target.value = "";
  }

  async function handleMetaFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const text = await f.text();
      const parsed = parseMetaAdsCsv(text);
      const id = "meta|" + (parsed.periodEnd || f.name);
      const snap = { id, periodStart: parsed.periodStart, periodEnd: parsed.periodEnd, campaigns: parsed.campaigns, summary: parsed.summary, importedAt: Date.now(), fileName: f.name };
      const db = await openDB();
      await new Promise((res, rej) => { const tx = db.transaction("meta_ads","readwrite"); tx.objectStore("meta_ads").clear(); tx.oncomplete=res; tx.onerror=rej; });
      await idbPut("meta_ads", snap);
      setMetaSnaps([snap]);
      setTab("performance");
      flash(`Meta Ads: ${parsed.campaigns.length} campaign · ${parsed.periodStart} → ${parsed.periodEnd}`);
    } catch(err) {
      flash("Gagal baca file Meta: " + err.message);
    }
    e.target.value = "";
  }

  async function handleCogsFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const items = parseCOGSXlsx(buf);
      const db = await openDB();
      // Save to cogs_items (for Unit Economics / CM / PnL)
      await new Promise((res, rej) => { const tx = db.transaction("cogs_items","readwrite"); tx.objectStore("cogs_items").clear(); tx.oncomplete=res; tx.onerror=rej; });
      for (const item of items) await idbPut("cogs_items", item);
      setCogsItems(items);

      // Also save margin% to cogs store so marginFor() works in Performance/Strategy
      const newCogsMap = { ...cogsMap };
      for (const item of items) {
        if (!item.harga || !item.total) continue;
        const marginPct = Math.max(0, Math.min(100, (item.harga - item.total) / item.harga * 100));
        // Key by SKU (c: prefix) or name (n: prefix) — same as keyFor()
        const key = item.sku && item.sku !== "-" && item.sku !== item.nama
          ? "c:" + item.sku
          : "n:" + item.nama;
        const rec = { key, label: item.nama, marginPct };
        await idbPut("cogs", rec);
        newCogsMap[key] = rec;
      }
      setCogsMap(newCogsMap);
      flash(`COGS: ${items.length} SKU diimport · margin otomatis terhubung ke Performa Iklan`);
    } catch(err) {
      flash("Gagal baca file COGS: " + err.message);
    }
    e.target.value = "";
  }

  async function handleOrderFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const parsed = parseOrderXlsx(buf);
      const id = "orders|" + (parsed.periodEnd || f.name);
      const snap = { id, periodStart: parsed.periodStart, periodEnd: parsed.periodEnd, orders: parsed.orders, importedAt: Date.now(), fileName: f.name };
      const db = await openDB();
      await new Promise((res, rej) => { const tx = db.transaction("orders","readwrite"); tx.objectStore("orders").clear(); tx.oncomplete=res; tx.onerror=rej; });
      await idbPut("orders", snap);
      setOrderSnaps([snap]);
      setTab("area");
      flash(`Order: ${parsed.orders.length} pesanan (non-batal) · ${parsed.periodStart} → ${parsed.periodEnd}`);
    } catch(err) {
      flash("Gagal baca file order: " + err.message);
    }
    e.target.value = "";
  }

  async function handleTikTokFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const parsed = parseTikTokIncomeXlsx(buf);
      const id = "tiktok|" + (parsed.periodEnd || f.name);
      const snap = { id, channel: "tiktok", seller: parsed.seller, periodStart: parsed.periodStart, periodEnd: parsed.periodEnd, summary: parsed.summary, orders: parsed.orders, importedAt: Date.now(), fileName: f.name };
      const db = await openDB();
      await new Promise((res, rej) => { const tx = db.transaction("income_tiktok","readwrite"); tx.objectStore("income_tiktok").clear(); tx.oncomplete=res; tx.onerror=rej; });
      await idbPut("income_tiktok", snap);
      setTiktokSnaps([snap]);
      setTab("fee");
      flash(`TikTok: ${parsed.orders.length} pesanan · ${parsed.periodStart} → ${parsed.periodEnd}`);
    } catch(err) {
      flash("Gagal baca file TikTok: " + err.message);
    }
    e.target.value = "";
  }

  async function handleIncomeFile(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const parsed = parseIncomeXlsx(buf);
      const id = "income|" + (parsed.periodEnd || f.name);
      const snap = { id, seller: parsed.seller, periodStart: parsed.periodStart, periodEnd: parsed.periodEnd, summary: parsed.summary, orders: parsed.orders, importedAt: Date.now(), fileName: f.name };
      // Replace: clear old income data
      const db = await openDB();
      await new Promise((res, rej) => { const tx = db.transaction("income","readwrite"); tx.objectStore("income").clear(); tx.oncomplete=res; tx.onerror=rej; });
      await idbPut("income", snap);
      setIncomeSnaps([snap]);
      setTab("fee");
      flash(`Penghasilan: ${parsed.orders.length} pesanan · ${parsed.periodStart} → ${parsed.periodEnd}`);
    } catch(err) {
      flash("Gagal baca file penghasilan: " + err.message);
    }
    e.target.value = "";
  }

  // resolve COGS key for an ad: prefer product code, fallback to name
  const keyFor = (ad) => (ad.code && ad.code !== "-" ? "c:" + ad.code : "n:" + ad.name);
  const marginFor = (ad) => {
    const k = keyFor(ad);
    return cogsMap[k] ? cogsMap[k].marginPct : null;
  };

  const active = snapshots.find((s) => s.id === activeId);

  // pick the product snapshot whose period end is closest to active ad snapshot
  const activeProducts = useMemo(() => {
    if (!products.length) return null;
    if (!active) return products[0];
    let best = products[0], bestDiff = Infinity;
    for (const p of products) {
      const d = Math.abs((p.endTs || 0) - (active.endTs || 0));
      if (d < bestDiff) { bestDiff = d; best = p; }
    }
    return best;
  }, [products, active]);

  // Date sync check across all three file sources
  const syncStatus = useMemo(() => {
    const sources = [];
    if (active) sources.push({ name: "Iklan", period: active.periodEnd?.slice(0,7) });
    if (activeProducts) sources.push({ name: "Produk", period: activeProducts.periodEnd?.slice(0,7) });
    if (incomeSnaps[0]) sources.push({ name: "Shopee", period: incomeSnaps[0].periodEnd?.slice(0,7) });
    if (tiktokSnaps[0]) sources.push({ name: "TikTok", period: tiktokSnaps[0].periodEnd?.slice(0,7) });
    if (orderSnaps[0]) sources.push({ name: "Order", period: orderSnaps[0].periodEnd?.slice(0,7) });
    if (metaSnaps[0]) sources.push({ name: "Meta", period: metaSnaps[0].periodEnd?.slice(0,7) });
    if (tiktokAdsSnaps[0]) sources.push({ name: "TikTok Ads", period: tiktokAdsSnaps[0].periodEnd?.slice(0,7) });
    if (sources.length < 2) return null;
    const periods = [...new Set(sources.map(s => s.period).filter(Boolean))];
    if (periods.length === 1) return { ok: true, msg: `Semua file periode ${periods[0]} ✓` };
    return { ok: false, msg: `⚠ Periode beda: ${sources.map(s => s.name + " " + (s.period||"?")).join(" · ")} — hasil analisa lintas-channel perlu disamakan dulu` };
  }, [active, activeProducts, incomeSnaps, tiktokSnaps, orderSnaps, metaSnaps, tiktokAdsSnaps]);

  async function saveThresholds(next) {
    setThresholds(next);
    await idbPut("settings", { key: "thresholds", value: next });
  }

  const hasAds = !!active;
  const hasProd = !!activeProducts;

  return (
    <div style={S.root}>
      <style>{CSS}</style>

      {/* Top bar */}
      <header style={S.topbar}>
        <div style={S.brand}>
          <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9ArFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTmlyQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKeDHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAAJ8UlEQVR4nM2aSYxlVRmAv/+cc+8b6lVXFdXd2JMM3Si0RKARQQjRCCG6UDGYSBxYaExcYEAXYCKGRANo3LgxkYjERBcsHGLiAkEcFogtbWOYlLZJE+ZuhO6qel1vuOec38W9r/q+1/e9qmoB/ZObuu/cM/zzdEpijMoqoDo8RUTGfit/V1VEZGXOuHWj4+X1o/uWvwO41ZCvOuTNhiomrBWkLIEqTo07ZByXquZM2udUYbD3miXwdsF6JW3eIjzeNhgiQETWLfby/NG1qrryVK2rWrteFRtSoUnepmp81GZGPUR5v/LcAVFlIiatnYSDG+fORjeZpJtViK42d626Po5ZA6iUwGDSWg6ZRPR6YFQSo/uW40kZKr3Qerl0KjCMjCBCSbVOjK+Gi4QQdBx142CturqWecNzQvkLIuOd5EBS/wdxIIJEUEF1QFA1tytVaFTf36xIWbXXqERUIUYICsYYTMFxBTAhf9FJ6/V/JwERwRgAg8MCEH0AY4sZkVwSk2OtG/W/47LHSYgMIBaP0wgoEYuKYKSPEuipxYaERCMxERb7yoFn2+x7JuOR/a+wcz5w+00XIhqAANihsyaq0Go+f00xQSJCH0VBE0AwCgSLzRyNhoATwPL4c0vcfOcBDj1/GgtZl8V+ynWXOyJgENAUo4KWjqokoAqxKgmsRRo2WDTWCU5QwNDHmAzMFH1RnjuyzN+ePMKeXRt5fUnZ/3xCrT5Ps9mn11kmqfcGqCJqUc3lOAncJMQmETeesAwfI6I1rAjHfcJPfnGAX+7r8uxLNV54PnL3rXDBecJUChLAS8Crzy0aRaSPkjBQIUFyqVacOzESr9Uj5e5PwQhiDfWBAUbLa53AD37d5l+vbmZ2uo60FkmbkASDjynWdLFkJCTYWKhPTIgCRorfIqDrUKFJICKIFzKrRMlwQbDWQuyyJE2eeznj8acXeHDvUS46u8X1120lbW2g3qjhxOOCh+jJSPACiMFigVDicZFKaMH5CsQHeFeq0GpeKCYR0T41UnDCi8d6bJ5u8N17/8k9v3L4bsrrx5vUPtkgBZL+Mhpb+MKxa6EQogP1EFQ0f4DcA7mTwllV5jrkZMfl7ieBjzgxLPYid9zzIp+/7QmWYuTVFzyHezXcfAs7F3D1bq5WGAQzwaPnhJVj8KjOlwkYKuqr4kCZoFEpxBixxnLwNeXGOw/xwKOw810bsAjTpk7TpNRizj+LEEUJrgbWYSyIsYBgBDAGIxaMQcQiElZwkIp0YsDcsoYMGfFqNpAbq+GNXuDmuw7wxyebzGy0TAVLNIYFFlheVo7bJr12h6wnqAr9zgK9dorzSr+zRNQZvCZ02228rWFcRnc5w2cRU5yjud2OyYpOQGUgqyJIC+47Z7nvoVd4eL9hZn6K5V6XaLvETuDGT5/J1VcqNesIPuHsd6bUrPL9G89gYdmiCajW2XN+g6k05b7bdpAJBIEQlTNPFzSChIRgFJWI02GNOImAVQg8QQyCMblXePixDiHdCDFgokM0okHZc+48e84t862PhoyrPrAJ6BVsOC1/D8onPryJ0ZShR0BMoBYSFIPKiUBWlZ6fZANlGKZYV/Dq9CyKQ6WLxWKCxRhPiBkhWkSSIoIqxhpCyHJeaZ7zpzZl74EFvv69fTQ3b+Wq9yXccPUcm2ca2CBEByoZosmKKVdVaLCOtooCGhRBuHi3Rbsv58gl4JM+CFhTIyEhUUgxpNRwwZESqRGoSSBRRYzQ7ioHXmrwuydb3Hp34Ppv/INnXslwxpB4RzRCsJPTiHURYI1BjBBj5IaPbeeaSxMW/r1Aux3p9SJZzFPfaCLRRqLR3BDFoGoIwRFx2DT3/ClCktSZaSqnz8yy96k5vn3vs3QFRDwSHaInZ6Ojz4oNjLY7yuOQG3A+B3a0Ij/61jn89LeH+f2fjzBnU8RZVDPQSBQDFOFfDIlLAPBkPHO4x/Zph7OGHi2CD8TYptmaZt9TGS++1mPXZkuIkQS7ksqNi0/rKmhyNwq+45ixnpuv3cSXr53HeyG1go8ZEi34PHAhATXCIwcX+MPfF9j/dOCxR5e546YdbNtlyKSDxFnU91HbpdNrsHBcc7SG6uPx4MrUjbYzylQbY4rfimvmOT0odZJ8l76HtDFSgxjeWPR89a6D/PXQPI2pFtqBruT5jfV9hAwjjhADU60eM61cWkZkKJUebTxMzEarua+Dlfz4/sM8cdBSa2b43jJbWo4vfnY7Dzz0EnufVmaTKd7wba64UPjgRVsIzNFqbaVZ79LrHcMRMTHFaAtv+4ixZMdqXHDFImfMW+iDphA1YgqOjOJW2ZUYorDo1ax844R3//mDr3P/X2aotQy+I+zcnPG5zzh+86dX+dmD08y2pjm23Gax67nqku0Ys0wwLxOpk0kCajEEQtLGpC2Wlz1bG12+8qldJMYTRRBV8qg2bMij+FamEit/tdQUUEWL8WajzoZZpdlwBDvF1IYuKdBsWlozjlbL0nMpTVfHBUVDSuob1CSlLQtE28UyjaYtOkuejY0jfOdr27h8V50QAiQWGz2qssK1KtUemwvlufioyAo5iOStkJCH/xAjIUYUCGroxYTgBckM1jtUlJB2EDONWEcamkBC10f8G0f50Pun+OYX3sll754jCx3EpkQCRgeoTc6OK1VI82I0J36FqMFWipG8ayYi+ftgXBNqsYlLMjRt42sb6FphMUs42g0shkXS/lFMlrJjrs4Pb9nCx6/ZwpSF0FdMKkQ8JtZQ8YhEIB0iYrSRPN6NyqDYYCQ+DEQZMaJgQEzu90Po0+0f49jxJp0lQ78D01G5ZNsSl22ts/u9DXbv2MaV59WZmzOc9dHTwUe8V0gDogYbEwyRYEPutnW4uzWafLpxujX4fcL55C5WUbzJ6Gctlnp9ekuwod6mo5HZhmf39j7vObfB2dvfwUcubtBM4d7bL6We5H2KgnZ8iCgZVlLERFQNoi6vrYs8CDWFFMr4DOM4MZAN9Y8LuxARnPdsaXY466we52w9jYt3bmJWI7d96XxuSS0bG4IlAPmBqTN4nydyIiAmL2hEauSVmC0kPFDSWi57Wb0imHhLWV0jGA69cJykLmzalFCXPKBlxDwzVUG9ElURA1jFF99EZWSvydXfSQytwieEcNJoVXvxRCCjaMJGCAaNkah5SwWbEcQj6jAkGDUI5TK+GvG1jI2DsSo0qbefFT5aTBcxDqsOiRFCDZEUlVwd1ARQk9eHq7jDUwWJMepaL7jLAQ4ElYL7IogooqZoKheeQ8iDUfnAMTlN1Xnl76eUjY77v4Ucp9zABq8y8DAmJ26lOTWIHmMQ+G/vIyYScCp3ZOu5saw6Y70ErXih9bYYywee6l3vavutpfHsyvpflbJWwVt5ezkKq3mk/wACeozss2wqqAAAAABJRU5ErkJggg==" style={{ width: 44, height: 44, objectFit: "contain" }} alt="Sellio" />
          <div>
            <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
              <div style={S.brandName}>SELLIO</div>
              <div style={{ fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.accent, letterSpacing: 0.3 }}>for Seller</div>
            </div>
            <div style={S.brandSub}>satu dashboard untuk semua keputusan bisnis marketplace kamu</div>
          </div>
        </div>
        <div style={S.topActions}>
          {active && <span style={S.storeTag}>{active.store}</span>}
          <button
            style={{ fontFamily: mono, fontSize: 12, fontWeight: 600, padding: "8px 14px", borderRadius: 8, border: `1px solid ${C.line}`, background: C.panel2, color: C.muted, cursor: "pointer", letterSpacing: 0.3 }}
            onClick={() => setPanduanOpen(true)}>
            Panduan
          </button>
          {/* Import button + dropdown */}
          <div style={{ position: "relative" }}>
            <button style={S.importBtn} onClick={() => setImportMenuOpen(v => !v)}>
              + Import Data {importMenuOpen ? "▲" : "▼"}
            </button>
            {importMenuOpen && (
              <div style={S.importMenu} onMouseLeave={() => setImportMenuOpen(false)}>

                {/* Group: Data Iklan */}
                <div style={{ fontFamily: sans, fontSize: 10, fontWeight: 700, color: C.dim, letterSpacing: 1, textTransform: "uppercase", padding: "6px 10px 4px" }}>Data Iklan</div>
                {[
                  { label: "Shopee Ads", sub: "Shopee Ads Manager → Laporan → Export CSV", icon: "📊", action: () => adFileRef.current.click(), accent: C.accent },
                  { label: "TikTok GMV Max", sub: "TikTok Ads Manager → Campaign Overview → Export", icon: "⬛", action: () => tiktokAdsFileRef.current.click(), accent: "#010101" },
                  { label: "TikTok Per Campaign", sub: "TikTok Ads Manager → Campaign → Creative → Export", icon: "⬛", action: () => tiktokCampaignFileRef.current.click(), accent: "#fe2c55" },
                  { label: "Meta Ads", sub: "Meta Ads Manager → Export → CSV", icon: "🔵", action: () => metaFileRef.current.click(), accent: "#1877f2" },
                ].map((item, i) => (
                  <button key={i} style={S.importMenuItem} onClick={() => { item.action(); setImportMenuOpen(false); }}>
                    <span style={{ fontSize: 16, lineHeight: 1, minWidth: 20 }}>{item.icon}</span>
                    <div>
                      <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: item.accent }}>{item.label}</div>
                      <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>{item.sub}</div>
                    </div>
                  </button>
                ))}

                <div style={{ borderTop: `1px solid ${C.line}`, margin: "6px 0" }} />

                {/* Group: Data Keuangan */}
                <div style={{ fontFamily: sans, fontSize: 10, fontWeight: 700, color: C.dim, letterSpacing: 1, textTransform: "uppercase", padding: "4px 10px" }}>Data Keuangan</div>
                {[
                  { label: "Penghasilan Shopee", sub: "Shopee → Keuangan → Penghasilan Saya → Download", icon: "🟠", action: () => incomeFileRef.current.click(), accent: "#f96d00" },
                  { label: "Penghasilan TikTok", sub: "TikTok Shop → Keuangan → Pendapatan → Download", icon: "⚫", action: () => tiktokFileRef.current.click(), accent: C.muted },
                ].map((item, i) => (
                  <button key={i} style={S.importMenuItem} onClick={() => { item.action(); setImportMenuOpen(false); }}>
                    <span style={{ fontSize: 16, lineHeight: 1, minWidth: 20 }}>{item.icon}</span>
                    <div>
                      <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: item.accent }}>{item.label}</div>
                      <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>{item.sub}</div>
                    </div>
                  </button>
                ))}

                <div style={{ borderTop: `1px solid ${C.line}`, margin: "6px 0" }} />

                {/* Group: Data Produk & Pesanan */}
                <div style={{ fontFamily: sans, fontSize: 10, fontWeight: 700, color: C.dim, letterSpacing: 1, textTransform: "uppercase", padding: "4px 10px" }}>Produk &amp; Pesanan</div>
                {[
                  { label: "Produk Shopee", sub: "Shopee → Bisnis Saya → Performa Produk → Export", icon: "📦", action: () => prodFileRef.current.click(), accent: C.good },
                  { label: "Pesanan Shopee", sub: "Shopee → Pesanan → Export Pesanan → XLSX", icon: "📍", action: () => orderFileRef.current.click(), accent: "#a29bfe" },
                ].map((item, i) => (
                  <button key={i} style={S.importMenuItem} onClick={() => { item.action(); setImportMenuOpen(false); }}>
                    <span style={{ fontSize: 16, lineHeight: 1, minWidth: 20 }}>{item.icon}</span>
                    <div>
                      <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: item.accent }}>{item.label}</div>
                      <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>{item.sub}</div>
                    </div>
                  </button>
                ))}

                <div style={{ borderTop: `1px solid ${C.line}`, margin: "6px 0" }} />

                {/* Group: COGS */}
                <div style={{ fontFamily: sans, fontSize: 10, fontWeight: 700, color: C.dim, letterSpacing: 1, textTransform: "uppercase", padding: "4px 10px" }}>COGS / HPP</div>
                <button style={S.importMenuItem} onClick={() => { cogsFileRef.current.click(); setImportMenuOpen(false); }}>
                  <span style={{ fontSize: 16, lineHeight: 1, minWidth: 20 }}>💰</span>
                  <div>
                    <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.good }}>Template COGS</div>
                    <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>Download template → isi HPP → import balik</div>
                  </div>
                </button>

                <div style={{ borderTop: `1px solid ${C.line}`, margin: "6px 0" }} />
                <button style={{ ...S.importMenuItem, opacity: 0.7 }} onClick={() => { confirm("Reset semua data? Semua snapshot iklan, data produk, COGS, dan stok akan dihapus.", deleteAllData); setImportMenuOpen(false); }}>
                  <span style={{ fontSize: 16, minWidth: 20 }}>↺</span>
                  <div>
                    <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.bad }}>Reset Semua Data</div>
                    <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>Hapus semua data, mulai dari awal</div>
                  </div>
                </button>
              </div>
            )}
          </div>
          <input ref={adFileRef} type="file" accept=".csv" hidden onChange={handleAdFile} />
          <input ref={prodFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleProdFile} />
          <input ref={incomeFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleIncomeFile} />
          <input ref={tiktokFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleTikTokFile} />
          <input ref={orderFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleOrderFile} />
          <input ref={cogsFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleCogsFile} />
          <input ref={metaFileRef} type="file" accept=".csv" hidden onChange={handleMetaFile} />
          <input ref={tiktokAdsFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleTikTokAdsFile} />
          <input ref={tiktokCampaignFileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleTikTokCampaignFile} />
        </div>
      </header>

      {panduanOpen && <PanduanModal onClose={() => setPanduanOpen(false)} onGoToOverview={() => { setPanduanOpen(false); setTab("overview"); }} />}
      {(hasAds || hasProd) && (
        <div style={S.srcStrip}>
          <span style={S.srcChip}>
            <span style={{ ...S.srcDot, background: hasAds ? C.good : C.dim }} />
            File iklan {hasAds ? `· ${active.periodEnd}` : " -  belum ada"}
          </span>
          <span style={S.srcChip}>
            <span style={{ ...S.srcDot, background: hasProd ? C.good : C.dim }} />
            File produk {hasProd ? `· ${activeProducts.periodEnd}` : " -  belum ada"}
          </span>
          {hasAds && hasProd && <span style={S.srcFull}>✓ mode lengkap: iklan × produk</span>}
          {syncStatus && !syncStatus.ok && (
            <span style={{ fontFamily: mono, fontSize: 11, color: C.watch, border: `1px solid ${C.watch}44`, background: C.watch+"12", padding: "3px 9px", borderRadius: 5 }}>{syncStatus.msg}</span>
          )}
          {syncStatus && syncStatus.ok && (
            <span style={{ fontFamily: mono, fontSize: 11, color: C.good, border: `1px solid ${C.good}44`, background: C.good+"12", padding: "3px 9px", borderRadius: 5 }}>{syncStatus.msg}</span>
          )}
        </div>
      )}

      {/* snapshot strip */}
      {snapshots.length > 0 && (
        <div style={S.snapStrip}>
          <span style={S.snapStripLabel}>SNAPSHOT IKLAN</span>
          {snapshots.map((s) => (
            <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 2 }}>
              <button
                onClick={() => setActiveId(s.id)}
                style={{ ...S.snapChip, ...(s.id === activeId ? S.snapChipActive : {}), borderRadius: "20px 0 0 20px", borderRight: "none" }}>
                {s.periodEnd || s.fileName}
              </button>
              <button
                onClick={() => confirm("Hapus snapshot " + (s.periodEnd || s.fileName) + "?", () => deleteSnapshot(s.id))}
                title="Hapus snapshot"
                style={{ ...S.snapChip, borderRadius: "0 20px 20px 0", padding: "5px 8px", color: C.dim, fontSize: 11, background: s.id === activeId ? C.accent+"1a" : C.panel, borderColor: s.id === activeId ? C.accent+"66" : C.line }}>
                ✕
              </button>
            </div>
          ))}
        </div>
      )}

      {/* tabs */}
      <nav style={S.tabs}>
        {[["overview", "Overview"], ["strategy", "Strategi"], ["fee", "Fee Marketplace"], ["performance", "Performa Iklan"], ["product", "Performa Produk"], ["inventory", "Inventory"], ["forecast", "Forecast & Stok"], ["cogs", "COGS / Margin"], ["area", "Peta Distribusi"], ["unitec", "Unit Economics"], ["cm", "Contribution Margin"], ["pnl", "Simulasi L/R"]].map(([k, lbl]) => (
          <button key={k} onClick={() => setTab(k)}
            style={{ ...S.tab, ...(tab === k ? S.tabActive : {}) }}>{lbl}</button>
        ))}
      </nav>

      <main style={S.main}>
        {!hasAds && !hasProd && ["overview","performance","strategy","forecast","cogs"].includes(tab) && <EmptyState onImportAds={() => adFileRef.current.click()} onImportProd={() => prodFileRef.current.click()} />}
        {hasAds && tab === "overview" && <Overview active={active} snapshots={snapshots} marginFor={marginFor} thresholds={thresholds} activeProducts={activeProducts} shopeeSnap={incomeSnaps[0]||null} tiktokSnap={tiktokSnaps[0]||null} orderSnap={orderSnaps[0]||null} metaSnap={metaSnaps[0]||null} tiktokAdsSnap={tiktokAdsSnaps[0]||null} cogsItems={cogsItems} onGoToCogs={() => setTab("cogs")} />}
        {hasAds && tab === "performance" && <Performance active={active} marginFor={marginFor} thresholds={thresholds} saveThresholds={saveThresholds} activeProducts={activeProducts} metaSnap={metaSnaps[0]||null} tiktokAdsSnap={tiktokAdsSnaps[0]||null} tiktokCampaignSnap={tiktokCampaignSnap} onImportTiktokCampaign={() => tiktokCampaignFileRef.current.click()} />}
        {tab === "product" && (hasProd
          ? <ProductTab snap={activeProducts} thresholds={thresholds} active={active} feeRate={(incomeSnaps[0]?.summary?.feeRateNetGmv) || (tiktokSnaps[0]?.summary?.feeRateNetGmv) || 0} />
          : <ProductEmpty onImport={() => prodFileRef.current.click()} />)}
        {tab === "inventory" && (hasProd
          ? <InventoryTab snap={activeProducts} active={active} thresholds={thresholds} stockMap={stockMap} setStockMap={setStockMap} />
          : <ProductEmpty onImport={() => prodFileRef.current.click()} />)}
        {tab === "strategy" && !hasAds && (
          <div style={{ padding: "40px 0", textAlign: "center" }}>
            <div style={{ fontFamily: mono, fontSize: 13, fontWeight: 700, color: C.muted, marginBottom: 12 }}>Strategi butuh data dari tab lain dulu</div>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.muted, lineHeight: 1.7, maxWidth: 480, margin: "0 auto", marginBottom: 20 }}>
              Untuk analisa strategi yang komprehensif, import minimal:<br/>
              <b style={{ color: C.ink }}>Shopee Ads</b> (wajib) · <b style={{ color: C.ink }}>Performa Produk</b> · <b style={{ color: C.ink }}>Penghasilan Shopee</b> · <b style={{ color: C.ink }}>COGS</b><br/>
              Makin banyak data yang diimport, makin tajam rekomendasinya.
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 10, maxWidth: 640, margin: "0 auto", textAlign: "left" }}>
              {[
                { icon: "📢", label: "Shopee Ads", desc: "Wajib · basis diagnosa iklan", done: hasAds },
                { icon: "📦", label: "Performa Produk", desc: "Stok & BCG Matrix", done: !!activeProducts },
                { icon: "💸", label: "Penghasilan Shopee/TikTok", desc: "Fee rate & retur", done: incomeSnaps.length > 0 || tiktokSnaps.length > 0 },
                { icon: "🧮", label: "COGS / HPP", desc: "Break-even & margin", done: cogsItems?.length > 0 },
              ].map((item, i) => (
                <div key={i} style={{ background: C.panel, border: `1px solid ${item.done ? C.good : C.line}`, borderLeft: `3px solid ${item.done ? C.good : C.dim}`, borderRadius: "0 8px 8px 0", padding: "10px 14px" }}>
                  <div style={{ fontFamily: sans, fontSize: 12, fontWeight: 700, color: item.done ? C.good : C.ink }}>{item.icon} {item.label} {item.done ? "✓" : ""}</div>
                  <div style={{ fontFamily: sans, fontSize: 11, color: C.muted }}>{item.desc}</div>
                </div>
              ))}
            </div>
          </div>
        )}
        {hasAds && tab === "strategy" && <Strategy active={active} marginFor={marginFor} thresholds={thresholds} activeProducts={activeProducts} cogsItems={cogsItems} shopeeSnap={incomeSnaps[0]||null} tiktokSnap={tiktokSnaps[0]||null} tiktokAdsSnap={tiktokAdsSnaps[0]||null} metaSnap={metaSnaps[0]||null} incomeSnaps={incomeSnaps} tiktokSnaps={tiktokSnaps} />}
        {tab === "forecast" && (hasAds
          ? <Forecast active={active} activeProducts={activeProducts} thresholds={thresholds} marginFor={marginFor} />
          : <div style={S.cmEmpty}>Forecast butuh file iklan untuk benchmark ROAS/CVR. Import <b>File Iklan (CSV)</b> dulu{!hasProd ? ", idealnya plus File Produk untuk alokasi per kuadran BCG." : "."}</div>)}
        {tab === "cogs" && <CogsTab active={active} cogsMap={cogsMap} setCogsMap={setCogsMap} keyFor={keyFor} flash={flash} cogsItems={cogsItems} onImportCogs={() => cogsFileRef.current.click()} onDownloadTemplate={() => generateCOGSTemplate()} />}
        {tab === "fee" && (incomeSnaps.length > 0 || tiktokSnaps.length > 0
          ? <FeeTab shopeeSnap={incomeSnaps[0] || null} tiktokSnap={tiktokSnaps[0] || null} />
          : <FeeEmpty onImportShopee={() => incomeFileRef.current.click()} onImportTiktok={() => tiktokFileRef.current.click()} />)}
        {tab === "area" && (orderSnaps.length > 0
          ? <AreaTab snap={orderSnaps[0]} />
          : <AreaEmpty onImport={() => orderFileRef.current.click()} />)}
        {tab === "unitec" && <UnitEconomicsTab active={active} activeProducts={activeProducts} cogsItems={cogsItems} shopeeSnap={incomeSnaps[0]||null} tiktokSnap={tiktokSnaps[0]||null} onDownloadTemplate={() => generateCOGSTemplate()} onImportCogs={() => cogsFileRef.current.click()} />}
        {tab === "cm" && <ContribMarginTab active={active} activeProducts={activeProducts} cogsItems={cogsItems} shopeeSnap={incomeSnaps[0]||null} tiktokSnap={tiktokSnaps[0]||null} onDownloadTemplate={() => generateCOGSTemplate()} onImportCogs={() => cogsFileRef.current.click()} />}
        {tab === "pnl" && <PnLTab
          active={active}
          shopeeSnap={incomeSnaps[0]||null}
          tiktokSnap={tiktokSnaps[0]||null}
          metaSnap={metaSnaps[0]||null}
          tiktokAdsSnap={tiktokAdsSnaps[0]||null}
          cogsItems={cogsItems}
          pnlInputs={pnlInputs}
          setPnlInputs={setPnlInputs}
          onGoToCogs={() => setTab("cogs")}
        />}
        {!hasAds && hasProd && tab !== "product" && (
          <div style={S.cmEmpty}>Tab ini butuh file iklan. Import <b>File Iklan (CSV)</b>, atau buka tab <b>Produk</b> untuk analisa dari file produk.</div>
        )}
      </main>

      {toast && <div style={S.toast}>{toast}</div>}

      {/* Custom confirm dialog (replaces window.confirm which is blocked in iframes) */}
      {confirmDlg && (
        <div style={S.dlgOverlay}>
          <div style={S.dlgBox}>
            <div style={S.dlgMsg}>{confirmDlg.msg}</div>
            <div style={S.dlgBtns}>
              <button style={S.dlgCancel} onClick={() => setConfirmDlg(null)}>Batal</button>
              <button style={S.dlgOk} onClick={() => { setConfirmDlg(null); confirmDlg.onOk(); }}>Ya, lanjutkan</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- Empty states -------------------------------------------------- */
function EmptyState({ onImportAds, onImportProd }) {
  const steps = [
    { num: "01", label: "File Iklan (CSV)", desc: "Shopee Ads → Laporan → Data Keseluruhan Iklan", color: C.accent, action: null, btnLabel: null },
    { num: "02", label: "File Produk (XLSX)", desc: "Shopee → Bisnis Saya → Performa Produk → Export", color: C.good, action: null, btnLabel: null },
    { num: "03", label: "Penghasilan Shopee (XLSX)", desc: "Shopee → Keuangan → Penghasilan Saya → Download", color: "#f96d00", action: null, btnLabel: null },
    { num: "04", label: "Penghasilan TikTok (XLSX)", desc: "TikTok Shop → Keuangan → Pendapatan → Download", color: C.muted, action: null, btnLabel: null },
    { num: "05", label: "Template COGS / HPP", desc: "Download template dari Import Data, isi HPP per SKU, import balik", color: C.good, action: null, btnLabel: null },
  ];
  return (
    <div style={S.empty}>
      <div style={S.emptyMark}>▦</div>
      <h2 style={S.emptyTitle}>Mulai dengan import data</h2>
      <p style={S.emptyText}>Sellio bekerja dengan file export dari marketplace kamu — Shopee, TikTok Shop, dan Meta Ads. Tidak perlu input manual. Klik <b>+ Import Data ▼</b> di kanan atas untuk mulai.</p>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 20, maxWidth: 560, width: "100%" }}>
        {steps.map((step, i) => (
          <div key={i} style={{ background: C.panel, border: `1px solid ${step.color}33`, borderLeft: `3px solid ${step.color}`, borderRadius: 8, padding: "14px 16px" }}>
            <div style={{ fontFamily: mono, fontSize: 10, color: step.color, letterSpacing: 1, marginBottom: 6 }}>{step.num}</div>
            <div style={{ fontFamily: mono, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 4 }}>{step.label}</div>
            <div style={{ fontFamily: sans, fontSize: 11.5, color: C.muted, lineHeight: 1.5, marginBottom: 12 }}>{step.desc}</div>

          </div>
        ))}
      </div>
    </div>
  );
}
function ProductEmpty({ onImport }) {
  return (
    <div style={S.empty}>
      <div style={S.emptyMark}>◳</div>
      <h2 style={S.emptyTitle}>Belum ada data produk</h2>
      <p style={S.emptyText}>
        Export <b>Detail SKU Induk</b> dari Shopee (Bisnis Saya → Performa Produk), lalu import via tombol <b>+ Import Data</b> di kanan atas → File Produk.
      </p>
      <button style={S.importBtn2Big} onClick={onImport}>+ Import File Produk</button>
    </div>
  );
}

/* ---------- Aggregate helper ---------------------------------------------- */
function aggregate(ads) {
  const a = ads.reduce((t, x) => ({
    spend: t.spend + x.spend, gmv: t.gmv + x.gmv, conv: t.conv + x.conv,
    sold: t.sold + x.sold, clicks: t.clicks + x.clicks, impr: t.impr + x.impr,
  }), { spend: 0, gmv: 0, conv: 0, sold: 0, clicks: 0, impr: 0 });
  a.roas = a.spend ? a.gmv / a.spend : 0;
  a.acos = a.gmv ? (a.spend / a.gmv) * 100 : 0;
  a.ctr = a.impr ? (a.clicks / a.impr) * 100 : 0;
  a.cvr = a.clicks ? (a.conv / a.clicks) * 100 : 0;
  a.cpa = a.conv ? a.spend / a.conv : 0;
  return a;
}

/* ---------- Overview ------------------------------------------------------ */
function Overview({ active, snapshots, marginFor, thresholds, activeProducts, shopeeSnap, tiktokSnap, orderSnap, metaSnap, tiktokAdsSnap, cogsItems, onGoToCogs }) {
  const agg = useMemo(() => aggregate(active.ads), [active]);

  // ad dependency: ad-attributed GMV vs total product sales (if product file present)
  const adDep = useMemo(() => {
    if (!activeProducts) return null;
    const totalSales = activeProducts.products.reduce((t, p) => t + p.sales, 0);
    if (!totalSales) return null;
    return { adGmv: agg.gmv, totalSales, pct: Math.min(100, (agg.gmv / totalSales) * 100) };
  }, [activeProducts, agg]);

  // contribution margin = GMV*margin% - spend, where margin known
  const cm = useMemo(() => {
    let knownGmv = 0, grossProfit = 0, unknown = 0;
    active.ads.forEach((ad) => {
      const m = marginFor(ad);
      if (m == null) { unknown += ad.gmv; return; }
      knownGmv += ad.gmv;
      grossProfit += ad.gmv * (m / 100);
    });
    const contribution = grossProfit - active.ads.reduce((t, ad) => t + (marginFor(ad) != null ? ad.spend : 0), 0);
    return { knownGmv, grossProfit, contribution, unknown };
  }, [active, marginFor]);

  // trend vs previous snapshot
  const prev = snapshots.filter((s) => s.endTs < active.endTs).sort((a, b) => b.endTs - a.endTs)[0];
  const prevAgg = prev ? aggregate(prev.ads) : null;
  const delta = (cur, p) => (p ? ((cur - p) / p) * 100 : null);

  // ── derived data for new sections ──
  const safeNum = v => (typeof v === "number" && !isNaN(v) ? v : 0);
  const sh = shopeeSnap?.summary;
  const tt = tiktokSnap?.summary;
  const totalNetGmv  = safeNum(sh?.netGmv) + safeNum(tt?.netGmv);
  const totalFee     = safeNum(sh?.totalFee) + safeNum(tt?.totalFee);
  const adGmv        = active ? active.ads.reduce((t,a) => t+a.gmv, 0) : 0;
  const shopeeSpend  = active ? active.ads.reduce((t,a) => t+a.spend, 0) : 0;
  const ttAdsSpend   = tiktokAdsSnap?.summary?.totalSpend || 0;
  const metaSpend    = metaSnap?.summary?.totalSpend || 0;
  const totalAdSpend = shopeeSpend + ttAdsSpend + metaSpend;
  const cogsOk       = cogsItems && cogsItems.length > 0;

  // Daily sales from order file
  const dailySales = useMemo(() => {
    if (!orderSnap) return [];
    const map = {};
    orderSnap.orders.forEach(o => {
      if (!o.tglDibuat) return;
      const d = o.tglDibuat.slice(0,10);
      if (!map[d]) map[d] = { tgl: d, pesanan: 0, pcs: 0 };
      map[d].pesanan++; map[d].pcs += o.jumlah;
    });
    return Object.values(map).sort((a,b) => a.tgl < b.tgl ? -1 : 1);
  }, [orderSnap]);

  // Top 5 kota & provinsi from order file
  const { topKota, topProv } = useMemo(() => {
    if (!orderSnap) return { topKota: [], topProv: [] };
    const kt = {}, pv = {};
    orderSnap.orders.forEach(o => {
      kt[o.kota] = (kt[o.kota]||0) + 1;
      pv[o.provinsi] = (pv[o.provinsi]||0) + 1;
    });
    const topKota = Object.entries(kt).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([n,c])=>({name:n,count:c}));
    const topProv = Object.entries(pv).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([n,c])=>({name:n,count:c}));
    return { topKota, topProv };
  }, [orderSnap]);

  // Top 5 customers
  const topCustomers = useMemo(() => {
    if (!orderSnap) return [];
    const map = {};
    orderSnap.orders.forEach(o => {
      const u = o.skuInduk ? o.skuInduk : "unknown"; // use a buyer field
      // Actually order snap has no buyer username — skip for now, use kota as proxy
    });
    // Parse from raw — orderSnap.orders has no buyer field currently
    return [];
  }, [orderSnap]);

  // Top colors from order Nama Variasi
  const topColors = useMemo(() => {
    if (!orderSnap) return [];
    const map = {};
    const colorWords = ["black","white","navy","beige","grey","gray","green","blue","red","brown","cream","khaki","olive","maroon","yellow","pink","purple","orange","hitam","putih","abu","biru","merah","hijau","coklat","krem"];
    orderSnap.orders.forEach(o => {
      if (!o.namaProduk) return;
      const lower = o.namaProduk.toLowerCase();
      colorWords.forEach(c => {
        if (lower.includes(c)) map[c] = (map[c]||0) + o.jumlah;
      });
    });
    return Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([c,n])=>({color:c,count:n}));
  }, [orderSnap]);

  return (
    <div>
      {/* 1. REVENUE AKTUAL */}
      {(shopeeSnap || tiktokSnap) && (
        <>
          <SectionLabel>REVENUE AKTUAL · SEMUA CHANNEL</SectionLabel>
          <DataDisclaimer />
          <div style={S.kpiGrid}>
            <Kpi label="Total Net GMV" value={rpShort(totalNetGmv)} dir="rev" accent={C.accent} big />
            {totalNetGmv > 0 && <Kpi label="Fee Rate Marketplace" value={(Math.abs(totalFee)/totalNetGmv*100).toFixed(1)+"%"} dir="cost" accent={C.bad} />}
            <Kpi label="Total Fee Marketplace" value={rpShort(Math.abs(totalFee))} dir="cost" accent={C.bad} />
            {sh && <Kpi label="Shopee GMV" value={rpShort(sh.netGmv||0)} dir="rev" accent={C.accent} />}
            {tt && <Kpi label="TikTok GMV" value={rpShort(tt.netGmv||0)} dir="rev" accent={"#fe2c55"} />}
            {active && <Kpi label="Revenue dari Iklan" value={rpShort(adGmv)} dir="rev" accent={C.muted} />}
            {totalAdSpend > 0 && <Kpi label="Total Ad Spending" value={rpShort(totalAdSpend)} dir="cost" accent={C.watch} />}
            {cogsOk
              ? <Kpi label="COGS Coverage" value={cogsItems.length+" SKU"} dir="neutral" accent={C.good} />
              : <div style={{ background: C.panel, border: `1px solid ${C.watch}44`, borderRadius: 8, padding: "13px 15px", cursor: "pointer", borderTop: `3px solid ${C.watch}` }} onClick={onGoToCogs}>
                  <div style={{ fontFamily: mono, fontSize: 10, color: C.watch, letterSpacing: 0.8, marginBottom: 6 }}>COGS / UNIT ECONOMICS</div>
                  <div style={{ fontFamily: mono, fontSize: 13, color: C.watch }}>Belum diisi →</div>
                  <div style={{ fontFamily: sans, fontSize: 11, color: C.dim, marginTop: 3 }}>Klik untuk isi COGS & unlock CM</div>
                </div>
            }
          </div>

          {/* Daily sales chart */}
          {dailySales.length > 0 && (() => {
            const maxPesanan = Math.max(...dailySales.map(d=>d.pesanan), 1);
            return (
              <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: "14px 16px", marginTop: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 12 }}>
                  <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1 }}>DAILY ORDERS · SHOPEE{tt ? "" : " ONLY"}</div>
                  <div style={{ fontFamily: mono, fontSize: 11, color: C.dim }}>{orderSnap.periodStart} → {orderSnap.periodEnd}</div>
                </div>
                {(() => {
                  const W = 900, H = 72, pad = 4;
                  const minV = 0, maxV = maxPesanan;
                  const pts = dailySales.map((d, i) => {
                    const x = pad + (i / Math.max(dailySales.length-1,1)) * (W - pad*2);
                    const y = H - pad - ((d.pesanan - minV) / Math.max(maxV - minV, 1)) * (H - pad*2);
                    return `${x},${y}`;
                  }).join(" ");
                  const areaBottom = dailySales.map((d, i) => {
                    const x = pad + (i / Math.max(dailySales.length-1,1)) * (W - pad*2);
                    return `${x},${H-pad}`;
                  });
                  const areaPath = `M ${pts.split(" ")[0]} L ${pts.split(" ").join(" L ")} L ${areaBottom.reverse().join(" L ")} Z`;
                  return (
                    <svg viewBox={`0 0 ${W} ${H}`} style={{ width:"100%", height:72, overflow:"visible" }} preserveAspectRatio="none">
                      <defs>
                        <linearGradient id="lineGrad" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor={C.accent} stopOpacity="0.3"/>
                          <stop offset="100%" stopColor={C.accent} stopOpacity="0.02"/>
                        </linearGradient>
                      </defs>
                      <path d={areaPath} fill="url(#lineGrad)" />
                      <polyline points={pts} fill="none" stroke={C.accent} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
                      {dailySales.map((d, i) => {
                        const x = pad + (i / Math.max(dailySales.length-1,1)) * (W - pad*2);
                        const y = H - pad - ((d.pesanan - minV) / Math.max(maxV - minV, 1)) * (H - pad*2);
                        return d.pesanan === maxPesanan
                          ? <circle key={i} cx={x} cy={y} r="4" fill={C.accent} stroke={C.panel} strokeWidth="2" />
                          : null;
                      })}
                    </svg>
                  );
                })()}
                <div style={{ display:"flex", justifyContent:"space-between", marginTop:5, fontFamily:mono, fontSize:10, color:C.dim }}>
                  <span>{dailySales[0]?.tgl?.slice(5)}</span>
                  <span style={{ color:C.muted }}>{orderSnap.orders.length} pesanan total · {orderSnap.orders.reduce((t,o)=>t+o.jumlah,0)} pcs</span>
                  <span>{dailySales[dailySales.length-1]?.tgl?.slice(5)}</span>
                </div>
              </div>
            );
          })()}

          {/* Top 5 area + colors */}
          {orderSnap && (topKota.length > 0 || topProv.length > 0) && (
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10, marginTop:10 }}>
              {/* Top Provinsi */}
              <div style={{ background:C.panel, border:`1px solid ${C.line}`, borderRadius:8, padding:"12px 14px" }}>
                <div style={{ fontFamily:mono, fontSize:10, color:C.muted, letterSpacing:1, marginBottom:10 }}>TOP 5 PROVINSI</div>
                {topProv.map((p,i) => (
                  <div key={i} style={{ marginBottom:8 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                      <span style={{ fontFamily:sans, fontSize:12, fontWeight:i<3?700:400 }}>{p.name}</span>
                      <span style={{ fontFamily:mono, fontSize:12, color:C.accent }}>{p.count}</span>
                    </div>
                    <div style={{ height:4, background:C.panel2, borderRadius:2, overflow:"hidden" }}>
                      <div style={{ width:(p.count/topProv[0].count*100)+"%", height:"100%", background:[C.accent,C.good,C.watch,C.muted,C.dim][i]||C.dim, borderRadius:2 }} />
                    </div>
                  </div>
                ))}
              </div>
              {/* Top Kota */}
              <div style={{ background:C.panel, border:`1px solid ${C.line}`, borderRadius:8, padding:"12px 14px" }}>
                <div style={{ fontFamily:mono, fontSize:10, color:C.muted, letterSpacing:1, marginBottom:10 }}>TOP 5 KOTA</div>
                {topKota.map((p,i) => (
                  <div key={i} style={{ marginBottom:8 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                      <span style={{ fontFamily:sans, fontSize:12, fontWeight:i<3?700:400 }}>{p.name.replace("KOTA ","").replace("KAB. ","")}</span>
                      <span style={{ fontFamily:mono, fontSize:12, color:C.good }}>{p.count}</span>
                    </div>
                    <div style={{ height:4, background:C.panel2, borderRadius:2, overflow:"hidden" }}>
                      <div style={{ width:(p.count/topKota[0].count*100)+"%", height:"100%", background:[C.good,C.accent,C.watch,C.muted,C.dim][i]||C.dim, borderRadius:2 }} />
                    </div>
                  </div>
                ))}
              </div>
              {/* Top Warna */}
              {topColors.length > 0 && (
                <div style={{ background:C.panel, border:`1px solid ${C.line}`, borderRadius:8, padding:"12px 14px" }}>
                  <div style={{ fontFamily:mono, fontSize:10, color:C.muted, letterSpacing:1, marginBottom:10 }}>TOP WARNA (dari nama produk)</div>
                  {topColors.map((p,i) => (
                    <div key={i} style={{ marginBottom:8 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                        <span style={{ fontFamily:sans, fontSize:12, fontWeight:i<3?700:400, textTransform:"capitalize" }}>{p.color}</span>
                        <span style={{ fontFamily:mono, fontSize:12, color:C.watch }}>{p.count} pcs</span>
                      </div>
                      <div style={{ height:4, background:C.panel2, borderRadius:2, overflow:"hidden" }}>
                        <div style={{ width:(p.count/topColors[0].count*100)+"%", height:"100%", background:[C.watch,C.accent,C.good,C.muted,C.dim,C.bad][i]||C.dim, borderRadius:2 }} />
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </>
      )}

      {/* 2. PERFORMA IKLAN */}
      <SectionLabel>PERFORMA IKLAN · {active.periodStart} → {active.periodEnd}</SectionLabel>
      <div style={S.kpiGrid}>
        <Kpi label="Spend Shopee Ads" value={rpShort(shopeeSpend)} dir="cost" accent={C.watch} />
        {tiktokAdsSnap && <Kpi label="Spend TikTok GMV Max" value={rpShort(ttAdsSpend)} dir="cost" accent={"#fe2c55"} />}
        {metaSnap && <Kpi label="Spend Meta Ads" value={rpShort(metaSpend)} dir="cost" accent={"#1877f2"} />}
        <Kpi label="GMV dari Iklan" value={rpShort(agg.gmv)} sub={prevAgg ? deltaStr(delta(agg.gmv, prevAgg.gmv)) : null} dir="rev" accent={C.accent} />
        <Kpi label="ROAS Shopee" value={agg.roas.toFixed(2) + "x"} sub={prevAgg ? deltaStr(delta(agg.roas, prevAgg.roas)) : null} dir="rev" big accent={agg.roas >= thresholds.targetRoas ? C.good : C.bad} />
        <Kpi label="ACOS" value={agg.acos.toFixed(1) + "%"} dir="cost" accent={C.watch} />
        <Kpi label="Konversi" value={nfmt(agg.conv)} dir="rev" accent={C.good} />
        <Kpi label="CVR" value={agg.cvr.toFixed(2) + "%"} dir="neutral" accent={C.muted} />
        <Kpi label="CPA" value={rpShort(agg.cpa)} dir="cost" accent={C.muted} />
      </div>

      {/* 3. KETERGANTUNGAN IKLAN */}
      {adDep && (
        <>
          <SectionLabel>KETERGANTUNGAN IKLAN</SectionLabel>
          <div style={S.cmRow}>
            <div style={S.cmCard}>
              <div style={S.cmLabel}>Penjualan total (semua channel)</div>
              <div style={S.cmVal}>{rp(adDep.totalSales)}</div>
            </div>
            <div style={S.cmCard}>
              <div style={S.cmLabel}>GMV dari iklan</div>
              <div style={S.cmVal}>{rp(adDep.adGmv)}</div>
            </div>
            <div style={{ ...S.cmCard, ...S.cmCardFinal }}>
              <div style={S.cmLabel}>% penjualan via iklan</div>
              <div style={{ ...S.cmVal, color: adDep.pct > 70 ? C.watch : C.good }}>{adDep.pct.toFixed(0)}%</div>
            </div>
            <div style={S.cmNote}>
              {adDep.pct > 70
                ? `⚠ ${adDep.pct.toFixed(0)}% penjualan bergantung iklan  -  organik tipis, rawan kalau spend dipotong.`
                : `${(100 - adDep.pct).toFixed(0)}% penjualan datang organik  -  basis sehat.`}
            </div>
          </div>
        </>
      )}

      {/* 4. CONTRIBUTION MARGIN */}
      <SectionLabel>CONTRIBUTION MARGIN</SectionLabel>
      <div style={S.cmRow}>
        {(() => {
          // Try marginFor first, fallback to cogsItems direct lookup
          const cogsMap2 = {};
          (cogsItems || []).forEach(item => {
            if (item.harga && item.total) {
              const pct = (item.harga - item.total) / item.harga * 100;
              cogsMap2[item.sku] = pct;
              cogsMap2[item.nama] = pct;
            }
          });
          const getMargin = (ad) => {
            const m = marginFor(ad);
            if (m != null) return m;
            // fallback: match by ad name against cogsItems
            if (!ad.name) return null;
            const name = ad.name.toLowerCase();
            for (const [k, v] of Object.entries(cogsMap2)) {
              if (k && name.includes(k.toLowerCase().slice(0,10))) return v;
            }
            return null;
          };
          let grossProfit = 0, adSpendKnown = 0, knownGmv = 0, unknownGmv = 0;
          active.ads.forEach(ad => {
            const m = getMargin(ad);
            if (m != null) {
              grossProfit += ad.gmv * (m / 100);
              adSpendKnown += ad.spend;
              knownGmv += ad.gmv;
            } else {
              unknownGmv += ad.gmv;
            }
          });
          const contribution = grossProfit - adSpendKnown;
          const hasData = knownGmv > 0 || (cogsItems?.length > 0 && cm.knownGmv > 0);

          if (!hasData) return (
            <div style={S.cmEmpty}>
              Set COGS / margin per produk di tab <b>COGS / Margin</b> untuk menghitung contribution margin.
            </div>
          );
          return (
            <>
              <div style={S.cmCard}>
                <div style={S.cmLabel}>Gross profit (GMV × margin)</div>
                <div style={S.cmVal}>{rp(knownGmv > 0 ? grossProfit : cm.grossProfit)}</div>
              </div>
              <div style={S.cmCard}>
                <div style={S.cmLabel}>− Ad spend</div>
                <div style={S.cmVal}>{rp(knownGmv > 0 ? adSpendKnown : active.ads.reduce((t,ad) => t+(marginFor(ad)!=null?ad.spend:0),0))}</div>
              </div>
              <div style={{ ...S.cmCard, ...S.cmCardFinal }}>
                <div style={S.cmLabel}>Contribution margin</div>
                <div style={{ ...S.cmVal, color: (knownGmv > 0 ? contribution : cm.contribution) >= 0 ? C.good : C.bad }}>{rp(knownGmv > 0 ? contribution : cm.contribution)}</div>
              </div>
              {unknownGmv > 0 && <div style={S.cmNote}>⚠ {rpShort(unknownGmv)} GMV belum punya COGS — belum dihitung.</div>}
            </>
          );
        })()}
      </div>

      {/* 5. STRATEGI HIGHLIGHT */}
      {(() => {
        const buckets = { scale: [], hold: [], fix: [] };
        active.ads.forEach(ad => {
          const m = marginFor(ad);
          const dx = diagnose(ad, m, thresholds);
          if (dx.some(d => d.tag === "SCALE")) buckets.scale.push(ad);
          else if (dx.some(d => d.tag === "DI BAWAH TARGET" || d.tag === "MAKAN MARGIN")) buckets.hold.push(ad);
          else if (dx.some(d => d.tag === "CTR RENDAH" || d.tag === "CVR RENDAH")) buckets.fix.push(ad);
        });
        const prods = activeProducts?.products || [];
        const lowStock = prods.filter(p => p.stock != null && p.stock < 10);
        const sh = shopeeSnap?.summary;
        const refundAlert = sh?.refundRateValue > 3;
        const feeAlert = sh?.feeRateNetGmv > 22;
        const priorities = [];
        if (buckets.scale.length > 0) priorities.push({ level: "good", text: `${buckets.scale.length} iklan siap scale — naikkan budget 20–30%`, tag: "SCALE" });
        if (buckets.hold.length > 0) priorities.push({ level: "bad", text: `${buckets.hold.length} iklan di bawah target — ketatkan atau pause`, tag: "HOLD" });
        if (lowStock.length > 0) priorities.push({ level: "watch", text: `${lowStock.length} produk stok kritis (<10 unit)`, tag: "RESTOCK" });
        if (refundAlert) priorities.push({ level: "bad", text: `Retur ${sh.refundRateValue.toFixed(1)}% — di atas normal, perlu diaudit`, tag: "RETUR" });
        if (feeAlert) priorities.push({ level: "watch", text: `Fee rate ${sh.feeRateNetGmv.toFixed(1)}% — cek komposisi voucher & promo`, tag: "FEE" });
        if (buckets.fix.length > 0) priorities.push({ level: "watch", text: `${buckets.fix.length} iklan CTR/CVR rendah — perlu fix materi atau listing`, tag: "FIX" });
        if (priorities.length === 0) return null;
        return (
          <>
            <SectionLabel>HIGHLIGHT STRATEGI</SectionLabel>
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 16 }}>
              {priorities.slice(0, 4).map((p, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, background: C.panel, border: `1px solid ${C.line}`, borderLeft: `3px solid ${p.level === "good" ? C.good : p.level === "bad" ? C.bad : C.watch}`, borderRadius: "0 8px 8px 0", padding: "9px 14px" }}>
                  <span style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: p.level === "good" ? C.good : p.level === "bad" ? C.bad : C.watch, background: (p.level === "good" ? C.good : p.level === "bad" ? C.bad : C.watch) + "18", padding: "2px 7px", borderRadius: 4, letterSpacing: 0.5, flexShrink: 0 }}>{p.tag}</span>
                  <span style={{ fontFamily: sans, fontSize: 12.5, color: C.ink }}>{p.text}</span>
                </div>
              ))}
              {priorities.length > 4 && <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, paddingLeft: 4 }}>+{priorities.length - 4} item lain — lihat tab Strategi untuk detail lengkap.</div>}
            </div>
          </>
        );
      })()}

      {/* 6. ACTION PLAN */}
      <SectionLabel>ACTION PLAN</SectionLabel>
      <ActionPlan active={active} marginFor={marginFor} agg={agg} thresholds={thresholds} cogsItems={cogsItems} />
    </div>
  );
}

function deltaStr(d) {
  if (d == null) return null;
  const s = (d >= 0 ? "▲ " : "▼ ") + Math.abs(d).toFixed(0) + "%";
  return { text: s, up: d >= 0 };
}

/* aggregated action plan: pulls worst offenders + scale candidates */
function ActionPlan({ active, marginFor, thresholds, cogsItems }) {
  const items = useMemo(() => {
    const scale = [], cut = [], fix = [];
    active.ads.forEach((ad) => {
      const m = marginFor(ad);
      const dx = diagnose(ad, m, thresholds);
      dx.forEach((d) => {
        if (d.tag === "SCALE") scale.push({ ad, d });
        else if (d.tag === "DI BAWAH TARGET" || d.tag === "MAKAN MARGIN") cut.push({ ad, d });
        else if (d.level === "watch" && (d.tag.includes("CTR") || d.tag.includes("CVR"))) fix.push({ ad, d });
      });
    });
    return { scale, cut, fix };
  }, [active, marginFor, thresholds]);

  const noMargin = active.ads.every((ad) => marginFor(ad) == null) && !(cogsItems?.length > 0);
  if (noMargin) return <div style={S.cmEmpty}>Action plan butuh margin. Isi COGS dulu untuk rekomendasi scale/hold yang akurat.</div>;

  return (
    <div style={S.planGrid}>
      <PlanCol title="SCALE" color={C.good} items={items.scale} />
      <PlanCol title="HOLD / CUT" color={C.bad} items={items.cut} />
      <PlanCol title="FIX MATERI" color={C.watch} items={items.fix} />
    </div>
  );
}
function PlanCol({ title, color, items }) {
  return (
    <div style={S.planCol}>
      <div style={{ ...S.planHead, color }}>{title} · {items.length}</div>
      {items.length === 0 && <div style={S.planEmpty}> - </div>}
      {items.slice(0, 8).map((it, i) => (
        <div key={i} style={S.planItem}>
          <div style={S.planAd}>{it.ad.name}</div>
          <div style={S.planMsg}>{it.d.msg}</div>
        </div>
      ))}
    </div>
  );
}

/* ---------- Performance --------------------------------------------------- */
function Performance({ active, marginFor, thresholds, saveThresholds, metaSnap, tiktokAdsSnap, tiktokCampaignSnap, onImportTiktokCampaign }) {
  const [channel, setChannel] = useState("shopee");
  const [sort, setSort] = useState("spend");
  const [dir, setDir] = useState(-1);
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(null);
  const rows = useMemo(() => {
    let r = active.ads.filter((a) => a.name.toLowerCase().includes(q.toLowerCase()));
    r = [...r].sort((a, b) => (a[sort] < b[sort] ? 1 : -1) * dir);
    return r;
  }, [active, sort, dir, q]);
  const head = (k, lbl, right) => (
    <th style={{ ...S.th, textAlign: right ? "right" : "left", cursor: "pointer" }}
      onClick={() => { if (sort === k) setDir(-dir); else { setSort(k); setDir(-1); } }}>
      {lbl}{sort === k ? (dir === -1 ? " ▾" : " ▴") : ""}
    </th>
  );
  return (
    <div>
      {/* Channel sub-nav */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16, borderBottom: `1px solid ${C.line}`, paddingBottom: 12 }}>
        {[["shopee","Shopee Ads"], ["meta","Meta Ads"], ["tiktokads","TikTok GMV Max"], ["all","Semua Channel"]].map(([k, lbl]) => {
          const disabled = (k === "meta" && !metaSnap) || (k === "tiktokads" && !tiktokAdsSnap);
          return (
            <button key={k} disabled={disabled} onClick={() => setChannel(k)}
              style={{ fontFamily: mono, fontSize: 12, fontWeight: 700, padding: "7px 16px", borderRadius: 6,
                cursor: disabled ? "not-allowed" : "pointer",
                border: `1px solid ${channel===k ? C.accent : C.line}`,
                background: channel===k ? C.accent+"1a" : "transparent",
                color: disabled ? C.dim : channel===k ? C.accent : C.muted,
                opacity: disabled ? 0.4 : 1 }}>
              {lbl}
              {((k === "meta" && !metaSnap) || (k === "tiktokads" && !tiktokAdsSnap)) && <span style={{ fontSize: 10, marginLeft: 6, color: C.dim }}>belum diimport</span>}
            </button>
          );
        })}
      </div>

      {/* Disclaimer */}
      <div style={{ display: "flex", gap: 10, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 8, padding: "12px 14px", marginBottom: 14, fontFamily: sans, fontSize: 11, color: C.muted, lineHeight: 1.6 }}>
        <span style={{ fontSize: 13, flexShrink: 0, marginTop: 1 }}>ℹ️</span>
        <div>
          <b style={{ color: C.ink }}>Data 100% identik dengan file ekspor resmi {channel === "tiktokads" ? "TikTok Shop" : channel === "meta" ? "Meta Ads Manager" : "Shopee Seller Center"}.</b>{" "}
          Sellio membaca file langsung tanpa modifikasi — angka yang tampil di sini adalah angka yang sama persis dengan yang ada di dashboard channel kamu. Jika ada perbedaan dengan laporan internal, kemungkinan bersumber dari perbedaan periode, timezone, atau metode agregasi di sisi platform.
          {" "}<b style={{ color: C.ink }}>Diagnosa dan saran yang ditampilkan adalah gambaran awal berbasis data.</b>{" "}
          Untuk keputusan strategis lebih lanjut, validasi dengan metrics lain sebelum eksekusi.{" "}
          <a href="https://wa.me/6282130311844" target="_blank" rel="noreferrer"
            style={{ color: C.accent, fontWeight: 700, textDecoration: "none" }}>
            Jadwalkan konsultasi →
          </a>
        </div>
      </div>

      {channel === "meta" && metaSnap && <MetaAdsView snap={metaSnap} />}
      {channel === "tiktokads" && tiktokAdsSnap && <TikTokAdsView snap={tiktokAdsSnap} campaignSnap={tiktokCampaignSnap} onImportCampaign={onImportTiktokCampaign} />}
      {channel === "all" && <AllChannelView active={active} metaSnap={metaSnap} tiktokAdsSnap={tiktokAdsSnap} />}
      {channel === "shopee" && <>
      <ThresholdPanel thresholds={thresholds} saveThresholds={saveThresholds} />

      {/* Scatter ROAS vs Spend + Ranking panel */}
      {active.ads.length > 0 && (() => {
        const ads = active.ads.filter(a => a.spend > 0);
        const maxSpend = Math.max(...ads.map(a => a.spend), 1);
        const maxRoas  = Math.max(...ads.map(a => a.roas), 1);
        const W = 340, H = 200, pad = 36;
        const targetRoas = thresholds.targetRoas || 5;
        const cx = a => pad + (a.spend / maxSpend) * (W - pad * 1.5);
        const cy = a => H - pad - (a.roas / (maxRoas * 1.1)) * (H - pad * 1.5);
        // Ranking by ROAS
        const ranked = [...ads].sort((a,b) => b.roas - a.roas).slice(0, 10);
        const maxBarSpend = Math.max(...ranked.map(a => a.spend), 1);
        return (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
            {/* Scatter */}
            <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, padding: "14px 16px" }}>
              <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1.2, marginBottom: 8 }}>ROAS vs SPEND · tiap titik = 1 iklan</div>
              <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", overflow: "visible" }}>
                {/* Grid */}
                {[0, 0.25, 0.5, 0.75, 1].map((r,i) => (
                  <g key={i}>
                    <line x1={pad} x2={W - pad/2} y1={H - pad - r*(H-pad*1.5)} y2={H - pad - r*(H-pad*1.5)} stroke={C.line} strokeWidth="0.5" />
                    <text x={pad - 4} y={H - pad - r*(H-pad*1.5) + 3} textAnchor="end" fontSize="7.5" fill={C.dim} fontFamily="monospace">{(maxRoas * 1.1 * r).toFixed(1)}x</text>
                  </g>
                ))}
                {/* Target ROAS line */}
                {targetRoas <= maxRoas * 1.1 && (
                  <line x1={pad} x2={W - pad/2}
                    y1={H - pad - (targetRoas / (maxRoas * 1.1)) * (H - pad * 1.5)}
                    y2={H - pad - (targetRoas / (maxRoas * 1.1)) * (H - pad * 1.5)}
                    stroke={C.accent} strokeWidth="0.8" strokeDasharray="4,3" />
                )}
                <text x={W - pad/2 + 2} y={H - pad - (targetRoas / (maxRoas * 1.1)) * (H - pad * 1.5) + 3} fontSize="7" fill={C.accent} fontFamily="monospace">target</text>
                {/* X axis */}
                <line x1={pad} x2={W - pad/2} y1={H - pad} y2={H - pad} stroke={C.line} strokeWidth="0.5" />
                <text x={pad} y={H - 4} fontSize="7.5" fill={C.dim} fontFamily="monospace">spend rendah</text>
                <text x={W - pad/2} y={H - 4} textAnchor="end" fontSize="7.5" fill={C.dim} fontFamily="monospace">spend tinggi</text>
                {/* Dots */}
                {ads.map((a, i) => {
                  const x = cx(a), y = cy(a);
                  const col = a.roas >= targetRoas * 1.5 ? C.good : a.roas >= targetRoas ? C.accent : a.roas >= targetRoas * 0.75 ? C.watch : C.bad;
                  const r = Math.max(4, Math.min(10, 4 + (a.spend / maxSpend) * 6));
                  const rowIdx = rows.findIndex(r => r.name === a.name);
                  const isSelected = open === rowIdx;
                  return (
                    <g key={i} style={{ cursor: "pointer" }}
                      onClick={() => {
                        setOpen(isSelected ? null : rowIdx);
                        setTimeout(() => {
                          const el = document.getElementById("adrow-" + rowIdx);
                          if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                        }, 60);
                      }}>
                      <circle cx={x} cy={y} r={isSelected ? r + 3 : r}
                        fill={col} fillOpacity={isSelected ? 1 : 0.75}
                        stroke={isSelected ? "#fff" : col} strokeWidth={isSelected ? 2 : 0.5} />
                      {isSelected && (
                        <text x={x + r + 4} y={y + 3} fontSize="8.5" fontFamily="monospace" fill={col} fontWeight="700">
                          {a.name.length > 20 ? a.name.slice(0,20)+"…" : a.name}
                        </text>
                      )}
                      <title>{`${a.name}\nSpend: ${rpShort(a.spend)}\nROAS: ${a.roas.toFixed(2)}x\nGMV: ${rpShort(a.gmv)}`}</title>
                    </g>
                  );
                })}
              </svg>
              <div style={{ display: "flex", gap: 10, marginTop: 4, flexWrap: "wrap" }}>
                {[[C.good, `≥${(targetRoas*1.5).toFixed(1)}x scale`],[C.accent,`≥${targetRoas}x on target`],[C.watch,"mendekati batas"],[C.bad,"di bawah target"]].map(([c,l])=>(
                  <div key={l} style={{ display:"flex", alignItems:"center", gap:4, fontFamily:mono, fontSize:9, color:C.muted }}>
                    <div style={{ width:8, height:8, borderRadius:"50%", background:c, opacity:0.8 }} />{l}
                  </div>
                ))}
              </div>
            </div>

            {/* Ranking horizontal bar */}
            <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, padding: "14px 16px" }}>
              <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1.2, marginBottom: 10 }}>RANKING ROAS · top 10 iklan</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {ranked.map((a, i) => {
                  const roasCol = a.roas >= targetRoas * 1.5 ? C.good : a.roas >= targetRoas ? C.accent : a.roas >= targetRoas * 0.75 ? C.watch : C.bad;
                  const spendW  = (a.spend / maxBarSpend * 100).toFixed(0);
                  const shortName = a.name.length > 22 ? a.name.slice(0, 22) + "…" : a.name;
                  const rowIdx = rows.findIndex(r => r.name === a.name);
                  const isSelected = open === rowIdx;
                  return (
                    <div key={i} style={{ cursor: "pointer" }} onClick={() => {
                      setOpen(isSelected ? null : rowIdx);
                      setTimeout(() => {
                        const el = document.getElementById("adrow-" + rowIdx);
                        if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                      }, 60);
                    }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3, background: isSelected ? roasCol+"18" : "transparent", borderRadius: 4, padding: "1px 4px" }}>
                        <span style={{ fontFamily: sans, fontSize: 10.5, color: isSelected ? roasCol : C.ink, fontWeight: isSelected ? 700 : 400, flex: 1, overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>{shortName}</span>
                        <div style={{ display: "flex", gap: 8, flexShrink: 0, marginLeft: 8 }}>
                          <span style={{ fontFamily: mono, fontSize: 10, fontWeight: 700, color: roasCol }}>{a.roas.toFixed(1)}x</span>
                          <span style={{ fontFamily: mono, fontSize: 10, color: C.dim }}>{rpShort(a.spend)}</span>
                        </div>
                      </div>
                      <div style={{ height: 5, background: C.panel2, borderRadius: 3, overflow: "hidden" }}>
                        <div style={{ width: spendW + "%", height: "100%", background: roasCol, opacity: isSelected ? 1 : 0.7, borderRadius: 3 }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        );
      })()}
      <div style={S.perfBar}>
        <SectionLabel inline>PER IKLAN · {active.ads.length}</SectionLabel>
        <input placeholder="cari iklan..." value={q} onChange={(e) => setQ(e.target.value)} style={S.search} />
      </div>
      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={{ ...S.th, width: 22 }}></th>
              {head("name", "Iklan")}
              <th style={S.th}>Bidding</th>
              {head("spend", "Spend", 1)}
              {head("gmv", "GMV", 1)}
              {head("roas", "ROAS", 1)}
              {head("acos", "ACOS", 1)}
              {head("ctr", "CTR", 1)}
              {head("cvr", "CVR", 1)}
              {head("conv", "Konv", 1)}
              <th style={{ ...S.th, textAlign: "left" }}>Status</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((ad, i) => {
              const m = marginFor(ad);
              const dx = diagnose(ad, m, thresholds);
              const lvl = worstLevel(dx);
              const isOpen = open === i;
              const effTarget = m != null && thresholds.autoRoasFromMargin ? Math.max(thresholds.targetRoas, 100 / m) : thresholds.targetRoas;
              return (
                <React.Fragment key={i}>
                  <tr style={{ ...S.tr, cursor: "pointer" }} id={"adrow-" + i} onClick={() => setOpen(isOpen ? null : i)}>
                    <td style={{ ...S.td, color: C.muted, fontFamily: mono }}>{isOpen ? "▾" : "▸"}</td>
                    <td style={S.td}>
                      <div style={S.adName}>{ad.name}</div>
                      <div style={S.adMeta}>{ad.status} · {ad.type}{m != null ? ` · margin ${m}%` : " · margin  - "}</div>
                    </td>
                    <td style={S.td}><span style={S.biddingTag}>{ad.bidding}</span></td>
                    <td style={S.tdR}>{rpShort(ad.spend)}</td>
                    <td style={S.tdR}>{rpShort(ad.gmv)}</td>
                    <td style={{ ...S.tdR, fontWeight: 700, color: ad.roas >= effTarget ? C.good : C.bad }}>{ad.roas.toFixed(2)}</td>
                    <td style={{ ...S.tdR, color: m != null ? (ad.acos > m ? C.bad : C.good) : C.ink }}>{ad.acos.toFixed(1)}%</td>
                    <td style={{ ...S.tdR, color: ad.ctr < thresholds.minCtr ? C.watch : C.ink }}>{ad.ctr.toFixed(2)}%</td>
                    <td style={{ ...S.tdR, color: ad.cvr < thresholds.minCvr ? C.watch : C.ink }}>{ad.cvr.toFixed(2)}%</td>
                    <td style={S.tdR}>{ad.conv}</td>
                    <td style={S.td}><Dot level={lvl} /></td>
                  </tr>
                  {isOpen && (
                    <tr style={S.diagRow}>
                      <td></td>
                      <td colSpan={9} style={S.diagCell}>
                        <div style={S.diagLabel}>DIAGNOSA · target ROAS {effTarget.toFixed(1)}x · CTR ≥{thresholds.minCtr}% · CVR ≥{thresholds.minCvr}%</div>
                        {dx.length === 0 && <div style={S.muted}>Tidak ada sinyal  -  set margin di tab COGS untuk diagnosa profitabilitas.</div>}
                        {dx.map((d, j) => (
                          <div key={j} style={{ ...S.diagItem, borderLeftColor: d.level === "bad" ? C.bad : d.level === "good" ? C.good : C.watch }}>
                            <span style={{ ...S.diagTag, color: d.level === "bad" ? C.bad : d.level === "good" ? C.good : C.watch }}>{d.tag}</span>
                            <span style={S.diagMsg}>{d.msg}</span>
                          </div>
                        ))}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </>}
    </div>
  );
}

function ThresholdPanel({ thresholds, saveThresholds }) {
  const set = (k, v) => saveThresholds({ ...thresholds, [k]: v });
  const numField = (k, label, suffix, step = 1) => (
    <div style={S.thField}>
      <label style={S.thLabel}>{label}</label>
      <div style={S.thInputWrap}>
        <input type="number" step={step} value={thresholds[k]}
          onChange={(e) => set(k, parseFloat(e.target.value) || 0)} style={S.thInput} />
        {suffix && <span style={S.thSuffix}>{suffix}</span>}
      </div>
    </div>
  );
  return (
    <div style={S.thPanel}>
      <div style={S.thHead}>
        <span style={S.thHeadLabel}>BENCHMARK THRESHOLD</span>
        <label style={S.thToggle}>
          <input type="checkbox" checked={thresholds.autoRoasFromMargin}
            onChange={(e) => set("autoRoasFromMargin", e.target.checked)} />
          target ROAS auto dari margin (pakai yang lebih tinggi)
        </label>
      </div>
      <div style={S.thFields}>
        {numField("targetRoas", "Target ROAS", "x", 0.5)}
        {numField("minCtr", "Min CTR", "%", 0.1)}
        {numField("minCvr", "Min CVR", "%", 0.1)}
        {numField("maxCpa", "Max CPA", "Rp", 5000)}
      </div>
    </div>
  );
}

function Dot({ level }) {
  const c = level === "bad" ? C.bad : level === "good" ? C.good : C.watch;
  const txt = level === "bad" ? "perlu aksi" : level === "good" ? "sehat" : "pantau";
  return <span style={{ display: "inline-flex", alignItems: "center", gap: 7, fontFamily: mono, fontSize: 11.5, color: c }}>
    <span style={{ width: 8, height: 8, borderRadius: 8, background: c }} />{txt}
  </span>;
}

/* ---------- Strategy ------------------------------------------------------ */
function Strategy({ active, marginFor, thresholds, activeProducts, cogsItems, shopeeSnap, tiktokSnap, tiktokAdsSnap, metaSnap }) {

  // ── 0. Data completeness ──
  const missing = [];
  if (!activeProducts) missing.push({ icon: "📦", label: "Performa Produk Shopee", hint: "untuk analisa stok & dead stock" });
  if (!shopeeSnap && !tiktokSnap) missing.push({ icon: "💸", label: "Penghasilan Shopee / TikTok", hint: "untuk fee rate & data retur" });
  if (!cogsItems?.length) missing.push({ icon: "🧮", label: "COGS / HPP", hint: "untuk break-even yang akurat" });
  if (!tiktokAdsSnap) missing.push({ icon: "⬛", label: "TikTok GMV Max", hint: "untuk analisa ROI TikTok" });

  const noMargin = active.ads.every((ad) => marginFor(ad) == null);

  // ── 1. Ad buckets (existing logic) ──
  const buckets = useMemo(() => {
    const b = { scale: [], hold: [], material: [], landing: [], healthy: [] };
    active.ads.forEach((ad) => {
      const m = marginFor(ad);
      const dx = diagnose(ad, m, thresholds);
      if (dx.some((d) => d.tag === "SCALE")) b.scale.push(ad);
      else if (dx.some((d) => d.tag === "DI BAWAH TARGET" || d.tag === "MAKAN MARGIN")) b.hold.push(ad);
      else if (dx.some((d) => d.tag === "CTR RENDAH")) b.material.push(ad);
      else if (dx.some((d) => d.tag === "CVR RENDAH")) b.landing.push(ad);
      else b.healthy.push(ad);
    });
    return b;
  }, [active, marginFor, thresholds]);

  // ── 2. Cross-tab health signals + BCG ──
  const bcgData = useMemo(() => {
    if (!activeProducts?.products?.length) return null;
    const adMap = {};
    active.ads.forEach(ad => { if (ad.code) adMap[ad.code] = ad; });
    const feeRate = shopeeSnap?.summary?.feeRateNetGmv || tiktokSnap?.summary?.feeRateNetGmv || 0;
    return classifyBCG(activeProducts.products, adMap, feeRate);
  }, [activeProducts, active, shopeeSnap, tiktokSnap]);
  const signals = useMemo(() => {
    const items = [];

    // Ads health
    const totalAds = active.ads.length;
    const badAds = buckets.hold.length + buckets.material.length + buckets.landing.length;
    const scaleAds = buckets.scale.length;
    const adHealth = totalAds ? ((totalAds - badAds) / totalAds * 100) : 0;
    items.push({
      area: "Performa Iklan",
      score: adHealth,
      good: `${scaleAds} iklan siap scale · ${buckets.healthy.length} healthy`,
      bad: badAds > 0 ? `${badAds} iklan perlu perhatian (hold/fix)` : null,
      icon: "📢"
    });

    // BCG health
    if (bcgData?.length > 0) {
      const stars    = bcgData.filter(p => p.quadrant === "star").length;
      const dogs     = bcgData.filter(p => p.quadrant === "dog").length;
      const cows     = bcgData.filter(p => p.quadrant === "cashcow").length;
      const question = bcgData.filter(p => p.quadrant === "question").length;
      const bcgScore = Math.max(0, Math.min(100, (stars * 3 + cows * 2 + question * 1) / Math.max(bcgData.length,1) * 25));
      items.push({
        area: "BCG Produk",
        score: bcgScore,
        good: stars > 0 ? `${stars} Star · ${cows} Cash Cow · ${question} Question Mark` : null,
        bad: dogs > bcgData.length * 0.4 ? `${dogs} Dog (${(dogs/bcgData.length*100).toFixed(0)}%) — terlalu dominan` : null,
        icon: "⭐"
      });
    }
    const sh = shopeeSnap?.summary;
    const tt = tiktokSnap?.summary;
    const feeRate = sh?.feeRateNetGmv || tt?.feeRateNetGmv || 0;
    const refundRate = sh?.refundRateValue || tt?.refundRateValue || 0;
    const feeScore = feeRate > 0 ? Math.max(0, 100 - (feeRate - 15) * 5) : null;
    if (feeScore != null) items.push({
      area: "Fee Marketplace",
      score: feeScore,
      good: feeRate < 20 ? `Fee rate ${feeRate.toFixed(1)}% — efisien` : null,
      bad: feeRate >= 25 ? `Fee rate ${feeRate.toFixed(1)}% — perlu dicek pricing` : refundRate > 3 ? `Retur ${refundRate.toFixed(1)}% — di atas normal` : null,
      icon: "💸"
    });

    // Inventory health
    const prods = activeProducts?.products || [];
    if (prods.length > 0) {
      const lowStock  = prods.filter(p => p.stock != null && p.stock < 10).length;
      const deadStock = prods.filter(p => p.stock != null && p.stock > 60 && p.roas < (thresholds.targetRoas || 5)).length;
      const invScore  = Math.max(0, 100 - (lowStock * 5) - (deadStock * 3));
      items.push({
        area: "Inventory",
        score: invScore,
        good: lowStock === 0 && deadStock === 0 ? "Stok semua produk aman" : null,
        bad: lowStock > 0 ? `${lowStock} produk stok kritis (<10)` : deadStock > 0 ? `${deadStock} produk dead stock (stok tinggi, ROAS rendah)` : null,
        icon: "📦"
      });
    }

    // COGS/margin health
    if (cogsItems?.length > 0) {
      const withMargin = active.ads.filter(ad => marginFor(ad) != null).length;
      const cogsScore  = totalAds ? (withMargin / totalAds * 100) : 0;
      items.push({
        area: "COGS / Margin",
        score: cogsScore,
        good: cogsScore >= 80 ? `${withMargin}/${totalAds} iklan terhubung ke margin` : null,
        bad: cogsScore < 50 ? `Hanya ${withMargin}/${totalAds} iklan punya data margin — break-even tidak akurat` : null,
        icon: "🧮"
      });
    }

    // TikTok GMV Max
    if (tiktokAdsSnap) {
      const roi = tiktokAdsSnap.summary.blendedRoi;
      const roiScore = Math.min(100, roi / 8 * 100);
      items.push({
        area: "TikTok GMV Max",
        score: roiScore,
        good: roi >= 5 ? `Blended ROI ${roi.toFixed(1)}x — healthy` : null,
        bad: roi < 3 ? `ROI ${roi.toFixed(1)}x — di bawah threshold efisien` : null,
        icon: "⬛"
      });
    }

    return items;
  }, [buckets, bcgData, active, marginFor, thresholds, activeProducts, cogsItems, shopeeSnap, tiktokSnap, tiktokAdsSnap]);

  // ── 3. Prioritized action items ──
  const actions = useMemo(() => {
    const list = [];

    if (buckets.scale.length > 0)
      list.push({ priority: 1, level: "good", tag: "SCALE SEKARANG", desc: `${buckets.scale.length} iklan ROAS jauh di atas target — naikkan budget 20–30% atau longgarkan target ROAS GMV Max. Jangan ubah materi yang menang.`, ads: buckets.scale.slice(0,3) });

    if (buckets.hold.length > 0)
      list.push({ priority: 2, level: "bad", tag: "STOP / KETATKAN", desc: `${buckets.hold.length} iklan ROAS di bawah break-even — pause atau ketatkan target ROAS. Jangan tambah budget untuk volume di iklan rugi.`, ads: buckets.hold.slice(0,3) });

    const sh = shopeeSnap?.summary;
    if (sh?.refundRateValue > 3)
      list.push({ priority: 3, level: "bad", tag: "AUDIT RETUR", desc: `Retur ${sh.refundRateValue.toFixed(1)}% dari GMV — di atas normal fashion 2–3%. Identifikasi SKU paling sering retur: sizing? foto tidak akurat? kualitas packaging?`, ads: [] });

    if (buckets.material.length > 0)
      list.push({ priority: 4, level: "watch", tag: "FIX MATERI IKLAN", desc: `${buckets.material.length} iklan CTR rendah — ganti thumbnail/hook. Uji 1 variasi baru per iklan sebelum scale.`, ads: buckets.material.slice(0,3) });

    if (buckets.landing.length > 0)
      list.push({ priority: 5, level: "watch", tag: "FIX HALAMAN PRODUK", desc: `${buckets.landing.length} iklan CVR rendah — bukan masalah iklan. Cek harga, foto, deskripsi, review. Jangan scale sampai CVR membaik.`, ads: buckets.landing.slice(0,3) });

    const lowStock = (activeProducts?.products || []).filter(p => p.stock != null && p.stock < 10);
    if (lowStock.length > 0)
      list.push({ priority: 6, level: "watch", tag: "RESTOCK SEGERA", desc: `${lowStock.length} produk stok <10 unit — risiko kehabisan saat iklan jalan. Bisa ganggu algoritma GMV Max.`, ads: lowStock.slice(0,3).map(p => ({ name: p.name, roas: p.roas })) });

    if (noMargin && active.ads.length > 0)
      list.push({ priority: 7, level: "watch", tag: "INPUT COGS", desc: "Belum ada data margin — break-even ROAS tidak bisa dihitung akurat. Isi template COGS di tab COGS/Margin untuk unlock diagnosa profitabilitas.", ads: [] });

    // BCG-based actions
    if (bcgData?.length > 0) {
      const stars   = bcgData.filter(p => p.quadrant === "star");
      const dogs    = bcgData.filter(p => p.quadrant === "dog");
      const cows    = bcgData.filter(p => p.quadrant === "cashcow");
      const qs      = bcgData.filter(p => p.quadrant === "question");

      if (stars.length > 0)
        list.push({ priority: 2.5, level: "good", tag: "SCALE STAR PRODUCTS", desc: `${stars.length} produk Star (share & CVR tinggi) — prioritaskan stok dan budget iklan di sini. Ini mesin revenue utama kamu.`, ads: stars.slice(0,3).map(p => ({ name: p.name, roas: p.roas })) });

      if (qs.length > 0)
        list.push({ priority: 4.5, level: "watch", tag: "KEMBANGKAN QUESTION MARK", desc: `${qs.length} produk CVR tinggi tapi share kecil — potensi naik jadi Star. Coba naikkan exposure lewat iklan atau voucher targeted.`, ads: qs.slice(0,3).map(p => ({ name: p.name, roas: p.roas })) });

      if (cows.length > 0 && cows.some(p => p.blendedCvr < 1.5))
        list.push({ priority: 5.5, level: "watch", tag: "FIX CVR CASH COW", desc: `${cows.filter(p=>p.blendedCvr<1.5).length} Cash Cow dengan CVR rendah — share besar tapi konversi bocor. Fix foto, harga, atau deskripsi untuk unlock revenue yang tertinggal.`, ads: cows.filter(p=>p.blendedCvr<1.5).slice(0,3).map(p => ({ name: p.name, roas: p.roas })) });

      if (dogs.length > dogs.length * 0 && dogs.length >= 3)
        list.push({ priority: 8, level: "bad", tag: "AUDIT DOG PRODUCTS", desc: `${dogs.length} produk Dog (share & CVR rendah) — evaluasi: apakah masih worth di-maintain? Pertimbangkan stop iklan, clearance, atau discontinue untuk free up modal kerja.`, ads: dogs.slice(0,3).map(p => ({ name: p.name, roas: p.roas })) });
    }

    // Marketplace-specific recommendations
    const shSum = shopeeSnap?.summary;
    const ttSum = tiktokSnap?.summary;
    const feeRate = shSum?.feeRateNetGmv || ttSum?.feeRateNetGmv || 0;
    if (feeRate > 22)
      list.push({ priority: 6.5, level: "watch", tag: "EFISIENSI FEE MARKETPLACE", desc: `Fee rate ${feeRate.toFixed(1)}% dari Net GMV — di atas rata-rata fashion. Cek: (1) proporsi voucher seller terlalu besar? (2) banyak transaksi via SPayLater yang fee-nya lebih tinggi? (3) ada iklan yang conversion-nya rendah tapi tetap jalan?`, ads: [] });

    const spaylater = shopeeSnap?.orders?.filter(o => o.metodePembayaran === "SPayLater").length || 0;
    const totalOrders = shopeeSnap?.orders?.length || 0;
    if (totalOrders > 0 && spaylater / totalOrders > 0.3)
      list.push({ priority: 9, level: "watch", tag: "DOMINASI SPAYLATER", desc: `${(spaylater/totalOrders*100).toFixed(0)}% pembeli pakai SPayLater — bagus untuk konversi, tapi pastikan cash flow tidak terganggu karena pelepasan dana yang lebih lambat.`, ads: [] });

    return list.sort((a,b) => a.priority - b.priority);
  }, [buckets, bcgData, shopeeSnap, tiktokSnap, activeProducts, noMargin, active, thresholds]);

  const scoreColor = s => s >= 75 ? C.good : s >= 50 ? C.watch : C.bad;

  return (
    <div>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 4 }}>
        <SectionLabel>RINGKASAN STRATEGI · LINTAS TAB</SectionLabel>
        <a href="https://wa.me/6282130311844" target="_blank" rel="noreferrer"
          style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: C.accent, textDecoration: "none", padding: "6px 12px", border: `1px solid ${C.accent}`, borderRadius: 6, whiteSpace: "nowrap" }}>
          Konsultasi →
        </a>
      </div>
      <div style={{ fontFamily: sans, fontSize: 12.5, color: C.muted, lineHeight: 1.6, marginBottom: 16 }}>
        Rekapan kondisi dari semua tab — iklan, fee, inventory, COGS. Diagnosa dan saran adalah gambaran awal berbasis data; validasi dengan konteks operasional sebelum eksekusi.
      </div>

      {/* Missing data notice */}
      {missing.length > 0 && (
        <div style={{ background: C.watch + "12", border: `1px solid ${C.watch}44`, borderRadius: 8, padding: "12px 16px", marginBottom: 16 }}>
          <div style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: C.watch, letterSpacing: 1.5, marginBottom: 8 }}>⚠ STRATEGI BELUM LENGKAP · {missing.length} DATA BELUM DIIMPORT</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.ink, marginBottom: 8 }}>Import data berikut untuk rekomendasi yang lebih tajam:</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {missing.map((m, i) => (
              <div key={i} style={{ fontFamily: sans, fontSize: 11, color: C.watch, background: C.watch+"18", border: `1px solid ${C.watch}44`, borderRadius: 6, padding: "4px 10px" }}>
                {m.icon} <b>{m.label}</b> — {m.hint}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Health scorecard */}
      {signals.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: C.muted, letterSpacing: 1.5, marginBottom: 10 }}>HEALTH CHECK · SEMUA AREA</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 10 }}>
            {signals.map((sig, i) => (
              <div key={i} style={{ background: C.panel, border: `1px solid ${C.line}`, borderLeft: `3px solid ${scoreColor(sig.score)}`, borderRadius: "0 8px 8px 0", padding: "12px 14px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontFamily: sans, fontSize: 12, fontWeight: 700, color: C.ink }}>{sig.icon} {sig.area}</span>
                  <span style={{ fontFamily: mono, fontSize: 13, fontWeight: 700, color: scoreColor(sig.score) }}>{sig.score.toFixed(0)}</span>
                </div>
                {/* Score bar */}
                <div style={{ height: 4, background: C.panel2, borderRadius: 2, marginBottom: 8 }}>
                  <div style={{ width: sig.score + "%", height: "100%", background: scoreColor(sig.score), borderRadius: 2, transition: "width 0.3s" }} />
                </div>
                {sig.good && <div style={{ fontFamily: sans, fontSize: 11, color: C.good, lineHeight: 1.4 }}>✓ {sig.good}</div>}
                {sig.bad  && <div style={{ fontFamily: sans, fontSize: 11, color: C.bad,  lineHeight: 1.4 }}>⚠ {sig.bad}</div>}
                {!sig.good && !sig.bad && <div style={{ fontFamily: sans, fontSize: 11, color: C.muted }}>Data tersedia</div>}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Prioritized action plan */}
      <div style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: C.muted, letterSpacing: 1.5, marginBottom: 10 }}>PRIORITAS AKSI · URUT DARI PALING MENDESAK</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 20 }}>
        {actions.map((a, i) => (
          <div key={i} style={{ background: C.panel, border: `1px solid ${C.line}`, borderLeft: `3px solid ${a.level === "good" ? C.good : a.level === "bad" ? C.bad : C.watch}`, borderRadius: "0 8px 8px 0", padding: "12px 16px", display: "flex", gap: 12, alignItems: "flex-start" }}>
            <div style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: C.panel, background: a.level === "good" ? C.good : a.level === "bad" ? C.bad : C.watch, borderRadius: 4, padding: "2px 7px", flexShrink: 0, marginTop: 1 }}>#{a.priority}</div>
            <div style={{ flex: 1 }}>
              <div style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: a.level === "good" ? C.good : a.level === "bad" ? C.bad : C.watch, marginBottom: 4 }}>{a.tag}</div>
              <div style={{ fontFamily: sans, fontSize: 12.5, color: C.ink, lineHeight: 1.6, marginBottom: a.ads.length > 0 ? 8 : 0 }}>{a.desc}</div>
              {a.ads.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {a.ads.map((ad, j) => (
                    <span key={j} style={{ fontFamily: mono, fontSize: 10, color: C.muted, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 4, padding: "2px 8px" }}>
                      {ad.name?.length > 28 ? ad.name.slice(0,28)+"…" : ad.name}
                      {ad.roas != null && <span style={{ color: ad.roas >= (thresholds.targetRoas||5) ? C.good : C.bad }}> · {ad.roas?.toFixed(1)}x</span>}
                    </span>
                  ))}
                </div>
              )}
            </div>
          </div>
        ))}
        {actions.length === 0 && (
          <div style={{ fontFamily: sans, fontSize: 13, color: C.muted, padding: "20px 0", textAlign: "center" }}>Semua area terlihat sehat — tidak ada aksi mendesak saat ini.</div>
        )}
      </div>

      {/* Ad buckets (simplified) */}
      <div style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: C.muted, letterSpacing: 1.5, marginBottom: 10 }}>DISTRIBUSI IKLAN · {active.ads.length} TOTAL</div>
      <div style={S.stratGrid}>
        <StratCard color={C.good}  title="SCALE"        when="ROAS jauh di atas target & spend signifikan"          action="Naikkan budget 20–30%/langkah atau longgarkan target ROAS GMV Max." ads={buckets.scale} />
        <StratCard color={C.bad}   title="HOLD / CUT"   when="ROAS di bawah break-even atau ACOS > margin"          action="Ketatkan target ROAS atau pause. Jangan scale iklan rugi." ads={buckets.hold} />
        <StratCard color={C.watch} title="FIX MATERI"   when="CTR rendah, impression cukup"                         action="Ganti visual/hook. Uji 1 variasi baru sebelum dorong budget." ads={buckets.material} />
        <StratCard color={C.watch} title="FIX LANDING"  when="CTR oke tapi CVR rendah"                              action="Cek harga, stok, foto, review. Jangan scale sebelum CVR membaik." ads={buckets.landing} />
        <StratCard color={C.ink}   title="HEALTHY"      when="Profit, tidak ada sinyal merah"                       action="Maintain. Pantau drift di snapshot harian." ads={buckets.healthy} />
      </div>
    </div>
  );
}
function StratCard({ color, title, when, action, ads }) {
  return (
    <div style={{ ...S.stratCard, borderTopColor: color }}>
      <div style={{ ...S.stratTitle, color }}>{title} · {ads.length}</div>
      <div style={S.stratWhen}><b>Kapan:</b> {when}</div>
      <div style={S.stratAction}>{action}</div>
      {ads.length > 0 && (
        <div style={S.stratAds}>
          {ads.slice(0, 6).map((a, i) => <div key={i} style={S.stratAdRow}><span>{a.name}</span><span style={S.muted}>ROAS {a.roas.toFixed(1)}</span></div>)}
          {ads.length > 6 && <div style={S.muted}>+{ads.length - 6} lagi</div>}
        </div>
      )}
    </div>
  );
}

/* ---------- COGS tab ------------------------------------------------------ */
function CogsTab({ active, cogsMap, setCogsMap, keyFor, flash, cogsItems, onImportCogs, onDownloadTemplate }) {
  const hasCogs = cogsItems && cogsItems.length > 0;

  // Connects to: Unit Economics, Contribution Margin, Simulasi L/R
  const connectedTabs = [
    { label: "Unit Economics", desc: "HPP per unit → gross & net margin per SKU" },
    { label: "Contribution Margin", desc: "CM total & % per SKU setelah fee + ads" },
    { label: "Simulasi L/R", desc: "HPP dipakai untuk estimasi laba kotor" },
  ];

  // Margin % legacy (still used by action plan)
  const products = useMemo(() => {
    if (!active) return [];
    const seen = new Map();
    active.ads.forEach((ad) => {
      const k = keyFor(ad);
      if (!seen.has(k)) seen.set(k, { key: k, label: ad.name, code: ad.code, gmv: 0 });
      seen.get(k).gmv += ad.gmv;
    });
    return [...seen.values()].sort((a, b) => b.gmv - a.gmv);
  }, [active, keyFor]);

  async function setMargin(key, label, val) {
    const marginPct = val === "" ? null : Math.max(0, Math.min(100, parseFloat(val) || 0));
    if (marginPct == null) {
      await idbDel("cogs", key);
      const m = { ...cogsMap }; delete m[key]; setCogsMap(m);
    } else {
      const rec = { key, label, marginPct };
      await idbPut("cogs", rec);
      setCogsMap({ ...cogsMap, [key]: rec });
    }
  }

  const [bulk, setBulk] = useState("");
  async function applyBulk() {
    const v = parseFloat(bulk);
    if (isNaN(v)) return;
    const m = { ...cogsMap };
    for (const p of products) {
      const rec = { key: p.key, label: p.label, marginPct: Math.max(0, Math.min(100, v)) };
      await idbPut("cogs", rec);
      m[p.key] = rec;
    }
    setCogsMap(m);
    flash(`Margin ${v}% diterapkan ke ${products.length} produk`);
  }

  return (
    <div>
      <SectionLabel>COGS / HPP — DATA MANAGEMENT</SectionLabel>

      {/* Connected tabs info */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 10, marginBottom: 20 }}>
        {connectedTabs.map((t, i) => (
          <div key={i} style={{ background: C.accent+"08", border: `1px solid ${C.accent}22`, borderLeft: `3px solid ${C.accent}`, borderRadius: 8, padding: "12px 14px" }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.accent, marginBottom: 3 }}>→ {t.label}</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.muted }}>{t.desc}</div>
          </div>
        ))}
      </div>

      {/* COGS Template section */}
      <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: "18px 20px", marginBottom: 20, boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
        <div style={{ fontFamily: sans, fontSize: 13.5, fontWeight: 700, color: C.ink, marginBottom: 6 }}>
          Template HPP per SKU {hasCogs && <span style={{ fontFamily: sans, fontSize: 11, fontWeight: 700, color: C.good, background: C.good+"12", border: `1px solid ${C.good}33`, padding: "2px 8px", borderRadius: 10, marginLeft: 8 }}>✓ {cogsItems.length} SKU diimport</span>}
        </div>
        <div style={{ fontFamily: sans, fontSize: 13, color: C.muted, lineHeight: 1.6, marginBottom: 14 }}>
          Download template Excel, isi kolom <b>HPP Produksi</b>, <b>HPP Packaging</b>, dan <b>Harga Jual</b> per SKU, lalu import balik. Data ini dipakai otomatis di Unit Economics, Contribution Margin, dan Simulasi L/R.
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button style={S.importBtnBig} onClick={onDownloadTemplate}>
            ⬇ Download Template COGS
          </button>
          <button style={S.importBtn2Big} onClick={onImportCogs}>
            + Import COGS yang sudah diisi
          </button>
        </div>
        {hasCogs && (
          <div style={{ marginTop: 16, borderTop: `1px solid ${C.line}`, paddingTop: 14 }}>
            <div style={{ fontFamily: sans, fontSize: 11, fontWeight: 700, color: C.muted, letterSpacing: 0.5, textTransform: "uppercase", marginBottom: 10 }}>Preview HPP yang sudah diimport</div>
            <div style={S.tableWrap}>
              <table style={S.table}>
                <thead><tr>
                  <th style={S.th}>SKU</th>
                  <th style={S.th}>Nama Internal</th>
                  <th style={{ ...S.th, textAlign: "right" }}>HPP Produksi</th>
                  <th style={{ ...S.th, textAlign: "right" }}>HPP Packaging</th>
                  <th style={{ ...S.th, textAlign: "right" }}>Total HPP</th>
                  <th style={{ ...S.th, textAlign: "right" }}>Harga Jual</th>
                  <th style={{ ...S.th, textAlign: "right" }}>Gross Margin</th>
                </tr></thead>
                <tbody>
                  {cogsItems.map((c, i) => {
                    const gm = c.harga > 0 ? ((c.harga - c.total) / c.harga * 100) : null;
                    return (
                      <tr key={i} style={S.tr}>
                        <td style={S.td}><span style={{ fontFamily: sans, fontSize: 12, color: C.muted }}>{c.sku}</span></td>
                        <td style={S.td}><div style={S.adName}>{c.nama || c.sku}</div></td>
                        <td style={S.tdR}>{c.hpp_prod > 0 ? rp(c.hpp_prod) : "—"}</td>
                        <td style={S.tdR}>{c.hpp_pack > 0 ? rp(c.hpp_pack) : "—"}</td>
                        <td style={{ ...S.tdR, fontWeight: 700 }}>{rp(c.total)}</td>
                        <td style={S.tdR}>{c.harga > 0 ? rp(c.harga) : "—"}</td>
                        <td style={{ ...S.tdR, fontWeight: 700, color: gm ? (gm >= 40 ? C.good : gm >= 20 ? C.watch : C.bad) : C.muted }}>
                          {gm != null ? gm.toFixed(1)+"%" : "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* Legacy margin % — for action plan compatibility */}
      {active && products.length > 0 && (
        <>
          <SectionLabel>MARGIN % PER IKLAN — untuk Action Plan</SectionLabel>
          <div style={S.cogsHelp}>
            Isi gross margin % per kampanye iklan untuk kalkulasi break-even ROAS dan Action Plan di tab Overview.
            Berbeda dengan HPP template di atas — ini input cepat berbasis persentase, bukan rupiah per SKU.
          </div>
          <div style={S.bulkRow}>
            <input placeholder="margin % untuk semua" value={bulk} onChange={(e) => setBulk(e.target.value)} style={S.search} />
            <button style={S.smallBtn} onClick={applyBulk}>Terapkan ke semua</button>
          </div>
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr>
                <th style={S.th}>Produk / Iklan</th>
                <th style={S.th}>Kode</th>
                <th style={{ ...S.th, textAlign: "right" }}>GMV</th>
                <th style={{ ...S.th, textAlign: "right" }}>Margin %</th>
                <th style={{ ...S.th, textAlign: "right" }}>BE ROAS</th>
              </tr></thead>
              <tbody>
                {products.map((p) => {
                  const m = cogsMap[p.key]?.marginPct;
                  return (
                    <tr key={p.key} style={S.tr}>
                      <td style={S.td}><div style={S.adName}>{p.label}</div></td>
                      <td style={S.td}><span style={S.muted}>{p.code && p.code !== "-" ? p.code : "—"}</span></td>
                      <td style={S.tdR}>{rpShort(p.gmv)}</td>
                      <td style={S.tdR}>
                        <input type="number" defaultValue={m ?? ""} placeholder="—"
                          onBlur={(e) => setMargin(p.key, p.label, e.target.value)}
                          style={S.marginInput} />
                      </td>
                      <td style={{ ...S.tdR, color: m ? C.good : C.muted }}>{m ? (100/m).toFixed(1) : "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

/* ---------- Product tab (BCG share × CVR) -------------------------------- */
function ProductTab({ snap, thresholds, active, feeRate }) {
  const adMap = useMemo(() => {
    const m = {};
    if (active) active.ads.forEach(ad => { if (ad.code && ad.code !== "-") m[ad.code] = ad; });
    return m;
  }, [active]);
  const classified = useMemo(() => classifyBCG(snap.products, adMap, feeRate || 0), [snap, adMap, feeRate]);
  const [openCode, setOpenCode] = useState(null);

  const adByCode = adMap; // reuse from classifyBCG

  const counts = classified.reduce((a, p) => (a[p.quadrant]++, a), { star: 0, cashcow: 0, question: 0, dog: 0 });

  return (
    <div>
      <SectionLabel>BCG MATRIX · {snap.products.length} SKU · {snap.periodStart} → {snap.periodEnd}</SectionLabel>
      <div style={S.bcgIntro}>
        Sumbu X = net revenue share{feeRate > 0 ? ` (fee-adjusted ${feeRate.toFixed(0)}%)` : " (gross)"}, sumbu Y = blended CVR{active ? " (produk + iklan)" : " (produk)"}.
        {" "}Star = scale; Cash Cow = fix konversi; Question Mark = kandidat iklan; Dog = evaluasi.
        {feeRate > 0 && <span style={{ color: C.good, marginLeft: 8 }}>✓ income file terhubung</span>}
        {active && <span style={{ color: C.good, marginLeft: 8 }}>✓ data iklan terhubung</span>}
      </div>
      <BCGScatter classified={classified} onPick={(c) => setOpenCode(c === openCode ? null : c)} openCode={openCode} />
      <div style={S.quadLegend}>
        {Object.entries(QUAD).map(([k, q]) => (
          <div key={k} style={S.quadLegendItem}>
            <span style={{ ...S.quadDot, background: q.color }} />
            <b style={{ color: q.color }}>{q.label}</b>
            <span style={S.muted}>{counts[k]} · {q.desc}</span>
          </div>
        ))}
      </div>
      <SectionLabel>DIAGNOSA &amp; SARAN  -  PER SKU</SectionLabel>
      <div>
        {classified.map((p) => {
          const q = QUAD[p.quadrant];
          const dx = diagnoseProduct(p, thresholds);
          const isOpen = openCode === p.code;
          const ad = adByCode[p.code];
          const borderCol = q.color;
          return (
            <div key={p.code} id={"sku-" + p.code} style={{ ...S.skuCard, borderLeftColor: borderCol }}>
              <div style={S.skuHead} onClick={() => setOpenCode(isOpen ? null : p.code)}>
                <div style={S.skuHeadL}>
                  <QuadTag q={p.quadrant} />
                  <span style={S.skuName}>{p.name}</span>
                </div>
                <div style={S.skuHeadR}>
                  <span style={S.skuStat}>share <b>{p.share.toFixed(1)}%</b></span>
                  <span style={S.skuStat}>CVR <b>{p.blendedCvr ? p.blendedCvr.toFixed(2) : p.cvr.toFixed(2)}%</b></span>
                  <span style={S.skuStat}>{rpShort(p.sales)}</span>
                  {p.roas != null && <span style={{ ...S.skuStat, color: p.roas >= 5 ? C.good : p.roas >= 3 ? C.watch : C.bad }}>ROAS {p.roas.toFixed(1)}</span>}
                  {p.adGmv > 0 && <span style={{ ...S.skuStat, color: C.muted }}>iklan {rpShort(p.adGmv)}</span>}
                  <span style={{ color: C.muted, fontFamily: mono }}>{isOpen ? "▾" : "▸"}</span>
                </div>
              </div>
              {isOpen && (
                <div style={S.skuBody}>

                  {dx.map((d, i) => (
                    <div key={i} style={{ ...S.skuDiag, borderLeftColor: d.level === "bad" ? C.bad : d.level === "good" ? C.good : C.watch }}>
                      <div style={{ ...S.skuDiagTitle, color: d.level === "bad" ? C.bad : d.level === "good" ? C.good : C.watch }}>{d.title}</div>
                      <div style={S.skuDiagMsg}>{d.msg}</div>
                    </div>
                  ))}
                  {ad && <CrossSignal p={p} ad={ad} thresholds={thresholds} />}
                  {!ad && p.quadrant === "question" && (
                    <div style={{ ...S.skuDiag, borderLeftColor: C.accent }}>
                      <div style={{ ...S.skuDiagTitle, color: C.accent }}>Silang iklan × produk</div>
                      <div style={S.skuDiagMsg}>CVR terbukti bagus tapi belum ada iklan produk sendiri. Kandidat kuat untuk dibuatkan iklan.</div>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}


// cross-cutting diagnosis combining the ad row and the product row for same code
function CrossSignal({ p, ad, thresholds }) {
  const msgs = [];
  if (p.quadrant === "star" && ad.roas >= thresholds.targetRoas) {
    msgs.push({ level: "good", t: "Star + iklan profitable", m: `Produk Star dengan iklan ROAS ${ad.roas.toFixed(1)}x di atas target. Scale aman  -  naikkan budget bertahap.` });
  }
  if (p.quadrant === "cashcow" && ad.roas >= thresholds.targetRoas) {
    msgs.push({ level: "watch", t: "Iklan oke, konversi produk yang rem", m: `Iklan ROAS ${ad.roas.toFixed(1)}x sebenarnya bagus, tapi CVR produk rendah. Jangan dorong budget lebih  -  perbaiki halaman produk dulu, baru scale. Tambah spend sekarang = bayar mahal untuk traffic yang bocor.` });
  }
  if (ad.roas < thresholds.targetRoas && (p.quadrant === "star" || p.quadrant === "cashcow")) {
    msgs.push({ level: "watch", t: "Produk laku, iklan rugi", m: `Produk ini share-nya tinggi tapi iklannya ROAS ${ad.roas.toFixed(1)}x di bawah target. Sebagian besar penjualan kemungkinan organik  -  iklan mungkin cuma kanibalisasi. Coba turunkan/pause iklan, lihat apakah penjualan total turun.` });
  }
  if (p.quadrant === "dog" && ad.spend > 0) {
    msgs.push({ level: "bad", t: "Buang spend di Dog", m: `Produk Dog masih dapat spend iklan ${rpShort(ad.spend)}. Hentikan  -  alokasikan ke Star/Question Mark.` });
  }
  if (!msgs.length) return null;
  return (
    <>
      {msgs.map((x, i) => (
        <div key={i} style={{ ...S.skuDiag, borderLeftColor: C.accent, background: C.accent + "0c" }}>
          <div style={{ ...S.skuDiagTitle, color: C.accent }}>⇄ {x.t}</div>
          <div style={S.skuDiagMsg}>{x.m}</div>
        </div>
      ))}
    </>
  );
}

/* ---------- Forecast tab (benchmark-driven planner / allocator) ---------- */
// Builds per-product benchmarks from the ad×product join, then either:
//  - PLANNER: target revenue -> required spend & traffic KPIs (per product)
//  - ALLOCATOR: fixed budget -> projected revenue, split by BCG quadrant priority
function Forecast({ active, activeProducts, thresholds, marginFor }) {
  const [mode, setMode] = useState("planner");
  const [targetRev, setTargetRev] = useState(300000000); // Rp/month
  const [budget, setBudget] = useState(30000000);        // Rp/month
  const [overrides, setOverrides] = useState({});        // code -> {roas,cvr,aov}

  // store-level fallback benchmark
  const storeAgg = useMemo(() => aggregate(active.ads), [active]);
  const storeRoas = storeAgg.roas || 8;

  // build per-product benchmark by joining ad row + product row on code
  const benchmarks = useMemo(() => {
    const adByCode = {};
    active.ads.forEach((ad) => { if (ad.code && ad.code !== "-") adByCode[ad.code] = ad; });
    const classified = activeProducts ? classifyBCG(activeProducts.products) : [];
    const prodByCode = {};
    classified.forEach((p) => { prodByCode[p.code] = p; });

    // universe: prefer product list (has share/quadrant); fallback to ad rows
    let universe;
    if (classified.length) {
      universe = classified.map((p) => {
        const ad = adByCode[p.code];
        return {
          code: p.code, name: p.name, quadrant: p.quadrant, share: p.share,
          roas: ad ? ad.roas : storeRoas, hasAd: !!ad,
          cvr: p.cvr || (ad ? ad.cvr : storeAgg.cvr),
          aov: ad && ad.conv ? ad.gmv / ad.conv : (p.orders ? p.sales / p.orders : storeAgg.cpa ? 0 : 0),
          ctr: ad ? ad.ctr : p.ctr,
          sales: p.sales,
        };
      });
    } else {
      universe = active.ads.filter((a) => a.code && a.code !== "-").map((ad) => ({
        code: ad.code, name: ad.name, quadrant: "star", share: 0,
        roas: ad.roas, hasAd: true, cvr: ad.cvr, aov: ad.conv ? ad.gmv / ad.conv : 0, ctr: ad.ctr, sales: ad.gmv,
      }));
    }
    // apply overrides
    return universe.map((b) => {
      const o = overrides[b.code] || {};
      return { ...b, roas: o.roas ?? b.roas, cvr: o.cvr ?? b.cvr, aov: o.aov ?? b.aov };
    });
  }, [active, activeProducts, overrides, storeRoas, storeAgg]);

  // quadrant priority weights for allocator
  const QW = { star: 1.0, question: 0.8, cashcow: 0.35, dog: 0 };

  // PLANNER: distribute target revenue across products by current sales mix (eligible only),
  // then compute required spend = rev/roas, clicks = conv/cvr, impr = clicks/ctr.
  const planner = useMemo(() => {
    const eligible = benchmarks.filter((b) => b.quadrant !== "dog" && b.roas > 0);
    const mixBase = eligible.reduce((t, b) => t + Math.max(b.sales, 1), 0) || 1;
    const rows = eligible.map((b) => {
      const rev = targetRev * (Math.max(b.sales, 1) / mixBase);
      const spend = b.roas ? rev / b.roas : 0;
      const conv = b.aov ? rev / b.aov : 0;
      const clicks = b.cvr ? conv / (b.cvr / 100) : 0;
      const impr = b.ctr ? clicks / (b.ctr / 100) : 0;
      const feasible = b.quadrant !== "cashcow"; // cashcow flagged: fix CVR before scaling
      return { ...b, rev, spend, conv, clicks, impr, feasible };
    }).sort((a, b) => b.rev - a.rev);
    const tot = rows.reduce((t, r) => ({ rev: t.rev + r.rev, spend: t.spend + r.spend, conv: t.conv + r.conv, impr: t.impr + r.impr }), { rev: 0, spend: 0, conv: 0, impr: 0 });
    return { rows, tot, blendedRoas: tot.spend ? tot.rev / tot.spend : 0 };
  }, [benchmarks, targetRev]);

  // ALLOCATOR: split budget by quadrant-weighted share, project revenue = spend*roas
  const allocator = useMemo(() => {
    const weighted = benchmarks.map((b) => ({ ...b, w: (QW[b.quadrant] ?? 0) * Math.max(b.share || b.sales / 1e6, 0.1) }));
    const totW = weighted.reduce((t, b) => t + b.w, 0) || 1;
    const rows = weighted.filter((b) => b.w > 0).map((b) => {
      const spend = budget * (b.w / totW);
      const rev = spend * b.roas;
      const conv = b.aov ? rev / b.aov : 0;
      return { ...b, spend, rev, conv };
    }).sort((a, b) => b.spend - a.spend);
    const tot = rows.reduce((t, r) => ({ spend: t.spend + r.spend, rev: t.rev + r.rev }), { spend: 0, rev: 0 });
    return { rows, tot, blendedRoas: tot.spend ? tot.rev / tot.spend : 0 };
  }, [benchmarks, budget]);

  const setOv = (code, field, val) => {
    setOverrides((o) => ({ ...o, [code]: { ...o[code], [field]: val === "" ? undefined : parseFloat(val) } }));
  };

  const hasProd = !!activeProducts;

  return (
    <div>
      <div style={S.fcModeRow}>
        <div style={S.fcToggle}>
          <button style={{ ...S.fcToggleBtn, ...(mode === "planner" ? S.fcToggleOn : {}) }} onClick={() => setMode("planner")}>Target → Spend</button>
          <button style={{ ...S.fcToggleBtn, ...(mode === "allocator" ? S.fcToggleOn : {}) }} onClick={() => setMode("allocator")}>Budget → Revenue</button>
        </div>
        {mode === "planner" ? (
          <div style={S.fcInputWrap}>
            <label style={S.fcInputLabel}>Target revenue / bulan</label>
            <div style={S.thInputWrap}><span style={S.thSuffix}>Rp</span>
              <input type="number" step={10000000} value={targetRev} onChange={(e) => setTargetRev(parseFloat(e.target.value) || 0)} style={S.thInput} /></div>
          </div>
        ) : (
          <div style={S.fcInputWrap}>
            <label style={S.fcInputLabel}>Budget iklan / bulan</label>
            <div style={S.thInputWrap}><span style={S.thSuffix}>Rp</span>
              <input type="number" step={5000000} value={budget} onChange={(e) => setBudget(parseFloat(e.target.value) || 0)} style={S.thInput} /></div>
          </div>
        )}
      </div>

      <div style={S.fcDisclaimer}>
        Berbasis benchmark periode sekarang (ROAS, CVR, AOV per produk dari join iklan × produk).
        Ini kalkulator alokasi, <b>bukan prediksi tren waktu</b>  -  itu butuh beberapa bulan snapshot.
        {!hasProd && " Tanpa file produk, alokasi pakai ROAS toko rata-rata sebagai proxy."}
      </div>

      {mode === "planner" ? (
        <>
          <div style={S.fcKpiRow}>
            <FcKpi label="Target revenue" value={rpShort(targetRev)} />
            <FcKpi label="Spend dibutuhkan" value={rpShort(planner.tot.spend)} accent />
            <FcKpi label="Blended ROAS" value={planner.blendedRoas.toFixed(2) + "x"} />
            <FcKpi label="Konversi target" value={nfmt(planner.tot.conv)} />
            <FcKpi label="Impression dibeli" value={nfmt(planner.tot.impr)} />
          </div>
          <SectionLabel>RENCANA PER PRODUK · target → KPI</SectionLabel>
          <FcDisclaimerCashcow rows={planner.rows} />
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr>
                <th style={S.th}>Produk</th>
                <th style={S.th}>Kuadran</th>
                <th style={{ ...S.th, textAlign: "right" }}>Target rev</th>
                <th style={{ ...S.th, textAlign: "right" }}>Spend</th>
                <th style={{ ...S.th, textAlign: "right" }}>ROAS*</th>
                <th style={{ ...S.th, textAlign: "right" }}>CVR*</th>
                <th style={{ ...S.th, textAlign: "right" }}>Klik</th>
                <th style={{ ...S.th, textAlign: "right" }}>Impresi</th>
              </tr></thead>
              <tbody>
                {planner.rows.map((r) => (
                  <tr key={r.code} style={S.tr}>
                    <td style={S.td}><div style={S.adName}>{r.name}</div>{!r.hasAd && <div style={S.adMeta}>ROAS proxy toko</div>}{!r.feasible && <div style={{ ...S.adMeta, color: C.watch }}>fix CVR dulu sebelum dorong spend</div>}</td>
                    <td style={S.td}><QuadTag q={r.quadrant} /></td>
                    <td style={S.tdR}>{rpShort(r.rev)}</td>
                    <td style={{ ...S.tdR, color: C.accent }}>{rpShort(r.spend)}</td>
                    <td style={S.tdR}><OvInput val={overrides[r.code]?.roas} ph={r.roas.toFixed(1)} onSet={(v) => setOv(r.code, "roas", v)} /></td>
                    <td style={S.tdR}><OvInput val={overrides[r.code]?.cvr} ph={r.cvr.toFixed(2)} onSet={(v) => setOv(r.code, "cvr", v)} /></td>
                    <td style={S.tdR}>{nfmt(r.clicks)}</td>
                    <td style={S.tdR}>{nfmt(r.impr)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={S.fcFootnote}>* ROAS &amp; CVR bisa di-override per produk untuk skenario (mis. CVR naik setelah fix foto). Kosongkan = pakai benchmark sekarang.</div>
        </>
      ) : (
        <>
          <div style={S.fcKpiRow}>
            <FcKpi label="Budget" value={rpShort(budget)} />
            <FcKpi label="Proyeksi revenue" value={rpShort(allocator.tot.rev)} accent />
            <FcKpi label="Blended ROAS" value={allocator.blendedRoas.toFixed(2) + "x"} />
            <FcKpi label="Proyeksi konversi" value={nfmt(allocator.tot.conv)} />
          </div>
          <SectionLabel>ALOKASI PER PRODUK · dipandu kuadran BCG</SectionLabel>
          <div style={S.fcAllocNote}>Star ×1.0 · Question Mark ×0.8 · Cash Cow ×0.35 (ditahan, fix CVR dulu) · Dog dikecualikan.</div>
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr>
                <th style={S.th}>Produk</th>
                <th style={S.th}>Kuadran</th>
                <th style={{ ...S.th, textAlign: "right" }}>Alokasi spend</th>
                <th style={{ ...S.th, textAlign: "right" }}>% budget</th>
                <th style={{ ...S.th, textAlign: "right" }}>ROAS*</th>
                <th style={{ ...S.th, textAlign: "right" }}>Proyeksi rev</th>
              </tr></thead>
              <tbody>
                {allocator.rows.map((r) => (
                  <tr key={r.code} style={S.tr}>
                    <td style={S.td}><div style={S.adName}>{r.name}</div>{!r.hasAd && <div style={S.adMeta}>ROAS proxy toko</div>}</td>
                    <td style={S.td}><QuadTag q={r.quadrant} /></td>
                    <td style={{ ...S.tdR, color: C.accent }}>{rpShort(r.spend)}</td>
                    <td style={S.tdR}>{((r.spend / (budget || 1)) * 100).toFixed(1)}%</td>
                    <td style={S.tdR}><OvInput val={overrides[r.code]?.roas} ph={r.roas.toFixed(1)} onSet={(v) => setOv(r.code, "roas", v)} /></td>
                    <td style={S.tdR}>{rpShort(r.rev)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={S.fcFootnote}>* ROAS bisa di-override per produk. Proyeksi = spend × ROAS, dengan asumsi efisiensi bertahan saat budget naik (kenyataannya ROAS biasanya turun saat scale agresif  -  pakai sebagai batas atas).</div>
        </>
      )}
    </div>
  );
}
function FcKpi({ label, value, accent }) {
  return (
    <div style={{ ...S.kpi, ...(accent ? { background: C.panel2, borderColor: C.accent + "44" } : {}) }}>
      <div style={S.kpiLabel}>{label}</div>
      <div style={{ ...S.kpiVal, ...(accent ? { color: C.accent } : {}) }}>{value}</div>
    </div>
  );
}
/* ---------- Inventory & Iklan tab ---------------------------------------- */
// Senior consultant-grade analysis: Stok × Performa Iklan × BCG → keputusan
function InventoryTab({ snap, active, thresholds, stockMap, setStockMap }) {
  const periodDays = useMemo(() => periodDaysFromSnap(snap), [snap]);
  const classified = useMemo(() => classifyBCG(snap.products), [snap]);

  const adByCode = useMemo(() => {
    const m = {};
    if (active) active.ads.forEach((ad) => { if (ad.code && ad.code !== "-") m[ad.code] = ad; });
    return m;
  }, [active]);

  // build enriched SKU rows
  const skus = useMemo(() => {
    return classified.map((p) => {
      const ad = adByCode[p.code];
      const stockUnits = stockMap[p.code] ?? null;
      const dailyRate = p.unitsSold / periodDays;
      const dio = stockUnits != null && dailyRate > 0 ? Math.round(stockUnits / dailyRate) : null;
      const fc = forecastDemand(dailyRate, periodDays, 30);
      const forecastUnits = fc.units;
      const forecastMult = fc.mult;
      const restockQty = stockUnits != null ? Math.max(0, forecastUnits - stockUnits) : null;
      const dioInfo = dioLabel(dio);
      return { ...p, ad, stockUnits, dailyRate, dio, dioInfo, forecastUnits, forecastMult, restockQty };
    });
  }, [classified, adByCode, stockMap, periodDays]);

  const withStock = skus.filter((s) => s.stockUnits != null).length;
  const totalRestockQty = skus.reduce((t, s) => t + (s.restockQty || 0), 0);

  async function setStock(code, val) {
    const units = val === "" ? null : Math.max(0, parseInt(val, 10) || 0);
    const next = { ...stockMap };
    if (units === null) { delete next[code]; await idbDel("stock", code); }
    else { next[code] = units; await idbPut("stock", { code, units }); }
    setStockMap(next);
  }

  return (
    <div>
      <SectionLabel>INVENTORY &amp; IKLAN · {snap.products.length} SKU · {snap.periodStart} → {snap.periodEnd}</SectionLabel>
      <div style={S.invIntro}>
        Analisa terpadu: posisi stok × performa iklan × BCG quadrant → keputusan per SKU.
        Daily sales = unit terjual ÷ {periodDays} hari periode. Forecast = daily rate × 30 hari.
        {!active && <span style={{ color: C.watch }}> Import file iklan untuk analisa ROAS & rekomendasi iklan.</span>}
      </div>

      {/* Summary bar */}
      <div style={S.kpiGrid}>
        <Kpi label="SKU dipantau" value={`${withStock} / ${skus.length}`} dir="neutral" />
        <Kpi label="Total restock bulan ini" value={`${nfmt(totalRestockQty)} unit`} dir="neutral" />
        <Kpi label="SKU stok kritis (<14hr)" value={skus.filter(s => s.dio !== null && s.dio < 14).length} dir="cost" />
        <Kpi label="SKU overstock (>90hr)" value={skus.filter(s => s.dio !== null && s.dio > 90).length} dir="cost" />
        {active && <Kpi label="SKU perlu pause iklan" value={skus.filter(s => { const r = s.ad && s.dio !== null ? stockAdRec(s.dio, s.ad.roas, thresholds.targetRoas, s.quadrant, s.ad, thresholds) : null; return r && (r.action === "PAUSE + RESTOCK" || r.action === "STOP IKLAN" || r.action === "PAUSE + PO" || r.action === "STOP & AUDIT"); }).length} dir="cost" />}
        {active && <Kpi label="SKU siap scale" value={skus.filter(s => { const r = s.ad && s.dio !== null ? stockAdRec(s.dio, s.ad.roas, thresholds.targetRoas, s.quadrant, s.ad, thresholds) : null; return r && r.action === "SCALE"; }).length} dir="rev" />}
      </div>

      {/* Main table */}
      <SectionLabel>STATUS PER SKU</SectionLabel>
      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead><tr>
            <th style={{ ...S.th, minWidth: 200 }}>SKU</th>
            <th style={S.th}>BCG</th>
            <th style={{ ...S.th, textAlign: "right" }}>Stok</th>
            <th style={{ ...S.th, textAlign: "right" }}>Daily Rate</th>
            <th style={{ ...S.th, textAlign: "right" }}>Days of Stock</th>
            <th style={{ ...S.th, textAlign: "right" }}>
              {(() => { const fc = forecastDemand(1, 30, 30); const m = ["","Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"][fc.month]; return `Forecast ${m} (×${fc.mult})`; })()}
            </th>
            <th style={{ ...S.th, textAlign: "right" }}>Restock Qty</th>
            {active && <th style={{ ...S.th, textAlign: "right" }}>ROAS</th>}
            {active && <th style={{ ...S.th, textAlign: "right" }}>ACOS</th>}
            <th style={{ ...S.th, minWidth: 140 }}>Keputusan</th>
          </tr></thead>
          <tbody>
            {skus.map((s) => {
              const rec = s.ad && s.dio !== null ? stockAdRec(s.dio, s.ad.roas, thresholds.targetRoas, s.quadrant, s.ad, thresholds) : null;
              const recColor = rec ? (rec.level === "bad" ? C.bad : rec.level === "good" ? C.good : C.watch) : C.muted;
              return (
                <tr key={s.code} style={S.tr}>
                  <td style={S.td}>
                    <div style={S.adName}>{s.name}</div>
                    <div style={S.adMeta}>CVR {s.cvr.toFixed(2)}% · {rpShort(s.sales)}</div>
                  </td>
                  <td style={S.td}><QuadTag q={s.quadrant} /></td>
                  <td style={S.tdR}>
                    <input type="number" min="0"
                      defaultValue={s.stockUnits ?? ""}
                      placeholder=" - "
                      onBlur={(e) => setStock(s.code, e.target.value)}
                      style={{ ...S.marginInput, width: 72, textAlign: "right" }} />
                  </td>
                  <td style={S.tdR}>{s.dailyRate > 0 ? s.dailyRate.toFixed(1) + "/hr" : " - "}</td>
                  <td style={{ ...S.tdR, fontWeight: 700, color: s.dioInfo ? s.dioInfo.color : C.muted }}>
                    {s.dio !== null ? <>{s.dio} hr<br /><span style={{ fontFamily: mono, fontSize: 10, fontWeight: 400 }}>{s.dioInfo?.tag}</span></> : " - "}
                  </td>
                  <td style={S.tdR}>
                    {s.forecastUnits > 0
                      ? <>{nfmt(s.forecastUnits)}{s.forecastMult !== 1 && <span style={{ fontFamily: mono, fontSize: 10, color: s.forecastMult > 1 ? C.good : C.watch, marginLeft: 4 }}>×{s.forecastMult}</span>}</>
                      : " - "}
                  </td>
                  <td style={{ ...S.tdR, color: s.restockQty > 0 ? C.watch : C.muted }}>
                    {s.restockQty !== null ? (s.restockQty > 0 ? nfmt(s.restockQty) : "✓ cukup") : " - "}
                  </td>
                  {active && <td style={{ ...S.tdR, fontWeight: 700, color: s.ad ? (s.ad.roas >= thresholds.targetRoas ? C.good : C.bad) : C.muted }}>{s.ad ? s.ad.roas.toFixed(2) : " - "}</td>}
                  {active && <td style={{ ...S.tdR }}>{s.ad ? s.ad.acos.toFixed(1) + "%" : " - "}</td>}
                  <td style={S.td}>
                    {rec
                      ? <span style={{ ...S.pill, color: recColor, borderColor: recColor + "55", background: recColor + "12" }}>{rec.action}</span>
                      : s.stockUnits === null
                        ? <span style={S.muted}>isi stok</span>
                        : !s.ad
                          ? <span style={S.muted}>no ads data</span>
                          : <span style={{ ...S.muted }}> - </span>}
                  </td>
                </tr>
              );
            })}
            {/* totals row */}
            <tr style={{ ...S.tr, background: C.panel2 }}>
              <td style={{ ...S.td, fontFamily: mono, fontSize: 11.5, fontWeight: 700 }} colSpan={2}>TOTAL</td>
              <td style={S.tdR}><b>{nfmt(skus.reduce((t, s) => t + (s.stockUnits || 0), 0))}</b></td>
              <td style={S.tdR}><b>{(skus.reduce((t, s) => t + s.dailyRate, 0)).toFixed(1)}/hr</b></td>
              <td colSpan={2}></td>
              <td style={{ ...S.tdR, color: C.watch }}><b>{nfmt(totalRestockQty)}</b></td>
              {active && <td colSpan={3}></td>}
            </tr>
          </tbody>
        </table>
      </div>

      {/* Diagnosa per SKU  -  consultant-grade */}
      <SectionLabel>DIAGNOSA &amp; SARAN PER SKU</SectionLabel>
      {withStock === 0 && !active && (
        <div style={S.cmEmpty}>Isi stok di kolom tabel dan/atau import file iklan untuk diagnosa lengkap.</div>
      )}
      {skus.filter(s => s.stockUnits != null || s.ad).map((s) => (
        <InventoryDiagnosa key={s.code} s={s} thresholds={thresholds} periodDays={periodDays} />
      ))}
    </div>
  );
}

function InventoryDiagnosa({ s, thresholds, periodDays }) {
  const [open, setOpen] = useState(false);
  const rec = s.ad && s.dio !== null ? stockAdRec(s.dio, s.ad.roas, thresholds.targetRoas, s.quadrant, s.ad, thresholds) : null;
  const QUAD_L = QUAD[s.quadrant];
  const recColor = rec ? (rec.level === "bad" ? C.bad : rec.level === "good" ? C.good : C.watch) : QUAD_L.color;

  // build full consultant-grade analysis
  const analysis = useMemo(() => {
    const pts = [];
    const q = s.quadrant;
    const ad = s.ad;
    const dio = s.dio;
    const roas = ad?.roas;
    const tr = thresholds.targetRoas;
    const beRoas = 100 / (thresholds.targetRoas || 8); // simplified, use targetRoas as floor

    // 1. BCG context
    if (q === "star") {
      pts.push({ type: "bcg", level: "good", title: "BCG: Star  -  mesin revenue", body: `Kontribusi ${s.share.toFixed(1)}% total GMV toko (${rpShort(s.sales)}). Produk ini adalah aset utama  -  keputusan di sini berdampak paling besar ke revenue toko.` });
    } else if (q === "cashcow") {
      pts.push({ type: "bcg", level: "watch", title: "BCG: Cash Cow  -  share tinggi, konversi bocor", body: `Share ${s.share.toFixed(1)}% dengan CVR ${s.cvr.toFixed(2)}% di bawah median toko. Traffic dan exposure sudah ada  -  upside terbesar bukan di spend iklan, tapi di perbaikan halaman produk (foto, harga, deskripsi). Naikin CVR 0.5% di share ini = impact revenue yang signifikan tanpa tambah spend.` });
    } else if (q === "question") {
      pts.push({ type: "bcg", level: "watch", title: "BCG: Question Mark  -  potensi belum terbuka", body: `CVR ${s.cvr.toFixed(2)}% bagus tapi share baru ${s.share.toFixed(1)}%. Produk terbukti bisa menutup penjualan  -  yang kurang cuma exposure. Keputusan kunci: apakah worth invest iklan lebih besar untuk naikan share?` });
    } else {
      pts.push({ type: "bcg", level: "bad", title: "BCG: Dog  -  evaluasi kelayakan", body: `Share ${s.share.toFixed(1)}% dan CVR ${s.cvr.toFixed(2)}% keduanya di bawah median. Sebelum keputusan apapun, identifikasi dulu: apakah ini karena listing belum optimal (foto/harga/deskripsi), atau memang demand pasarnya tidak ada? Iklan di produk ini tanpa fix fundamental = buang modal.` });
    }

    // 2. Stok assessment
    if (dio !== null) {
      if (s.dioInfo.level === "critical") {
        pts.push({ type: "stock", level: "bad", title: `Stok kritis  -  ${dio} hari tersisa`, body: `Dengan laju penjualan ${s.dailyRate.toFixed(1)} unit/hari, stok akan habis dalam ${dio} hari. Forecast demand bulan depan: ${nfmt(s.forecastUnits)} unit${s.forecastMult !== 1 ? ` (rate sekarang ×${s.forecastMult} seasonal multiplier  -  demand ${s.forecastMult > 1 ? "lebih tinggi dari biasanya" : "lebih rendah dari biasanya"})` : ""}. Butuh restock ${nfmt(s.restockQty || 0)} unit segera. Stockout = ranking turun, momentum iklan hilang, dan review negatif. Ini risiko paling mahal untuk brand yang sedang tumbuh.` });
      } else if (s.dioInfo.level === "low") {
        pts.push({ type: "stock", level: "watch", title: `Stok tipis  -  ${dio} hari tersisa`, body: `Window 14–30 hari  -  cukup untuk proses PO kalau supplier leadtime normal, tapi jangan tunda. Forecast bulan depan: ${nfmt(s.restockQty || 0)} unit dibutuhkan${s.forecastMult !== 1 ? ` (×${s.forecastMult} seasonal  -  ${s.forecastMult > 1 ? "peak season mendekat, demand akan naik" : "off-peak, demand lebih rendah"})` : ""}. Konfirmasi PO sekarang dan pertimbangkan kurangi budget iklan sambil menunggu restock.` });
      } else if (s.dioInfo.level === "overstock") {
        pts.push({ type: "stock", level: "watch", title: `Overstock  -  ${dio} hari tersisa`, body: `Stok ${nfmt(s.stockUnits)} unit dengan laju ${s.dailyRate.toFixed(1)}/hari akan habis dalam ${dio} hari  -  modal tertahan terlalu lama. Opportunity cost: modal yang tertahan di stok mati ini harusnya bisa berputar ke SKU lain. Target maksimum ideal 60–90 hari DIO untuk fashion.` });
      } else {
        pts.push({ type: "stock", level: "good", title: `Stok aman  -  ${dio} hari tersisa`, body: `Posisi stok sehat. Forecast demand bulan depan: ${nfmt(s.forecastUnits)} unit${s.forecastMult !== 1 ? ` (×${s.forecastMult} seasonal  -  ${s.forecastMult > 1 ? "demand diperkirakan naik, pastikan buffer stok cukup" : "demand lebih rendah, tidak perlu agresif restock"})` : ""}. ${s.restockQty === 0 ? "Stok mencukupi, tidak perlu restock segera." : `Pertimbangkan PO ${nfmt(s.restockQty || 0)} unit untuk buffer.`}` });
      }
    } else {
      pts.push({ type: "stock", level: "neutral", title: "Stok belum diinput", body: "Isi posisi stok di tabel untuk analisa DIO, forecast kebutuhan restock, dan rekomendasi iklan berbasis kondisi stok." });
    }

    // 3. Ad performance
    if (ad) {
      const acos = ad.acos;
      if (roas >= tr * 1.5) {
        pts.push({ type: "ads", level: "good", title: `Iklan performa kuat  -  ROAS ${roas.toFixed(2)}x`, body: `ROAS ${roas.toFixed(2)}x jauh di atas target ${tr}x dengan spend ${rpShort(ad.spend)}. Mode bidding: ${ad.bidding}. CTR ${ad.ctr.toFixed(2)}%, CVR iklan ${ad.cvr.toFixed(2)}%. Iklan ini sudah proven  -  keputusan scale hanya dibatasi oleh kondisi stok (lihat assessment di atas).` });
      } else if (roas >= tr) {
        pts.push({ type: "ads", level: "good", title: `Iklan on target  -  ROAS ${roas.toFixed(2)}x`, body: `ROAS ${roas.toFixed(2)}x memenuhi target ${tr}x. ACOS ${acos.toFixed(1)}%. Pertahankan  -  tidak perlu intervensi. Pantau drift mingguan: kalau ROAS mulai turun konsisten 2+ minggu, audit materi dan placement.` });
      } else {
        pts.push({ type: "ads", level: "bad", title: `Iklan di bawah target  -  ROAS ${roas.toFixed(2)}x`, body: `ROAS ${roas.toFixed(2)}x di bawah target ${tr}x. ACOS ${acos.toFixed(1)}%, spend ${rpShort(ad.spend)}/periode. ${/GMV Max/i.test(ad.bidding) ? `Mode GMV Max: naikkan target ROAS satu langkah (jangan agresif, beri 3–5 hari learning period tiap perubahan).` : `Mode manual: audit keyword/placement, pertimbangkan longgarkan bid di jam sepi.`} CTR ${ad.ctr.toFixed(2)}%${ad.ctr < thresholds.minCtr ? "  -  di bawah ambang, materi perlu diganti." : " oke."} CVR iklan ${ad.cvr.toFixed(2)}%${ad.cvr < thresholds.minCvr ? "  -  di bawah ambang, cek harga dan halaman produk." : " oke."}` });
      }
    } else {
      pts.push({ type: "ads", level: "neutral", title: "Tidak ada data iklan", body: `Produk ini tidak terdeteksi di file iklan  -  kemungkinan berjalan via Group Ads (Kode Produk tidak tercatat) atau belum diiklankan. ${q === "question" ? "Question Mark dengan CVR bagus = kandidat prioritas untuk dibuatkan iklan produk sendiri." : ""}` });
    }

    // 4. Integrated decision  -  the money paragraph
    const decision = buildDecision(s, ad, dio, roas, tr, thresholds);
    if (decision) pts.push({ type: "decision", level: decision.level, title: "▶ Keputusan & Langkah Konkret", body: decision.body });

    return pts;
  }, [s, thresholds, periodDays]);

  const urgency = analysis.find(p => p.type === "decision")?.level || "neutral";
  const urgencyColor = urgency === "bad" ? C.bad : urgency === "good" ? C.good : C.watch;

  return (
    <div style={{ ...S.skuCard, borderLeftColor: urgencyColor, marginBottom: 10 }}>
      <div style={S.skuHead} onClick={() => setOpen(!open)}>
        <div style={S.skuHeadL}>
          <QuadTag q={s.quadrant} />
          <span style={S.skuName}>{s.name}</span>
        </div>
        <div style={S.skuHeadR}>
          <span style={S.skuStat}>share <b>{s.share.toFixed(1)}%</b></span>
          {s.dio !== null && <span style={{ ...S.skuStat, color: s.dioInfo.color, fontWeight: 700 }}>DIO {s.dio}hr · {s.dioInfo.tag}</span>}
          {s.ad && <span style={{ ...S.skuStat, color: s.ad.roas >= thresholds.targetRoas ? C.good : C.bad }}>ROAS {s.ad.roas.toFixed(2)}</span>}
          {rec && <span style={{ ...S.pill, color: recColor, borderColor: recColor + "44", background: recColor + "12", marginLeft: 4 }}>{rec.action}</span>}
          <span style={{ color: C.muted, fontFamily: mono }}>{open ? "▾" : "▸"}</span>
        </div>
      </div>
      {open && (
        <div style={S.skuBody}>
          {analysis.map((pt, i) => {
            const c = pt.level === "bad" ? C.bad : pt.level === "good" ? C.good : pt.level === "watch" ? C.watch : pt.type === "decision" ? urgencyColor : C.line;
            return (
              <div key={i} style={{ ...S.skuDiag, borderLeftColor: c, background: pt.type === "decision" ? c + "10" : "transparent", marginTop: pt.type === "decision" ? 14 : 8 }}>
                <div style={{ ...S.skuDiagTitle, color: c }}>{pt.title}</div>
                <div style={S.skuDiagMsg}>{pt.body}</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// Build the integrated decision paragraph  -  the most important output

/* ---------- Decision engine (nuanced, senior-grade) ---------------------- */
// ad performance classification  -  separate from stock decision
function classifyAdPerf(ad, thresholds) {
  if (!ad) return null;
  const { roas, acos, ctr, cvr, spend, bidding } = ad;
  const tr = thresholds.targetRoas;
  const isGMX = /GMV Max/i.test(bidding);
  const roasStrong  = roas >= tr * 1.5;
  const roasOk      = roas >= tr;
  const roasMarginal = roas >= tr * 0.75 && roas < tr;
  const roasDead    = roas < tr * 0.75;
  const ctrLow      = ctr < thresholds.minCtr;
  const cvrLow      = cvr < thresholds.minCvr;
  const noVolume    = spend < 100000; // < 100rb spend  -  belum ada data cukup
  return { roas, tr, isGMX, roasStrong, roasOk, roasMarginal, roasDead, ctrLow, cvrLow, noVolume, spend, ctr, cvr, acos };
}

function buildDecision(s, ad, dio, roas, tr, thresholds) {
  const q = s.quadrant;
  const dioLvl = s.dioInfo?.level;
  const p = classifyAdPerf(ad, thresholds);

  // ── STOK KRITIS (<14hr) ──────────────────────────────────────────────────
  if (dioLvl === "critical") {
    if (!p) {
      return { action: "RESTOCK SEGERA", level: "bad",
        body: `Stok tinggal ${dio} hari dengan laju ${s.dailyRate.toFixed(1)} unit/hari. Kebutuhan restock: ${nfmt(s.restockQty || 0)} unit. Konfirmasi PO sekarang  -  ini satu-satunya prioritas. Tanpa data iklan, tidak bisa assess efisiensi spend.` };
    }
    if (p.roasDead || (p.ctrLow && p.cvrLow)) {
      // iklan beneran mati → pause
      return { action: "PAUSE IKLAN + RESTOCK", level: "bad",
        body: `Stok kritis ${dio} hari + iklan tidak efisien (ROAS ${roas.toFixed(1)}x < ${(tr * 0.75).toFixed(1)}x, CTR ${p.ctr.toFixed(2)}%, CVR ${p.cvr.toFixed(2)}%). Pause iklan sekarang  -  tidak ada alasan bayar untuk traffic yang tidak bisa closing dan stoknya hampir habis. Fokus ke: (1) Konfirmasi PO ${nfmt(s.restockQty || 0)} unit segera. (2) Restart iklan hanya setelah restock masuk dan audit materi selesai.` };
    }
    if (p.roasStrong) {
      // iklan bagus → jangan pause, turunkan budget
      return { action: "TURUN BUDGET + RESTOCK", level: "watch",
        body: `Stok kritis ${dio} hari tapi iklan masih sangat efisien (ROAS ${roas.toFixed(1)}x). Jangan pause  -  ranking dan momentum akan turun. Langkah tepat: (1) Kurangi budget harian 40–50% untuk memperlambat laju jual. (2) ${p.isGMX ? "Naikkan target ROAS GMV Max 1–2 langkah agar Shopee lebih selektif pilih pembeli." : "Kurangi bid manual di jam ramai."} (3) Konfirmasi PO ${nfmt(s.restockQty || 0)} unit  -  restock harus masuk sebelum stok habis. Target: pertahankan penjualan ${Math.round(s.dailyRate * 0.5).toFixed(0)}–${Math.round(s.dailyRate * 0.7).toFixed(0)} unit/hari sampai restock tiba.` };
    }
    // iklan marginal → kurangi budget
    return { action: "KURANGI BUDGET + RESTOCK", level: "watch",
      body: `Stok tinggal ${dio} hari, iklan ROAS ${roas.toFixed(1)}x marginal (target ${tr}x). Kurangi budget 30–40% sekarang  -  bukan karena iklan jelek, tapi karena stok tidak cukup untuk aggressive selling. Konfirmasi PO ${nfmt(s.restockQty || 0)} unit. Jangan pause total  -  pertahankan aktivitas iklan minimal agar ranking tidak turun signifikan.` };
  }

  // ── STOK TIPIS (14–30hr) ────────────────────────────────────────────────
  if (dioLvl === "low") {
    if (!p) return { action: "KURANGI AKTIVITAS + PO", level: "watch",
      body: `Stok ${dio} hari  -  window PO terbatas. Konfirmasi restock ${nfmt(s.restockQty || 0)} unit segera. Tanpa data iklan tidak bisa assess spend.` };
    if (p.roasStrong) {
      return { action: "NAIKKAN TARGET ROAS", level: "watch",
        body: `Iklan sangat efisien (ROAS ${roas.toFixed(1)}x) tapi stok tinggal ${dio} hari  -  window sempit. Strategi: ${p.isGMX ? `naikkan target ROAS GMV Max dari ${tr}x ke ${(tr * 1.3).toFixed(1)}x  -  Shopee akan lebih selektif, volume turun tapi margin per unit naik` : "kurangi bid di placement ramai, fokus di jam konversi tinggi"}. Ini bukan karena iklan bermasalah, tapi biar jual lebih efisien sambil stok dihemat. Segera PO ${nfmt(s.restockQty || 0)} unit  -  begitu restock confirmed, kembalikan ke target ROAS semula.` };
    }
    if (p.roasDead) {
      return { action: "PAUSE IKLAN + PO PRIORITAS", level: "bad",
        body: `Stok tipis ${dio} hari + iklan tidak efisien (ROAS ${roas.toFixed(1)}x, spend ${rpShort(p.spend)}/periode sia-sia). Pause iklan  -  tidak worth membakar budget untuk produk yang stoknya tipis dan iklannya tidak convert. Prioritas: PO restock ${nfmt(s.restockQty || 0)} unit. Restart dan audit materi iklan fresh setelah restock masuk.` };
    }
    return { action: "KURANGI BUDGET", level: "watch",
      body: `Stok ${dio} hari, iklan ROAS ${roas.toFixed(1)}x ${p.roasOk ? "oke" : "di bawah target"}. Kurangi budget 20–30% untuk perlambat laju jual. Konfirmasi PO ${nfmt(s.restockQty || 0)} unit. Stockout saat iklan aktif = ranking turun + review negatif + momentum hilang  -  biaya yang lebih mahal dari biaya iklan itu sendiri.` };
  }

  // ── STOK OVERSTOCK (>90hr) ───────────────────────────────────────────────
  if (dioLvl === "overstock") {
    if (!p) return { action: "GERAKIN STOK (ORGANIK)", level: "watch",
      body: `DIO ${dio} hari  -  modal tertahan terlalu lama. Target ideal fashion: 60–90 hari DIO. Aktifkan flash sale atau voucher seller untuk akselerasi perputaran. Tahan restock sampai DIO turun ke bawah 45 hari.` };
    if (p.roasDead) {
      return { action: "STOP IKLAN + CLEARANCE", level: "bad",
        body: `Double problem: stok numpuk ${dio} hari + iklan tidak efisien (ROAS ${roas.toFixed(1)}x < ${(tr * 0.75).toFixed(1)}x). Tidak ada justifikasi membakar budget iklan di kondisi ini. Langkah wajib: (1) Stop iklan sekarang  -  alokasikan budget ke SKU Star/Question Mark. (2) Flash sale agresif: diskon 15–20% selama 7 hari, target habiskan ${Math.round((s.stockUnits || 0) * 0.4)} unit. (3) Jika tidak bergerak dalam 14 hari, pertimbangkan bundle dengan SKU lain atau clearance.` };
    }
    if (p.roasOk) {
      return { action: "REM BUDGET", level: "watch",
        body: `Iklan masih profitable (ROAS ${roas.toFixed(1)}x) tapi tidak perlu didorong  -  stok sudah ${dio} hari. Kurangi budget iklan 40–60%, alihkan ke percepat perputaran stok via flash sale atau promo. Target: turunkan DIO ke bawah 60 hari. Jangan restock sampai DIO < 45 hari. Budget yang dihemat lebih baik dialokasikan ke SKU dengan stok lebih sehat.` };
    }
    return { action: "REM + OPTIMALKAN", level: "watch",
      body: `Overstock ${dio} hari dengan iklan marginal (ROAS ${roas.toFixed(1)}x). Kurangi budget, aktifkan promo organik. ${p.ctrLow ? "CTR rendah  -  ganti thumbnail." : ""} ${p.cvrLow ? "CVR rendah  -  audit harga vs kompetitor." : ""} Fokus gerakin stok dulu sebelum bicara optimasi iklan.` };
  }

  // ── STOK AMAN (30–90hr) ──────────────────────────────────────────────────
  if (!p) return { action: "INPUT IKLAN", level: "neutral",
    body: `Stok aman (${dio !== null ? dio + " hari" : "data belum ada"}). Import file iklan untuk analisa performa dan keputusan budget.` };

  // SCALE  -  kondisi ideal
  if (dioLvl === "safe" && p.roasStrong && (q === "star" || q === "question")) {
    const scaleStep = p.isGMX ? `longgarkan target ROAS GMV Max 0.5 langkah (dari ${tr}x ke ${(tr * 0.9).toFixed(1)}x) atau naikkan budget harian 20–25%` : `naikkan budget 15–20% per langkah, tunggu 3 hari sebelum step berikutnya`;
    return { action: "SCALE", level: "good",
      body: `Kondisi optimal: stok aman ${dio} hari, ROAS ${roas.toFixed(1)}x kuat, BCG ${q === "star" ? "Star" : "Question Mark"} dengan CVR terbukti. Ini window scale yang jarang. Langkah: (1) ${scaleStep}. (2) Jangan ubah materi yang sedang menang. (3) Set tripwire: kalau DIO turun ke bawah 21 hari atau ROAS drop lebih dari 15% dalam seminggu  -  rem kembali. ${q === "question" ? "(4) Kalau setelah scale 2 minggu share naik ke atas median, posisi BCG akan upgrade ke Star." : ""}` };
  }

  // SCALE untuk cashcow  -  hanya kalau CVR sudah oke
  if (dioLvl === "safe" && p.roasStrong && q === "cashcow") {
    return { action: "FIX CVR DULU", level: "watch",
      body: `ROAS ${roas.toFixed(1)}x kuat tapi ini Cash Cow  -  CVR produk ${s.cvr.toFixed(2)}% rendah. Nambah budget di CVR yang bocor = bayar mahal untuk traffic yang tidak closing. Fix dulu: (1) Audit foto utama  -  apakah thumbnail lebih menarik dari kompetitor? (2) Cek harga  -  apakah sudah kompetitif di kategori? (3) Lihat review  -  ada pola keluhan yang bisa diatasi? Begitu CVR naik 0.3–0.5 poin persentase, baru scale budget  -  impact-nya jauh lebih besar dari nambah spend sekarang.` };
  }

  // Stok aman + on target → maintain
  if (dioLvl === "safe" && p.roasOk) {
    return { action: "MAINTAIN", level: "good",
      body: `Posisi sehat: stok ${dio} hari, ROAS ${roas.toFixed(1)}x on target. Tidak ada intervensi mendesak. ${p.isGMX ? "GMV Max sedang berjalan stabil  -  jangan ubah target ROAS kalau tidak ada sinyal penurunan." : "Mode manual stabil."} Pantau weekly: kalau ROAS drift turun 2+ minggu berturut-turut, audit placement dan materi. Kalau DIO mendekati 21 hari tanpa PO confirmed  -  mulai proses restock.` };
  }

  // Stok aman + marginal ROAS → optimalkan
  if (dioLvl === "safe" && p.roasMarginal) {
    return { action: "OPTIMALKAN", level: "watch",
      body: `Stok aman tapi ROAS ${roas.toFixed(1)}x masih ${((1 - roas/tr)*100).toFixed(0)}% di bawah target ${tr}x. ${p.ctrLow ? `CTR ${p.ctr.toFixed(2)}% rendah  -  masalah di atas funnel: materi/thumbnail tidak menarik. Coba ganti visual utama atau hook teks.` : ""} ${p.cvrLow ? `CVR ${p.cvr.toFixed(2)}% rendah  -  masalah di bawah funnel: halaman produk tidak closing. Audit harga vs kompetitor dan foto produk.` : ""} ${p.isGMX ? `Ketatkan target ROAS GMV Max 0.5 langkah untuk filter traffic yang kurang qualified.` : `Kurangi bid di placement dengan CTR rendah.`} Jangan nambah budget sebelum ROAS naik ke target.` };
  }

  // Stok aman + iklan mati → stop dan audit
  if (dioLvl === "safe" && p.roasDead) {
    return { action: "STOP & AUDIT MATERI", level: "bad",
      body: `Stok aman tapi iklan tidak efisien: ROAS ${roas.toFixed(1)}x jauh di bawah target ${tr}x, spend ${rpShort(p.spend)} terbuang. ${p.ctrLow && p.cvrLow ? "CTR dan CVR keduanya rendah  -  masalah sistemik, bukan hanya materi." : p.ctrLow ? "CTR rendah  -  traffic tidak tertarik klik. Ganti materi iklan." : "CVR rendah  -  klik ada tapi tidak closing. Cek halaman produk."} Langkah: (1) Pause iklan sekarang  -  stop bleeding. (2) Audit menyeluruh: materi, harga, foto, review. (3) Relaunch dengan materi baru dan target ROAS lebih ketat. (4) Kalau sudah 3x relaunch masih di bawah target, pertimbangkan apakah kategori ini worth diiklankan.` };
  }

  // Dog dengan stok
  if (q === "dog") {
    const adLine = p ? ` Iklan ROAS ${roas?.toFixed(1)}x  -  ${p.roasOk ? "walau profitable, modal lebih efisien dialihkan ke SKU BCG yang lebih kuat." : "tidak efisien. Stop iklan sekarang."}` : "";
    return { action: p?.roasDead ? "STOP IKLAN + FIX LISTING" : "EVALUASI FUNDAMENTAL", level: "bad",
      body: `Dog: share ${s.share.toFixed(1)}% dan CVR ${s.cvr.toFixed(2)}% keduanya di bawah median toko.${adLine} Framework evaluasi: (1) Masalah listing? → foto profesional, repricing kompetitif, deskripsi ulang. (2) Masalah demand? → riset keyword, lihat volume pencarian kategori. (3) Masalah timing musim? → tahan dan relaunch di peak season. Kalau dalam 30 hari setelah intervensi tidak ada pergerakan berarti, pertimbangkan clearance untuk recovery modal daripada mempertahankan SKU yang menguras perhatian operasional.` };
  }

  return null;
}


/* ============================================================================
   FEE MARKETPLACE TAB
   ========================================================================== */

function FeeKpi({ label, value, color, sub, wide, highlight }) {
  return (
    <div style={{
      background: highlight ? (color ? color + "18" : C.panel) : C.panel,
      border: highlight ? `2px solid ${color || C.bad}` : `1px solid ${color ? color+"33" : C.line}`,
      borderTop: highlight ? `3px solid ${color || C.bad}` : (color ? `3px solid ${color}` : `1px solid ${C.line}`),
      borderRadius: 8, padding: "14px 16px",
      ...(wide ? { gridColumn: "span 2" } : {}),
      ...(highlight ? { boxShadow: `0 0 0 1px ${color || C.bad}22, 0 2px 8px ${color || C.bad}18` } : {})
    }}>
      <div style={{ fontFamily: mono, fontSize: 10, color: highlight ? (color || C.bad) : C.muted, letterSpacing: 1, textTransform: "uppercase", marginBottom: 8, fontWeight: highlight ? 700 : 400 }}>{label}</div>
      <div style={{ fontFamily: mono, fontSize: highlight ? 24 : 20, fontWeight: 700, color: color || C.ink, lineHeight: 1.1 }}>{value}</div>
      {sub && <div style={{ fontFamily: mono, fontSize: 11, color: C.muted, marginTop: 5 }}>{sub}</div>}
    </div>
  );
}

function PanduanModal({ onClose, onGoToOverview }) {
  const sections = [
    {
      title: "DATA IKLAN",
      color: C.accent,
      bg: C.accent + "12",
      items: [
        { label: "Shopee Ads", fmt: "CSV", path: "Pusat Promosi → Iklan Shopee → Semua Iklan Produk → Pilih Periode → Download Data Keseluruhan Iklan → Download", note: "Pastikan pilih 'Data Keseluruhan Iklan', bukan data per produk." },
        { label: "TikTok GMV Max", fmt: "XLSX", path: "Seller Dashboard → Pemasaran → Iklan Toko → Ekspor Data → Pilih Periode → Unduh", note: null },
        { label: "Meta Ads", fmt: "CSV", path: "Meta Ads Manager → Laporan → Pilih Kolom → Pilih Periode → Export CSV", note: "Kolom: Campaign Name, Spend, Results, ROAS, CTR, Impressions, Add to Cart." },
      ]
    },
    {
      title: "DATA KEUANGAN",
      color: C.good,
      bg: C.good + "12",
      items: [
        { label: "Penghasilan Shopee", fmt: "XLSX", path: "Keuangan → Penghasilan Saya → Pilih Periode → Export Data → Download", note: "Dipakai untuk kalkulasi fee marketplace, waterfall revenue, dan analisa retur." },
        { label: "Penghasilan TikTok", fmt: "XLSX", path: "TikTok Shop Seller Center → Keuangan → Penarikan Dana → Pilih Semua → Klik Unduh → Pilih Periode → Download", note: "Pastikan pilih 'Semua' sebelum klik Unduh." },
      ]
    },
    {
      title: "PRODUK & PESANAN",
      color: C.watch,
      bg: C.watch + "12",
      items: [
        { label: "Performa Produk Shopee", fmt: "XLSX", path: "Data → Performa Toko → Klik Produk → Performa Produk → Pilih Periode → Download Data", note: "Dipakai untuk tab Performa Produk, BCG Matrix, Inventory, dan Forecast." },
        { label: "Pesanan Shopee", fmt: "XLSX", path: "Pesanan → Pesanan Saya → Export → Pilih Periode → Download", note: "Dipakai untuk tab Peta Distribusi." },
      ]
    },
    {
      title: "COGS / HPP",
      color: "#8b5cf6",
      bg: "#8b5cf612",
      items: [
        { label: "Template COGS", fmt: "XLSX", path: "Tab COGS / Margin → ⬇ Download Template COGS → Isi HPP Produksi, HPP Packaging, Harga Jual → Import balik via + Import COGS", note: "HPP Packaging default Rp 6.200 — sesuaikan jika berbeda. Data HPP dipakai di Unit Economics, Contribution Margin, dan Simulasi L/R." },
      ]
    },
  ];

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 100, display: "flex", alignItems: "flex-start", justifyContent: "flex-end" }}
      onClick={onClose}>
      <div style={{ width: 480, height: "100vh", background: C.panel, borderLeft: `1px solid ${C.line}`, overflowY: "auto", boxShadow: "-4px 0 24px rgba(0,0,0,0.18)", display: "flex", flexDirection: "column" }}
        onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${C.line}`, position: "sticky", top: 0, background: C.panel, zIndex: 1 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
            <div style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: C.accent, letterSpacing: 1.5 }}>PANDUAN IMPORT DATA</div>
            <button onClick={onClose} style={{ background: "none", border: "none", color: C.muted, fontSize: 18, cursor: "pointer", lineHeight: 1, padding: "2px 6px" }}>✕</button>
          </div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.muted, lineHeight: 1.5 }}>
            Pastikan semua file menggunakan <b style={{ color: C.ink }}>periode yang sama</b> agar perbandingan antar channel akurat.
          </div>
        </div>

        {/* Content */}
        <div style={{ padding: "16px 24px", flex: 1 }}>
          {sections.map((sec, si) => (
            <div key={si} style={{ marginBottom: 20 }}>
              <div style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: sec.color, letterSpacing: 1.8, textTransform: "uppercase", marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ display: "inline-block", width: 12, height: 2, background: sec.color, borderRadius: 2 }} />
                {sec.title}
              </div>
              {sec.items.map((item, ii) => (
                <div key={ii} style={{ background: C.panel2, border: `1px solid ${C.line}`, borderLeft: `3px solid ${sec.color}`, borderRadius: "0 8px 8px 0", padding: "12px 14px", marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
                    <span style={{ fontFamily: mono, fontSize: 12, fontWeight: 700, color: C.ink }}>{item.label}</span>
                    <span style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: sec.color, background: sec.bg, padding: "2px 6px", borderRadius: 4, letterSpacing: 0.5 }}>{item.fmt}</span>
                  </div>
                  <div style={{ fontFamily: mono, fontSize: 10.5, color: C.muted, lineHeight: 1.7 }}>
                    {item.path.split("→").map((step, pi, arr) => (
                      <span key={pi}>
                        <span style={{ color: pi === arr.length - 1 ? C.good : C.ink, fontWeight: pi === arr.length - 1 ? 700 : 400 }}>{step.trim()}</span>
                        {pi < arr.length - 1 && <span style={{ color: C.dim }}> → </span>}
                      </span>
                    ))}
                  </div>
                  {item.note && <div style={{ fontFamily: sans, fontSize: 11, color: C.dim, marginTop: 6, lineHeight: 1.5 }}>{item.note}</div>}
                </div>
              ))}
            </div>
          ))}

          {/* Disclaimer */}
          <div style={{ background: C.watch + "15", border: `1px solid ${C.watch}55`, borderRadius: 8, padding: "12px 14px", marginTop: 4 }}>
            <div style={{ fontFamily: mono, fontSize: 9, fontWeight: 700, color: C.watch, letterSpacing: 1.5, marginBottom: 6 }}>⚠ DISCLAIMER DATA</div>
            <div style={{ fontFamily: sans, fontSize: 11, color: C.ink, lineHeight: 1.6 }}>
              Data bersumber langsung dari file ekspor Shopee & TikTok Shop <b>tanpa modifikasi</b>. Perbedaan dengan laporan akuntansi (Accurate, Jurnal, ERP) bisa terjadi karena perbedaan metode pengakuan revenue, timing pelepasan dana, atau komponen yang dikapitalisasi berbeda. Angka di Sellio mencerminkan data aktual dari platform, bukan angka buku.
            </div>
          </div>
        </div>

        {/* Footer CTA */}
        <div style={{ padding: "16px 24px", borderTop: `1px solid ${C.line}`, background: C.panel }}>
          <button onClick={onGoToOverview}
            style={{ width: "100%", background: C.accent, color: "#fff", border: "none", borderRadius: 8, padding: "11px 0", fontFamily: mono, fontSize: 13, fontWeight: 700, cursor: "pointer", letterSpacing: 0.3 }}>
            Mulai Import Data →
          </button>
        </div>
      </div>
    </div>
  );
}

function DataDisclaimer() {
  return (
    <div style={{ display: "flex", alignItems: "flex-start", gap: 8, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontFamily: sans, fontSize: 12, color: C.muted, lineHeight: 1.55 }}>
      <span style={{ fontSize: 13, marginTop: 1, flexShrink: 0 }}>ℹ️</span>
      <span><b style={{ color: C.ink }}>Data bersumber langsung dari file ekspor Shopee & TikTok Shop tanpa modifikasi.</b> Perbedaan dengan laporan akuntansi bisa terjadi karena perbedaan metode pengakuan revenue, timing pelepasan dana, atau komponen yang dikapitalisasi berbeda (mis. diskon dicatat berbeda antara sistem POS/ERP vs laporan marketplace).</span>
    </div>
  );
}

function FeeEmpty({ onImportShopee, onImportTiktok }) {
  return (
    <div style={S.empty}>
      <div style={S.emptyMark}>◈</div>
      <h2 style={S.emptyTitle}>Belum ada data penghasilan</h2>
      <p style={S.emptyText}>Import file penghasilan dari satu atau dua channel. Tab ini hitung fee rate, waterfall revenue, dan comparison Shopee vs TikTok.</p>
      <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap", marginTop: 20 }}>
        <div style={S.emptyCard}>
          <div style={S.emptyCardTitle}>Shopee</div>
          <div style={S.emptyCardDesc}>Keuangan → Penghasilan Saya → Download XLSX</div>
        </div>
        <div style={S.emptyCard}>
          <div style={S.emptyCardTitle}>TikTok Shop</div>
          <div style={S.emptyCardDesc}>Keuangan → Pendapatan → Download → "Detail pesanan" XLSX</div>
        </div>
      </div>
      <p style={{ ...S.emptyText, marginTop: 16 }}>Import via tombol <b>+ Import Data ▼</b> di kanan atas.</p>
    </div>
  );
}

function FeeTab({ shopeeSnap, tiktokSnap }) {
  const [activeChannel, setActiveChannel] = React.useState("combined");
  const snap = activeChannel === "tiktok" ? tiktokSnap : activeChannel === "shopee" ? shopeeSnap : null;
  const s = snap ? snap.summary : null;

  // Combined totals
  const combined = useMemo(() => {
    const sh = shopeeSnap?.summary;
    const tt = tiktokSnap?.summary;
    if (!sh && !tt) return null;
    const safeNum = (v) => (typeof v === "number" && !isNaN(v) ? v : 0);
    const c = {
      hargaAsli:       safeNum(sh?.hargaAsli)    + safeNum(tt?.gmvBruto),
      diskonProduk:    safeNum(sh?.diskonProduk)  + safeNum(tt?.diskonPenjual),
      netGmv:          safeNum(sh?.netGmv)        + safeNum(tt?.netGmv),
      totalFee:        safeNum(sh?.totalFee)      + safeNum(tt?.totalFee),
      totalDilepas:    safeNum(sh?.totalDilepas)  + safeNum(tt?.totalDilepas),
      komisiAMS:       safeNum(sh?.komisiAMS)     + safeNum(tt?.komisiPlatform),
      biayaAdmin:      safeNum(sh?.biayaAdmin)    + safeNum(tt?.komisiAfiliasi) + safeNum(tt?.komisiDinamis),
      biayaLayanan:    safeNum(sh?.biayaLayanan),
      biayaProses:     safeNum(sh?.biayaProses)   + safeNum(tt?.biayaProses),
      adSpendTiktok:   safeNum(tt?.adSpend),
    };
    c.feeRateNetGmv = c.netGmv ? Math.abs(c.totalFee) / c.netGmv * 100 : 0;
    c.feeRateGross  = c.hargaAsli ? Math.abs(c.totalFee) / c.hargaAsli * 100 : 0;
    c.takeRate      = c.netGmv ? (1 - c.totalDilepas / c.netGmv) * 100 : 0;
    c.shopeeShare   = c.netGmv && sh ? safeNum(sh.netGmv) / c.netGmv * 100 : 0;
    c.tiktokShare   = c.netGmv && tt ? safeNum(tt.netGmv) / c.netGmv * 100 : 0;
    // Combined refund metrics
    c.refundOrderCount = safeNum(sh?.refundOrderCount) + safeNum(tt?.refundOrderCount);
    c.refundValue      = safeNum(sh?.refundValue)      + safeNum(tt?.refundValue);
    c.totalOrderCount  = safeNum(sh?.totalOrderCount)  + safeNum(tt?.totalOrderCount);
    c.refundRateOrder  = c.totalOrderCount ? (c.refundOrderCount / c.totalOrderCount) * 100 : 0;
    c.refundRateValue  = c.hargaAsli ? (c.refundValue / c.hargaAsli) * 100 : 0;
    return c;
  }, [shopeeSnap, tiktokSnap]);

  // For single channel view, use its summary normalized
  const displaySummary = useMemo(() => {
    if (activeChannel === "combined") return combined;
    if (activeChannel === "shopee" && shopeeSnap) {
      const s = { ...shopeeSnap.summary };
      // Fallback: compute refund metrics from orders if not cached (old DB_VER)
      if (s.totalOrderCount == null && shopeeSnap.orders?.length > 0) {
        const refundOrders = shopeeSnap.orders.filter(o => o.pengembalianDana < 0);
        s.totalOrderCount  = shopeeSnap.orders.length;
        s.refundOrderCount = refundOrders.length;
        s.refundValue      = refundOrders.reduce((t,o) => t + Math.abs(o.pengembalianDana||0), 0);
        s.refundRateOrder  = shopeeSnap.orders.length ? (refundOrders.length / shopeeSnap.orders.length) * 100 : 0;
        s.refundRateValue  = s.hargaAsli ? (s.refundValue / s.hargaAsli) * 100 : 0;
      }
      return s;
    }
    if (activeChannel === "tiktok" && tiktokSnap) {
      const t = { ...tiktokSnap.summary };
      // Fallback: compute refund metrics from orders if not cached (old DB_VER)
      if (t.totalOrderCount == null && tiktokSnap.orders?.length > 0) {
        const refundOrders = tiktokSnap.orders.filter(o => o.subtotalRefund < 0);
        t.totalOrderCount  = tiktokSnap.orders.length;
        t.refundOrderCount = refundOrders.length;
        t.refundValue      = refundOrders.reduce((t2,o) => t2 + Math.abs(o.subtotalRefund||0), 0);
        t.refundRateOrder  = tiktokSnap.orders.length ? (refundOrders.length / tiktokSnap.orders.length) * 100 : 0;
        t.refundRateValue  = t.gmvBruto ? (t.refundValue / t.gmvBruto) * 100 : 0;
      }
      return { ...t, hargaAsli: t.gmvBruto, diskonProduk: t.diskonPenjual };
    }
    return null;
  }, [activeChannel, shopeeSnap, tiktokSnap, combined]);

  if (!combined) return <div style={S.cmEmpty}>Tidak ada data income.</div>;
  const ds = displaySummary;

  // Channel switcher labels
  const channelLabel = activeChannel === "combined" ? "Semua Channel" : activeChannel === "shopee" ? "Shopee" : "TikTok Shop";
  const periodLabel = activeChannel === "combined"
    ? [shopeeSnap, tiktokSnap].filter(Boolean).map(s => (s.channel==="tiktok"?"TikTok ":"Shopee ") + (s.periodStart||"") + "→" + (s.periodEnd||"")).join("  ·  ")
    : snap ? ((snap.periodStart||"") + " → " + (snap.periodEnd||"")) : "";

  // Waterfall — adapts per channel
  const isTikTok = activeChannel === "tiktok";
  const waterfall = [
    { label: "GMV Bruto (Harga Asli)", value: ds?.hargaAsli || ds?.gmvBruto || 0, type: "gross" },
    { label: isTikTok ? "Diskon Penjual" : "Total Diskon Produk", value: ds?.diskonProduk || ds?.diskonPenjual || 0, type: "deduct" },
    ...(!isTikTok && ds?.refund ? [{ label: "Refund ke Pembeli", value: ds.refund, type: "deduct" }] : []),
    ...(!isTikTok && ds?.voucherSeller ? [{ label: "Voucher Penjual", value: ds.voucherSeller, type: "deduct" }] : []),
    { label: "NET GMV", value: ds?.netGmv || 0, type: "subtotal" },
    ...(!isTikTok && ds?.ongkirBuyer != null ? [{ label: "Net Ongkir", value: (ds.ongkirBuyer||0)+(ds.gratisOngkirShopee||0)+(ds.ongkirDiteruskan||0)+(ds.promoOngkirSeller||0), type: "deduct" }] : []),
    { label: isTikTok ? "Biaya Komisi Platform" : "Biaya Komisi AMS", value: ds?.komisiAMS || ds?.komisiPlatform || 0, type: "fee" },
    { label: isTikTok ? "Komisi Afiliasi" : "Biaya Administrasi (inkl. PPN 11%)", value: ds?.biayaAdmin || ds?.komisiAfiliasi || 0, type: "fee" },
    ...(!isTikTok ? [{ label: "Biaya Layanan", value: ds?.biayaLayanan || 0, type: "fee" }] : []),
    ...(isTikTok && ds?.komisiDinamis ? [{ label: "Komisi Dinamis", value: ds.komisiDinamis, type: "fee" }] : []),
    { label: "Biaya Proses Pesanan", value: ds?.biayaProses || 0, type: "fee" },
    { label: "TOTAL FEE MARKETPLACE", value: ds?.totalFee || 0, type: "feetotal" },
    { label: "TOTAL DILEPAS KE SALDO", value: ds?.totalDilepas || 0, type: "result" },
  ].filter(r => r.value !== 0 || r.type === "result" || r.type === "subtotal");

  const feeComponents = activeChannel === "tiktok" ? [
    { label: "Biaya Komisi Platform", value: ds?.komisiPlatform || ds?.komisiAMS || 0 },
    { label: "Komisi Afiliasi", value: ds?.komisiAfiliasi || 0 },
    { label: "Komisi Dinamis", value: ds?.komisiDinamis || 0 },
    { label: "Biaya Proses Pesanan", value: ds?.biayaProses || 0 },
  ] : activeChannel === "shopee" ? [
    { label: "Biaya Layanan", value: ds?.biayaLayanan || 0 },
    { label: "Biaya Administrasi (inkl. PPN 11%)", value: ds?.biayaAdmin || 0 },
    { label: "Biaya Komisi AMS", value: ds?.komisiAMS || 0 },
    { label: "Biaya Proses Pesanan", value: ds?.biayaProses || 0 },
  ] : [
    { label: "Shopee: Biaya Layanan", value: shopeeSnap?.summary?.biayaLayanan || 0 },
    { label: "Shopee: Biaya Admin", value: shopeeSnap?.summary?.biayaAdmin || 0 },
    { label: "Shopee: Komisi AMS", value: shopeeSnap?.summary?.komisiAMS || 0 },
    { label: "TikTok: Komisi Platform", value: tiktokSnap?.summary?.komisiPlatform || 0 },
    { label: "TikTok: Komisi Afiliasi", value: tiktokSnap?.summary?.komisiAfiliasi || 0 },
    { label: "TikTok: Komisi Dinamis", value: tiktokSnap?.summary?.komisiDinamis || 0 },
  ].filter(f => f.value !== 0);
  const totalFeeAbs = Math.abs(ds?.totalFee || 0);

  const rowColor = (type) => {
    if (type === "gross") return C.ink;
    if (type === "subtotal") return C.accent;
    if (type === "feetotal") return C.bad;
    if (type === "result") return C.good;
    if (type === "fee") return C.bad;
    if (type === "deduct") return C.watch;
    return C.muted;
  };
  const rowBg = (type) => {
    if (type === "subtotal") return C.accent + "0d";
    if (type === "feetotal") return C.bad + "0d";
    if (type === "result") return C.good + "0d";
    return "transparent";
  };

  return (
    <div>
      <SectionLabel>FEE MARKETPLACE · {channelLabel}{periodLabel ? " · " + periodLabel : ""}</SectionLabel>
      <DataDisclaimer />
      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        {[["combined","Semua Channel"], ["shopee","Shopee"], ["tiktok","TikTok Shop"]].map(([k, lbl]) => {
          const disabled = (k === "shopee" && !shopeeSnap) || (k === "tiktok" && !tiktokSnap);
          return (
            <button key={k} disabled={disabled} onClick={() => setActiveChannel(k)}
              style={{ fontFamily: mono, fontSize: 11.5, fontWeight: 700, padding: "6px 14px", borderRadius: 6,
                cursor: disabled ? "not-allowed" : "pointer",
                border: `1px solid ${activeChannel===k ? C.accent : C.line}`,
                background: activeChannel===k ? C.accent+"1a" : C.panel,
                color: disabled ? C.dim : activeChannel===k ? C.accent : C.muted,
                opacity: disabled ? 0.4 : 1 }}>
              {lbl}
            </button>
          );
        })}
      </div>

      {/* Missing channel warning */}
      {activeChannel === "combined" && (!shopeeSnap || !tiktokSnap) && (
        <div style={{ background: C.watch+"15", border: `1px solid ${C.watch}55`, borderRadius: 8, padding: "10px 14px", marginBottom: 12, fontFamily: sans, fontSize: 13, lineHeight: 1.5 }}>
          <b style={{ color: C.watch }}>⚠ Data belum lengkap —</b>
          {!shopeeSnap && <span> Shopee belum diimport.</span>}
          {!tiktokSnap && <span> TikTok belum diimport.</span>}
          <span style={{ color: C.muted }}> Angka di bawah hanya dari channel yang sudah ada. Import kedua channel untuk perbandingan penuh.</span>
        </div>
      )}

      {/* KPI strip — color coded */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 10 }}>
        <FeeKpi label="GMV Bruto (Harga Asli)" value={rpShort(ds?.hargaAsli || ds?.gmvBruto || 0)} color={C.muted} sub={"sebelum diskon & fee"} />
        <FeeKpi label="Net GMV (setelah diskon)" value={rpShort(ds?.netGmv || 0)} color={C.accent} sub={"basis fee rate"} />
        <FeeKpi label="Total Fee Marketplace" value={rpShort(Math.abs(ds?.totalFee || 0))} color={C.bad} sub={channelLabel + " · fee diambil marketplace"} />
        <FeeKpi label="Fee Rate vs Net GMV" value={(ds?.feeRateNetGmv || 0).toFixed(1)+"%"} color={C.bad} sub={"basis pricing & margin"} highlight />
        <FeeKpi label="Total Dilepas ke Saldo" value={rpShort(ds?.totalDilepas || 0)} color={C.good} sub={"masuk saldo penjual"} />
        <FeeKpi label="Fee Rate vs Harga Asli" value={(ds?.feeRateGross || 0).toFixed(1)+"%"} color={C.watch} />
        <FeeKpi label="Take Rate Total" value={(ds?.takeRate || 0).toFixed(1)+"%"} color={C.watch} sub={"% Net GMV yg tidak masuk saldo"} />
        {snap?.orders && <FeeKpi label="Jumlah Pesanan" value={snap.orders.length.toLocaleString("id-ID")} color={C.muted} sub={(snap.periodStart||"") + " → " + (snap.periodEnd||"")} />}
        {ds?.totalOrderCount > 0 && <FeeKpi
          label="Retur / Refund (order)"
          value={(ds.refundRateOrder || 0).toFixed(1) + "%"}
          color={ds.refundRateOrder > 5 ? C.bad : ds.refundRateOrder > 2 ? C.watch : C.good}
          sub={`${ds.refundOrderCount} dari ${ds.totalOrderCount} order · ${rpShort(ds.refundValue || 0)}`}
        />}
        {ds?.totalOrderCount > 0 && <FeeKpi
          label="Retur / Refund (nilai)"
          value={(ds.refundRateValue || 0).toFixed(2) + "%"}
          color={ds.refundRateValue > 5 ? C.bad : ds.refundRateValue > 2 ? C.watch : C.good}
          sub={"dari total GMV Bruto · " + (ds.refundRateValue < 1 ? "sehat ✓" : ds.refundRateValue < 3 ? "normal" : "perlu dicek")}
        />}
      </div>

      {/* Fee rate context */}
      <div style={{ ...S.invIntro, marginTop: 16, fontFamily: sans }}>
        <b style={{ color: C.bad }}>Fee rate {(ds?.feeRateNetGmv||0).toFixed(1)}% dari Net GMV ({channelLabel})</b> — dari setiap Rp 100 penjualan bersih, Rp {(ds?.feeRateNetGmv||0).toFixed(0)} langsung ke marketplace sebelum COGS dan ad spend.
        {ds?.hargaAsli && (ds?.diskonProduk || ds?.diskonPenjual) ? <>{" "}Diskon {rpShort(Math.abs(ds.diskonProduk||ds.diskonPenjual||0))} ({ds.hargaAsli ? (Math.abs(ds.diskonProduk||ds.diskonPenjual||0)/ds.hargaAsli*100).toFixed(0) : 0}% dari harga asli) — pastikan ini sudah di-factor ke pricing.</> : null}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginTop: 8 }}>

        {/* Waterfall */}
        <div>
          <SectionLabel>WATERFALL PENDAPATAN</SectionLabel>
          <div style={{ ...S.tableWrap }}>
            <table style={S.table}>
              <tbody>
                {waterfall.map((row, i) => (
                  <tr key={i} style={{ ...S.tr, background: rowBg(row.type) }}>
                    <td style={{ ...S.td, fontFamily: sans, fontWeight: (row.type === "subtotal" || row.type === "feetotal" || row.type === "result") ? 700 : 400, fontSize: (row.type === "subtotal" || row.type === "feetotal" || row.type === "result") ? 13.5 : 13, paddingLeft: (row.type === "fee" || row.type === "deduct") ? 24 : 12, letterSpacing: (row.type === "subtotal" || row.type === "feetotal" || row.type === "result") ? 0 : "normal" }}>
                      {row.label}
                    </td>
                    <td style={{ ...S.tdR, fontFamily: mono, color: rowColor(row.type), fontWeight: (row.type === "subtotal" || row.type === "feetotal" || row.type === "result") ? 700 : 400 }}>
                      {row.value >= 0 ? "" : "−"}{rpShort(Math.abs(row.value))}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Fee breakdown donut-style bars */}
        <div>
          <SectionLabel>BREAKDOWN FEE · Rp {(totalFeeAbs/1e6).toFixed(1)}jt total</SectionLabel>
          <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: 16 }}>
            {feeComponents.map((fc, i) => {
              const pctOfFee = totalFeeAbs ? Math.abs(fc.value) / totalFeeAbs * 100 : 0;
              const pctOfNetGmv = ds?.netGmv ? Math.abs(fc.value) / ds.netGmv * 100 : 0;
              const barColors = [C.bad, "#e06c75", "#d19a66", C.watch];
              return (
                <div key={i} style={{ marginBottom: 18 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, alignItems: "baseline" }}>
                    <span style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 600 }}>{fc.label}</span>
                    <span style={{ fontFamily: mono, fontSize: 12.5, color: C.bad, fontWeight: 700 }}>{rpShort(Math.abs(fc.value))}</span>
                  </div>
                  <div style={{ background: C.panel2, borderRadius: 4, height: 8, overflow: "hidden" }}>
                    <div style={{ width: pctOfFee + "%", background: barColors[i], height: "100%", borderRadius: 4 }} />
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
                    <span style={{ fontFamily: mono, fontSize: 10.5, color: C.muted }}>{pctOfFee.toFixed(1)}% dari total fee</span>
                    <span style={{ fontFamily: mono, fontSize: 10.5, color: C.muted }}>{pctOfNetGmv.toFixed(1)}% dari Net GMV</span>
                  </div>
                </div>
              );
            })}
            <div style={{ borderTop: `1px solid ${C.line}`, paddingTop: 12, marginTop: 4 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
                <span style={{ fontFamily: mono, fontSize: 11, letterSpacing: 1, color: C.muted, textTransform: "uppercase" }}>Total Fee</span>
                <span style={{ fontFamily: mono, fontSize: 14, fontWeight: 700, color: C.bad }}>{rpShort(totalFeeAbs)}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
                <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>vs Net GMV</span>
                <span style={{ fontFamily: mono, fontSize: 11, color: C.bad, fontWeight: 700 }}>{(ds?.feeRateNetGmv||0).toFixed(1)}%</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>vs Harga Asli</span>
                <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>{(ds?.feeRateGross||0).toFixed(1)}%</span>
              </div>
            </div>
          </div>

          {/* Metode pembayaran breakdown */}
          {snap?.orders?.length > 0 && (() => {
            const isTT = activeChannel === "tiktok";
            const methods = {};
            snap.orders.forEach(o => {
              const key = isTT
                ? (o.sumberPesanan || "TikTok Shop")
                : (o.metodePembayaran || "—");
              methods[key] = (methods[key] || 0) + 1;
            });
            const sorted = Object.entries(methods).sort((a,b) => b[1]-a[1]);
            const label = isTT ? "SUMBER PESANAN" : "METODE PEMBAYARAN";
            return (
              <div style={{ marginTop: 16 }}>
                <SectionLabel>{label} · {snap.orders.length} pesanan</SectionLabel>
                <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: 14 }}>
                  {sorted.map(([method, count], i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", borderBottom: i < sorted.length-1 ? `1px solid ${C.line}` : "none" }}>
                      <span style={{ fontFamily: sans, fontSize: 12.5 }}>{method}</span>
                      <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                        <div style={{ background: C.panel2, borderRadius: 3, height: 6, width: 80, overflow: "hidden" }}>
                          <div style={{ width: (count/(snap.orders.length||1)*100)+"%", background: C.accent, height: "100%" }} />
                        </div>
                        <span style={{ fontFamily: mono, fontSize: 11.5, color: C.muted, minWidth: 60, textAlign: "right" }}>{count} ({(count/(snap.orders.length||1)*100).toFixed(0)}%)</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            );
          })()}
        </div>
      </div>

      {/* Per-order avg */}
      {(activeChannel !== "combined" && snap?.orders?.length > 0) && (() => {
        const n = snap.orders.length;
        const isTT = activeChannel === "tiktok";
        const avgHarga   = snap.orders.reduce((t,o) => t + (isTT ? (o.subtotalBruto||0) : (o.hargaAsli||0)), 0) / n;
        const avgDilepas = snap.orders.reduce((t,o) => t + (isTT ? (o.settlement||0) : (o.totalDilepas||0)), 0) / n;
        const avgFee     = isTT
          ? snap.orders.reduce((t,o) => t + Math.abs(o.totalBiaya||0), 0) / n
          : snap.orders.reduce((t,o) => t + Math.abs((o.komisiAMS||0)+(o.biayaAdmin||0)+(o.biayaLayanan||0)+(o.biayaProses||0)), 0) / n;
        const avgDiskon  = snap.orders.reduce((t,o) => t + Math.abs(isTT ? (o.diskonPenjual||0) : (o.diskonProduk||0)), 0) / n;
        const feeDisp    = avgDilepas > 0 ? (avgFee / (avgDilepas + avgFee) * 100) : 0;
        const realisasi  = avgHarga > 0 ? (avgDilepas / avgHarga * 100) : 0;
        return (
          <div style={{ marginTop: 8 }}>
            <SectionLabel>RATA-RATA PER PESANAN</SectionLabel>
            <div style={S.kpiGrid}>
              <Kpi label="Avg Harga Asli" value={rpShort(avgHarga)} dir="rev" accent={C.muted} />
              <Kpi label="Avg Diskon" value={rpShort(avgDiskon)} dir="cost" accent={C.watch} />
              <Kpi label="Avg Fee Marketplace" value={rpShort(avgFee)} dir="cost" accent={C.bad} />
              <Kpi label="Avg Dilepas ke Saldo" value={rpShort(avgDilepas)} dir="rev" accent={C.good} />
              <Kpi label="Fee / Order (% dilepas)" value={feeDisp.toFixed(1)+"%"} dir="cost" big accent={C.bad} />
              <Kpi label="Realisasi vs Harga Asli" value={realisasi.toFixed(1)+"%"} dir="neutral" accent={C.muted} />
            </div>
          </div>
        );
      })()}

      {/* Diagnosa & Saran */}
      <FeeDiagnosa s={ds} />
    </div>
  );
}


function FeeDiagnosa({ s: rawS }) {
  const s = rawS || {};
  const items = useMemo(() => {
    const pts = [];
    const feeRate = s.feeRateNetGmv || 0;
    const diskonRate = s.hargaAsli ? Math.abs(s.diskonProduk || 0) / s.hargaAsli * 100 : 0;
    const ongkirNet = (s.ongkirBuyer||0) + (s.gratisOngkirShopee||0) + (s.ongkirDiteruskan||0) + (s.promoOngkirSeller||0);
    const layananShare = s.totalFee ? Math.abs(s.biayaLayanan) / Math.abs(s.totalFee) * 100 : 0;
    const adminShare = s.totalFee ? Math.abs(s.biayaAdmin) / Math.abs(s.totalFee) * 100 : 0;

    // 1. Overall fee rate assessment
    if (feeRate > 28) {
      pts.push({ level: "bad", title: `Fee rate tinggi: ${feeRate.toFixed(1)}% dari Net GMV`, body: `Lebih dari seperempat revenue bersih langsung ke Shopee. Threshold waspada: >25%. Di angka ini, bisnis harus punya gross margin produk minimal 45–50% agar masih ada ruang untuk COGS, logistik, dan profit setelah fee. Audit: apakah kamu aktif di program gratis ongkir XTRA? Itu salah satu kontributor terbesar Biaya Layanan.` });
    } else if (feeRate > 22) {
      pts.push({ level: "watch", title: `Fee rate moderat: ${feeRate.toFixed(1)}% dari Net GMV`, body: `Masih dalam range normal fashion marketplace (20–27%), tapi sudah di zona yang perlu diperhatikan. Pastikan pricing sudah factor fee ini — kalau harga jual dihitung dari COGS + margin tanpa add-back fee 24%, margin aktual bisa 6–10 poin lebih tipis dari yang tercatat.` });
    } else {
      pts.push({ level: "good", title: `Fee rate efisien: ${feeRate.toFixed(1)}% dari Net GMV`, body: `Di bawah 22% — tergolong efisien untuk fashion marketplace. Pertahankan komposisi program promo yang sekarang dan hindari program berbayar yang tidak memberikan uplift konversi terukur.` });
    }

    // 2. Biaya Layanan — biggest fee component
    if (layananShare > 45) {
      pts.push({ level: "watch", title: `Biaya Layanan dominasi ${layananShare.toFixed(0)}% dari total fee (${rpShort(Math.abs(s.biayaLayanan))})`, body: `Biaya Layanan di Shopee terdiri dari: fee gratis ongkir XTRA (Kategori F), fee program hemat ongkir, dan fee layanan dasar. Kalau >45% fee kamu dari komponen ini, artinya kamu heavily exposed ke program gratis ongkir. Evaluasi: apakah setiap pesanan yang masuk dari program ini punya margin yang cukup setelah dikurangi fee-nya? Program gratis ongkir bagus untuk volume, tapi bisa jadi margin killer kalau AOV rendah.` });
    }

    // 3. Biaya Admin (PPN 11%) — structural, tidak bisa dikurangi
    if (adminShare > 38) {
      pts.push({ level: "neutral", title: `Biaya Administrasi ${rpShort(Math.abs(s.biayaAdmin))} (inkl. PPN 11%) — komponen struktural`, body: `Biaya Administrasi adalah komponen yang tidak bisa dihindari — ini include PPN 11% dari komisi. Tidak perlu dioptimasi karena sifatnya proporsional terhadap penjualan. Yang perlu dicatat: komponen ini harus masuk ke perhitungan harga pokok penjualan di pembukuan, bukan dianggap "biaya lain-lain".` });
    }

    // 4. Diskon produk rate
    if (diskonRate > 35) {
      pts.push({ level: "bad", title: `Diskon produk ${diskonRate.toFixed(0)}% dari harga asli — terlalu dalam`, body: `Diskon ${rpShort(Math.abs(s.diskonProduk))} dari harga asli ${rpShort(s.hargaAsli)}. Diskon >35% dari harga listing adalah sinyal bahwa harga asli sudah "di-mark up" untuk akomodasi diskon, atau kamu terlalu agresif di voucher/flash sale. Dampak ganda: (1) diskon ini mengurangi Net GMV yang jadi basis fee — fee tetap dihitung dari transaksi aktual. (2) pembeli yang terbiasa beli di harga diskon akan sulit convert di harga normal. Rekomendasi: audit SKU dengan diskon >40%, evaluasi apakah bisa mengurangi depth diskon dengan tetap stay competitive.` });
    } else if (diskonRate > 25) {
      pts.push({ level: "watch", title: `Diskon produk ${diskonRate.toFixed(0)}% dari harga asli — perlu dikontrol`, body: `Diskon di range 25–35% masih umum di fashion marketplace, tapi perlu dipantau. Pastikan setiap flash sale atau voucher yang kamu jalankan punya minimum order value yang protect margin. Hitung break-even: kalau gross margin 40% dan fee 24%, diskon >16% dari harga asli sudah mulai makan ke profit.` });
    } else {
      pts.push({ level: "good", title: `Diskon produk ${diskonRate.toFixed(0)}% dari harga asli — terkontrol`, body: `Depth diskon tergolong moderat. Artinya harga listing kamu sudah cukup kompetitif tanpa perlu heavy markdown, atau kamu selective dalam program promo. Pertahankan dan monitor per-SKU.` });
    }

    // 5. Net ongkir assessment
    if (ongkirNet < -3000000) {
      pts.push({ level: "watch", title: `Net ongkir negatif ${rpShort(Math.abs(ongkirNet))} — shipping cost signifikan`, body: `Kamu menanggung net ongkir ${rpShort(Math.abs(ongkirNet))} setelah diperhitungkan semua subsidi Shopee dan ongkir yang dibayar pembeli. Ini bukan masalah kalau sudah di-build ke harga jual. Tapi kalau tidak, ini "biaya tersembunyi" yang menggerus margin tanpa terlihat di P&L sederhana. Cek: apakah AOV-mu cukup untuk justify free shipping tanpa minimum order? Benchmark: untuk fashion, minimum order Rp 150–200rb sebelum offer free shipping.` });
    }

    // 6. Pricing recommendation
    const minPriceMultiplier = 1 / (1 - (feeRate/100) - 0.05); // fee + 5% buffer
    pts.push({ level: "neutral", title: "Implikasi ke Pricing", body: `Dengan fee rate ${feeRate.toFixed(1)}%, untuk target contribution margin 20% dari harga jual, kamu butuh gross margin produk minimal ${(feeRate + 20 + 5).toFixed(0)}% (fee ${feeRate.toFixed(0)}% + CM target 20% + buffer 5%). Kalau COGS produk kamu Rp X, harga jual minimum = COGS ÷ (1 − ${((feeRate + 25)/100).toFixed(2)}) = COGS × ${minPriceMultiplier.toFixed(2)}. Masukkan angka ini ke tab COGS / Margin untuk kalkulasi per SKU.` });

    return pts;
  }, [s]);

  const levelColor = (l) => l === "bad" ? C.bad : l === "good" ? C.good : l === "watch" ? C.watch : C.muted;

  return (
    <div style={{ marginTop: 8 }}>
      <SectionLabel>DIAGNOSA &amp; SARAN STRUKTUR FEE</SectionLabel>
      {items.map((pt, i) => (
        <div key={i} style={{ borderLeft: `3px solid ${levelColor(pt.level)}`, background: levelColor(pt.level) + "0a", borderRadius: "0 8px 8px 0", padding: "12px 16px", marginBottom: 10 }}>
          <div style={{ fontFamily: mono, fontSize: 11.5, fontWeight: 700, color: levelColor(pt.level), letterSpacing: 0.3, marginBottom: 6 }}>{pt.title}</div>
          <div style={{ fontSize: 13, color: C.ink, lineHeight: 1.6, fontFamily: sans }}>{pt.body}</div>
        </div>
      ))}
    </div>
  );
}


/* ============================================================================
   SEBARAN AREA TAB
   ========================================================================== */
function AreaEmpty({ onImport }) {
  return (
    <div style={S.empty}>
      <div style={S.emptyMark}>📍</div>
      <h2 style={S.emptyTitle}>Belum ada data pesanan</h2>
      <p style={S.emptyText}>
        Export <b>Data Pesanan</b> dari Shopee (Pesanan → Export Pesanan → pilih periode → Download XLSX),
        import via tombol <b>+ Import Data ▼</b> di kanan atas → Data Pesanan Shopee.
      </p>
      <button style={S.importBtn2Big} onClick={onImport}>+ Import Data Pesanan</button>
    </div>
  );
}

function AreaTab({ snap }) {
  const [view, setView] = React.useState("provinsi"); // "provinsi" | "kota"
  const [sortBy, setSortBy] = React.useState("pesanan"); // "pesanan" | "pcs"

  const { byProvinsi, byKota, highlights } = useMemo(() => {
    const orders = snap.orders || [];
    const total = orders.length;
    const totalPcs = orders.reduce((t, o) => t + o.jumlah, 0);
    const selesai = orders.filter(o => o.selesai).length;
    const cancelRate = snap.orders.length > 0
      ? 100 - (orders.length / (orders.length + (snap._rawTotal || orders.length)) * 100) : 0;

    // Group by provinsi
    const provMap = {};
    const kotaMap = {};
    orders.forEach(o => {
      const prov = o.provinsi || "Tidak diketahui";
      const kota = o.kota || "Tidak diketahui";
      if (!provMap[prov]) provMap[prov] = { pesanan: 0, pcs: 0, kota: {} };
      provMap[prov].pesanan++;
      provMap[prov].pcs += o.jumlah;
      provMap[prov].kota[kota] = (provMap[prov].kota[kota] || 0) + o.jumlah;
      if (!kotaMap[kota]) kotaMap[kota] = { pesanan: 0, pcs: 0, provinsi: prov };
      kotaMap[kota].pesanan++;
      kotaMap[kota].pcs += o.jumlah;
    });

    const byProvinsi = Object.entries(provMap)
      .map(([name, d]) => ({ name, pesanan: d.pesanan, pcs: d.pcs, topKota: Object.entries(d.kota).sort((a,b)=>b[1]-a[1]).slice(0,3).map(([k,v])=>`${k} (${v})`).join(", ") }))
      .sort((a,b) => b.pesanan - a.pesanan);

    const byKota = Object.entries(kotaMap)
      .map(([name, d]) => ({ name, pesanan: d.pesanan, pcs: d.pcs, provinsi: d.provinsi }))
      .sort((a,b) => b.pesanan - a.pesanan);

    // Highlights
    const topProv = byProvinsi[0];
    const topKota = byKota[0];
    const jawa = byProvinsi.filter(p => p.name.includes("JAWA") || p.name.includes("DKI") || p.name.includes("BANTEN") || p.name.includes("YOGYAKARTA"));
    const jawaPesanan = jawa.reduce((t, p) => t + p.pesanan, 0);
    const luarJawa = byProvinsi.filter(p => !jawa.find(j => j.name === p.name));
    const luarJawaPesanan = luarJawa.reduce((t, p) => t + p.pesanan, 0);
    const totalProvCount = byProvinsi.length;

    const highlights = [
      { icon: "🏆", label: "Provinsi terbesar", value: topProv?.name, sub: `${topProv?.pesanan} pesanan · ${((topProv?.pesanan||0)/total*100).toFixed(0)}% dari total` },
      { icon: "📍", label: "Kota terbesar", value: topKota?.name, sub: `${topKota?.pesanan} pesanan · ${topKota?.provinsi}` },
      { icon: "🗺️", label: "Jangkauan", value: `${totalProvCount} provinsi`, sub: `${byKota.length} kota/kabupaten terjangkau` },
      { icon: "☕", label: "Jawa vs Luar Jawa", value: `${((jawaPesanan/total)*100).toFixed(0)}% Jawa`, sub: `Luar Jawa ${((luarJawaPesanan/total)*100).toFixed(0)}% — ${luarJawaPesanan} pesanan` },
    ];

    return { byProvinsi, byKota, highlights };
  }, [snap]);

  const data = view === "provinsi" ? byProvinsi : byKota;
  const sorted = [...data].sort((a,b) => b[sortBy] - a[sortBy]);
  const maxVal = sorted[0]?.[sortBy] || 1;

  return (
    <div>
      <SectionLabel>SEBARAN AREA · {snap.periodStart} → {snap.periodEnd} · {snap.orders.length} pesanan aktif</SectionLabel>

      {/* Highlights */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 10, marginBottom: 16 }}>
        {highlights.map((h, i) => (
          <div key={i} style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: "14px 16px" }}>
            <div style={{ fontSize: 20, marginBottom: 8 }}>{h.icon}</div>
            <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1, textTransform: "uppercase", marginBottom: 4 }}>{h.label}</div>
            <div style={{ fontFamily: mono, fontSize: 15, fontWeight: 700, color: C.ink, marginBottom: 4 }}>{h.value}</div>
            <div style={{ fontFamily: sans, fontSize: 11.5, color: C.muted }}>{h.sub}</div>
          </div>
        ))}
      </div>

      {/* Controls */}
      <div style={{ display: "flex", gap: 8, marginBottom: 12, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 4 }}>
          {[["provinsi","Provinsi"],["kota","Kota/Kabupaten"]].map(([k,lbl]) => (
            <button key={k} onClick={() => setView(k)}
              style={{ fontFamily: mono, fontSize: 11.5, fontWeight: 700, padding: "6px 12px", borderRadius: 6, cursor: "pointer",
                border: `1px solid ${view===k ? C.accent : C.line}`,
                background: view===k ? C.accent+"1a" : C.panel,
                color: view===k ? C.accent : C.muted }}>
              {lbl}
            </button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 4, marginLeft: 8 }}>
          <span style={{ fontFamily: mono, fontSize: 10.5, color: C.dim, alignSelf: "center" }}>Sort:</span>
          {[["pesanan","Pesanan"],["pcs","PCS"]].map(([k,lbl]) => (
            <button key={k} onClick={() => setSortBy(k)}
              style={{ fontFamily: mono, fontSize: 11, padding: "4px 10px", borderRadius: 5, cursor: "pointer",
                border: `1px solid ${sortBy===k ? C.good : C.line}`,
                background: sortBy===k ? C.good+"1a" : "transparent",
                color: sortBy===k ? C.good : C.muted }}>
              {lbl}
            </button>
          ))}
        </div>
      </div>

      {/* Bar chart list */}
      <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, overflow: "hidden" }}>
        {/* Header */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 80px 60px", gap: 8, padding: "10px 16px", borderBottom: `1px solid ${C.line}`, background: C.panel2 }}>
          <span style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1 }}>{view === "provinsi" ? "PROVINSI" : "KOTA / KABUPATEN"}</span>
          <span style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1, textAlign: "right" }}>PESANAN</span>
          <span style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1, textAlign: "right" }}>PCS</span>
        </div>
        {sorted.slice(0, view === "provinsi" ? 30 : 40).map((row, i) => {
          const pct = row[sortBy] / maxVal * 100;
          const isTop3 = i < 3;
          return (
            <div key={i} style={{ padding: "11px 16px", borderBottom: `1px solid ${C.line}22`, background: i % 2 === 0 ? "transparent" : C.panel2+"44" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 80px 60px", gap: 8, alignItems: "center", marginBottom: 5 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  {isTop3 && <span style={{ fontFamily: mono, fontSize: 10, fontWeight: 700, color: [C.accent, C.good, C.watch][i], minWidth: 14 }}>{["①","②","③"][i]}</span>}
                  <div>
                    <div style={{ fontFamily: sans, fontSize: 13, fontWeight: isTop3 ? 700 : 400, color: isTop3 ? C.ink : C.muted }}>
                      {row.name}
                    </div>
                    {view === "provinsi" && row.topKota && (
                      <div style={{ fontFamily: sans, fontSize: 11, color: C.dim, marginTop: 2 }}>{row.topKota}</div>
                    )}
                    {view === "kota" && row.provinsi && (
                      <div style={{ fontFamily: mono, fontSize: 10, color: C.dim, marginTop: 1 }}>{row.provinsi}</div>
                    )}
                  </div>
                </div>
                <span style={{ fontFamily: mono, fontSize: 13, fontWeight: 600, color: isTop3 ? C.accent : C.ink, textAlign: "right" }}>{row.pesanan}</span>
                <span style={{ fontFamily: mono, fontSize: 12, color: C.muted, textAlign: "right" }}>{row.pcs}</span>
              </div>
              <div style={{ height: 4, background: C.panel2, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ width: pct+"%", height: "100%", background: isTop3 ? [C.accent, C.good, C.watch][i] : C.line, borderRadius: 2, transition: "width 0.3s" }} />
              </div>
            </div>
          );
        })}
        {sorted.length === 0 && (
          <div style={{ padding: 32, textAlign: "center", color: C.dim, fontFamily: sans }}>Tidak ada data.</div>
        )}
      </div>
    </div>
  );
}


/* ============================================================================
   UNIT ECONOMICS + CONTRIBUTION MARGIN TABS
   ========================================================================== */

// Helper: match cogsItems to a product name/SKU
function matchCogs(cogsItems, name, sku) {
  if (!cogsItems.length) return null;
  // 1. exact SKU match
  const exact = cogsItems.find(c => c.sku && sku && c.sku.trim().toUpperCase() === sku.trim().toUpperCase());
  if (exact) return exact;
  // 2. SKU prefix match (SKU starts with cogs sku or vice versa)
  const prefix = cogsItems.find(c => c.sku && sku && (
    sku.toUpperCase().startsWith(c.sku.toUpperCase()) ||
    c.sku.toUpperCase().startsWith(sku.toUpperCase())
  ));
  if (prefix) return prefix;
  // 3. name contains match
  const nameLower = (name||"").toLowerCase();
  const nameMatch = cogsItems.find(c => {
    const cn = (c.nama||c.sku||"").toLowerCase();
    return cn && nameLower && (nameLower.includes(cn) || cn.includes(nameLower));
  });
  return nameMatch || null;
}

function COGSEmptyPrompt({ onDownload, onImport }) {
  return (
    <div style={{ background: C.watch+"12", border: `1px solid ${C.watch}44`, borderRadius: 8, padding: "16px 20px", marginBottom: 16 }}>
      <div style={{ fontFamily: mono, fontSize: 11.5, fontWeight: 700, color: C.watch, marginBottom: 8 }}>⚠ Data COGS belum diimport</div>
      <div style={{ fontFamily: sans, fontSize: 13, color: C.ink, marginBottom: 12, lineHeight: 1.6 }}>
        Tab ini butuh data HPP per SKU. Download template, isi HPP produksi + packaging per SKU, lalu import balik.
      </div>
      <div style={{ display: "flex", gap: 10 }}>
        <button style={{ ...S.importBtnBig, marginTop: 0, fontSize: 12, padding: "8px 16px" }} onClick={onDownload}>⬇ Download Template COGS</button>
        <button style={{ ...S.importBtn2Big, marginTop: 0, fontSize: 12, padding: "8px 16px" }} onClick={onImport}>+ Import COGS</button>
      </div>
    </div>
  );
}

function UnitEconomicsTab({ active, activeProducts, cogsItems, shopeeSnap, tiktokSnap, onDownloadTemplate, onImportCogs }) {
  const hasCogs = cogsItems.length > 0;
  const feeRate = useMemo(() => {
    const sh = shopeeSnap?.summary?.feeRateNetGmv || 0;
    const tt = tiktokSnap?.summary?.feeRateNetGmv || 0;
    if (sh && tt) return (sh + tt) / 2;
    return sh || tt || 24; // fallback 24%
  }, [shopeeSnap, tiktokSnap]);

  const blendedAdRate = useMemo(() => {
    if (!active) return 0;
    const totalSpend = active.ads.reduce((t,a) => t+a.spend, 0);
    const totalGmv   = active.ads.reduce((t,a) => t+a.gmv, 0);
    return totalGmv ? totalSpend/totalGmv*100 : 0;
  }, [active]);

  const rows = useMemo(() => {
    if (!active) return [];
    const seen = new Map();
    active.ads.forEach(ad => {
      const k = ad.code && ad.code !== "-" ? ad.code : ad.name;
      if (!seen.has(k)) seen.set(k, { sku: ad.code, name: ad.name, gmv: 0, spend: 0, units: 0 });
      const r = seen.get(k);
      r.gmv += ad.gmv; r.spend += ad.spend; r.units += (ad.conv || 0);
    });
    return [...seen.values()].map(r => {
      const cogs = matchCogs(cogsItems, r.name, r.sku);
      const hpp = cogs?.total || 0;
      const harga = cogs?.harga || (r.units ? r.gmv/r.units : 0);
      const feeAmt = harga * (feeRate/100);
      const adAmt  = harga * (blendedAdRate/100);
      const grossMargin = harga - hpp;
      const grossMarginPct = harga ? grossMargin/harga*100 : 0;
      const netMargin = grossMargin - feeAmt - adAmt;
      const netMarginPct = harga ? netMargin/harga*100 : 0;
      return { ...r, hpp, harga, feeAmt, adAmt, grossMargin, grossMarginPct, netMargin, netMarginPct, hasCogs: !!cogs };
    }).sort((a,b) => b.gmv - a.gmv);
  }, [active, cogsItems, feeRate, blendedAdRate]);

  return (
    <div>
      <SectionLabel>UNIT ECONOMICS PER SKU</SectionLabel>
      {!hasCogs && <COGSEmptyPrompt onDownload={onDownloadTemplate} onImport={onImportCogs} />}
      <div style={{ ...S.invIntro, marginBottom: 12 }}>
        Fee rate dipakai: <b style={{ color: C.bad }}>{feeRate.toFixed(1)}%</b>
        {shopeeSnap || tiktokSnap ? " (dari income file)" : " (default)"} ·
        Blended ad rate: <b style={{ color: C.watch }}>{blendedAdRate.toFixed(1)}%</b>
        {active ? " (dari file iklan)" : ""} ·
        HPP dari template COGS {hasCogs ? <span style={{ color: C.good }}>✓ {cogsItems.length} SKU</span> : <span style={{ color: C.dim }}>belum ada</span>}
      </div>
      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead><tr>
            <th style={S.th}>Produk</th>
            <th style={{ ...S.th, textAlign: "right" }}>Harga Jual</th>
            <th style={{ ...S.th, textAlign: "right" }}>HPP</th>
            <th style={{ ...S.th, textAlign: "right" }}>Fee Mkt</th>
            <th style={{ ...S.th, textAlign: "right" }}>Ad Cost</th>
            <th style={{ ...S.th, textAlign: "right" }}>Gross Margin</th>
            <th style={{ ...S.th, textAlign: "right" }}>Net Margin</th>
            <th style={{ ...S.th, textAlign: "right" }}>Net %</th>
          </tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} style={S.tr}>
                <td style={S.td}>
                  <div style={S.adName}>{r.name}</div>
                  {r.sku && r.sku !== "-" && <div style={{ fontFamily: mono, fontSize: 10, color: C.dim }}>{r.sku}</div>}
                  {!r.hasCogs && hasCogs && <div style={{ fontFamily: mono, fontSize: 10, color: C.watch }}>⚠ SKU tidak match</div>}
                </td>
                <td style={S.tdR}>{r.harga ? rpShort(r.harga) : <span style={{ color: C.dim }}>-</span>}</td>
                <td style={{ ...S.tdR, color: r.hasCogs ? C.ink : C.dim }}>{r.hpp ? rpShort(r.hpp) : "-"}</td>
                <td style={{ ...S.tdR, color: C.bad }}>{r.harga ? rpShort(r.feeAmt) : "-"}</td>
                <td style={{ ...S.tdR, color: C.watch }}>{r.harga ? rpShort(r.adAmt) : "-"}</td>
                <td style={{ ...S.tdR, color: r.grossMarginPct > 30 ? C.good : r.grossMarginPct > 15 ? C.watch : C.bad }}>
                  {r.hasCogs && r.harga ? `${r.grossMarginPct.toFixed(0)}%` : "-"}
                </td>
                <td style={{ ...S.tdR, color: r.netMarginPct > 10 ? C.good : r.netMarginPct > 0 ? C.watch : C.bad }}>
                  {r.hasCogs && r.harga ? rpShort(r.netMargin) : "-"}
                </td>
                <td style={{ ...S.tdR, fontWeight: 700, color: r.netMarginPct > 10 ? C.good : r.netMarginPct > 0 ? C.watch : C.bad }}>
                  {r.hasCogs && r.harga ? `${r.netMarginPct.toFixed(1)}%` : "-"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ContribMarginTab({ active, activeProducts, cogsItems, shopeeSnap, tiktokSnap, onDownloadTemplate, onImportCogs }) {
  const hasCogs = cogsItems.length > 0;
  const feeRate = useMemo(() => {
    const sh = shopeeSnap?.summary?.feeRateNetGmv || 0;
    const tt = tiktokSnap?.summary?.feeRateNetGmv || 0;
    if (sh && tt) return (sh+tt)/2;
    return sh || tt || 24;
  }, [shopeeSnap, tiktokSnap]);

  const blendedAdRate = useMemo(() => {
    if (!active) return 0;
    const totalSpend = active.ads.reduce((t,a) => t+a.spend,0);
    const totalGmv   = active.ads.reduce((t,a) => t+a.gmv,0);
    return totalGmv ? totalSpend/totalGmv*100 : 0;
  }, [active]);

  const rows = useMemo(() => {
    if (!active) return [];
    const seen = new Map();
    active.ads.forEach(ad => {
      const k = ad.code && ad.code !== "-" ? ad.code : ad.name;
      if (!seen.has(k)) seen.set(k, { sku: ad.code, name: ad.name, gmv: 0, spend: 0, units: 0 });
      const r = seen.get(k);
      r.gmv += ad.gmv; r.spend += ad.spend; r.units += (ad.conv||0);
    });
    return [...seen.values()].map(r => {
      const cogs = matchCogs(cogsItems, r.name, r.sku);
      const hppUnit = cogs?.total || 0;
      const units = r.units || (r.gmv && cogs?.harga ? r.gmv/cogs.harga : 0);
      const totalHPP  = hppUnit * units;
      const totalFee  = r.gmv * (feeRate/100);
      const totalAd   = r.gmv * (blendedAdRate/100);
      const cm        = r.gmv - totalHPP - totalFee - totalAd;
      const cmPct     = r.gmv ? cm/r.gmv*100 : 0;
      return { ...r, hppUnit, units, totalHPP, totalFee, totalAd, cm, cmPct, hasCogs: !!cogs };
    }).filter(r => r.gmv > 0).sort((a,b) => b.cm - a.cm);
  }, [active, cogsItems, feeRate, blendedAdRate]);

  const totalGmv = rows.reduce((t,r) => t+r.gmv, 0);
  const totalCM  = rows.reduce((t,r) => t+r.cm, 0);
  const blendedCM = totalGmv ? totalCM/totalGmv*100 : 0;

  return (
    <div>
      <SectionLabel>CONTRIBUTION MARGIN PER SKU</SectionLabel>
      {!hasCogs && <COGSEmptyPrompt onDownload={onDownloadTemplate} onImport={onImportCogs} />}

      {/* Summary KPIs */}
      {hasCogs && rows.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10, marginBottom: 16 }}>
          <Kpi label="Total GMV" value={rpShort(totalGmv)} dir="rev" accent={C.accent} />
          <Kpi label="Total CM" value={rpShort(totalCM)} dir={totalCM > 0 ? "rev" : "cost"} accent={totalCM > 0 ? C.good : C.bad} big />
          <Kpi label="Blended CM Rate" value={blendedCM.toFixed(1)+"%"} dir={blendedCM > 0 ? "rev" : "cost"} accent={blendedCM > 15 ? C.good : blendedCM > 5 ? C.watch : C.bad} big />
          <Kpi label="SKU Profitable" value={rows.filter(r=>r.cm>0).length+"/"+rows.length} dir="neutral" accent={C.good} />
          <Kpi label="Fee Rate" value={feeRate.toFixed(1)+"%"} dir="cost" accent={C.bad} />
          <Kpi label="Blended Ad Rate" value={blendedAdRate.toFixed(1)+"%"} dir="cost" accent={C.watch} />
        </div>
      )}

      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead><tr>
            <th style={S.th}>Produk</th>
            <th style={{ ...S.th, textAlign: "right" }}>GMV</th>
            <th style={{ ...S.th, textAlign: "right" }}>HPP Total</th>
            <th style={{ ...S.th, textAlign: "right" }}>Fee Mkt</th>
            <th style={{ ...S.th, textAlign: "right" }}>Ad Spend</th>
            <th style={{ ...S.th, textAlign: "right" }}>CM (Rp)</th>
            <th style={{ ...S.th, textAlign: "right" }}>CM %</th>
          </tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} style={{ ...S.tr, background: r.cm < 0 ? C.bad+"08" : "transparent" }}>
                <td style={S.td}>
                  <div style={S.adName}>{r.name}</div>
                  {!r.hasCogs && hasCogs && <div style={{ fontFamily: mono, fontSize: 10, color: C.watch }}>⚠ SKU tidak match</div>}
                </td>
                <td style={S.tdR}>{rpShort(r.gmv)}</td>
                <td style={{ ...S.tdR, color: r.hasCogs ? C.ink : C.dim }}>{r.hasCogs ? rpShort(r.totalHPP) : "-"}</td>
                <td style={{ ...S.tdR, color: C.bad }}>{rpShort(r.totalFee)}</td>
                <td style={{ ...S.tdR, color: C.watch }}>{rpShort(r.totalAd)}</td>
                <td style={{ ...S.tdR, fontWeight: 700, color: r.cm > 0 ? C.good : C.bad }}>
                  {r.hasCogs ? rpShort(r.cm) : "-"}
                </td>
                <td style={{ ...S.tdR, fontWeight: 700, color: r.cmPct > 15 ? C.good : r.cmPct > 0 ? C.watch : C.bad }}>
                  {r.hasCogs ? r.cmPct.toFixed(1)+"%" : "-"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!active && <div style={S.cmEmpty}>Butuh file iklan untuk kalkulasi GMV dan ad spend per SKU.</div>}
    </div>
  );
}


/* ============================================================================
   META ADS VIEW + ALL CHANNEL VIEW
   ========================================================================== */
function MetaAdsView({ snap }) {
  const s = snap.summary;
  const [sort, setSort] = useState("spend");
  const [dir, setDir] = useState(-1);
  const [filterObj, setFilterObj] = useState("all");
  const [openIdx, setOpenIdx] = useState(null);

  // Enrich campaigns with objective classification
  const enriched = useMemo(() => snap.campaigns.map(c => ({
    ...c, ...diagnoseMetaCampaign(c)
  })), [snap]);

  // Objective counts for filter tabs
  const objCounts = useMemo(() => {
    const counts = { all: enriched.length };
    enriched.forEach(c => { counts[c.obj] = (counts[c.obj] || 0) + 1; });
    return counts;
  }, [enriched]);

  const filtered = useMemo(() => {
    let r = filterObj === "all" ? enriched : enriched.filter(c => c.obj === filterObj);
    return [...r].sort((a,b) => (a[sort]<b[sort]?1:-1)*dir);
  }, [enriched, filterObj, sort, dir]);

  // Spend by objective
  const spendByObj = useMemo(() => {
    const m = {};
    enriched.forEach(c => { m[c.obj] = (m[c.obj] || 0) + c.spend; });
    return m;
  }, [enriched]);

  const head = (k, lbl, right) => (
    <th style={{ ...S.th, textAlign: right ? "right" : "left", cursor: "pointer" }}
      onClick={() => { if (sort===k) setDir(-dir); else { setSort(k); setDir(-1); } }}>
      {lbl}{sort===k ? (dir===-1?" ▾":" ▴") : ""}
    </th>
  );

  return (
    <div>
      <SectionLabel>META ADS · {snap.periodStart} → {snap.periodEnd} · {snap.campaigns.length} campaign</SectionLabel>

      {/* KPI summary */}
      <div style={S.kpiGrid}>
        <Kpi label="Total Spend" value={rpShort(s.totalSpend)} dir="cost" accent={C.watch} />
        <Kpi label="Blended ROAS (Konversi)" value={s.blendedRoas.toFixed(2)+"x"} dir="rev" big accent={s.blendedRoas >= 3 ? C.good : C.bad} />
        <Kpi label="Total Results" value={nfmt(s.totalResult)} dir="rev" accent={C.good} />
        <Kpi label="Total Impressi" value={nfmt(s.totalImpr)} dir="neutral" accent={C.muted} />
        <Kpi label="Blended CTR" value={s.blendedCtr.toFixed(2)+"%"} dir="neutral" accent={C.muted} />
        <Kpi label="Cost per Result" value={rpShort(s.blendedCpr)} dir="cost" accent={C.watch} />
      </div>

      {/* Spend allocation by objective */}
      <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: "14px 16px", marginTop: 12 }}>
        <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1, marginBottom: 10 }}>ALOKASI SPEND PER OBJECTIVE</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {Object.entries(META_OBJ).filter(([k]) => spendByObj[k]).map(([k, obj]) => (
            <div key={k} style={{ background: obj.color+"15", border: `1px solid ${obj.color}44`, borderRadius: 6, padding: "8px 12px", minWidth: 130 }}>
              <div style={{ fontFamily: mono, fontSize: 10, color: obj.color, marginBottom: 4 }}>{obj.icon} {obj.label.toUpperCase()}</div>
              <div style={{ fontFamily: mono, fontSize: 14, fontWeight: 700, color: C.ink }}>{rpShort(spendByObj[k] || 0)}</div>
              <div style={{ fontFamily: sans, fontSize: 11, color: C.muted }}>{s.totalSpend ? ((spendByObj[k]||0)/s.totalSpend*100).toFixed(0) : 0}% dari total spend</div>
            </div>
          ))}
        </div>
      </div>

      {/* Attribution note */}
      <div style={{ background: C.watch+"10", border: `1px solid ${C.watch}33`, borderRadius: 7, padding: "10px 14px", marginTop: 10, fontFamily: sans, fontSize: 12.5, color: C.ink, lineHeight: 1.6 }}>
        <b style={{ color: C.watch }}>Penting — beda metodologi ROAS:</b> ROAS Meta pakai atribusi 7-day click / 1-day view.
        ROAS Shopee pakai last-click dalam platform. Jangan compare langsung — gunakan total spend + total revenue aktual dari income file sebagai basis perbandingan efisiensi channel.
      </div>

      {/* Objective filter */}
      <div style={{ display: "flex", gap: 6, marginTop: 14, marginBottom: 10, flexWrap: "wrap" }}>
        {[["all","Semua"], ...Object.entries(META_OBJ).filter(([k]) => objCounts[k]).map(([k,v]) => [k, v.label])].map(([k, lbl]) => (
          <button key={k} onClick={() => setFilterObj(k)}
            style={{ fontFamily: mono, fontSize: 11, padding: "5px 12px", borderRadius: 5, cursor: "pointer",
              border: `1px solid ${filterObj===k ? (META_OBJ[k]?.color || C.accent) : C.line}`,
              background: filterObj===k ? (META_OBJ[k]?.color || C.accent)+"1a" : "transparent",
              color: filterObj===k ? (META_OBJ[k]?.color || C.accent) : C.muted }}>
            {k !== "all" && META_OBJ[k]?.icon + " "}{lbl} {objCounts[k] ? `(${objCounts[k]})` : ""}
          </button>
        ))}
      </div>

      {/* Campaign cards with diagnosis */}
      {filtered.map((c, i) => {
        const obj = META_OBJ[c.obj] || META_OBJ.other;
        const isOpen = openIdx === i;
        const worstLvl = c.items.reduce((w, it) => it.level==="bad" ? "bad" : w==="bad" ? "bad" : it.level==="watch" ? "watch" : w, "good");
        const borderCol = worstLvl === "bad" ? C.bad : worstLvl === "watch" ? C.watch : C.good;
        return (
          <div key={i} style={{ ...S.skuCard, borderLeftColor: obj.color, marginBottom: 8 }}>
            <div style={S.skuHead} onClick={() => setOpenIdx(isOpen ? null : i)}>
              <div style={S.skuHeadL}>
                <span style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: obj.color, background: obj.color+"15", border: `1px solid ${obj.color}44`, padding: "2px 8px", borderRadius: 4 }}>
                  {obj.icon} {obj.label}
                </span>
                <span style={S.skuName}>{c.name}</span>
              </div>
              <div style={S.skuHeadR}>
                <span style={S.skuStat}>spend <b>{rpShort(c.spend)}</b></span>
                {c.roas > 0 && <span style={{ ...S.skuStat, color: c.roas >= 3 ? C.good : C.watch }}>ROAS <b>{c.roas.toFixed(2)}x</b></span>}
                <span style={S.skuStat}>CTR <b>{c.ctr.toFixed(2)}%</b></span>
                <span style={S.skuStat}>{nfmt(c.results)} results</span>
                <span style={{ color: C.muted, fontFamily: mono }}>{isOpen ? "▾" : "▸"}</span>
              </div>
            </div>
            {isOpen && (
              <div style={{ padding: "12px 16px", borderTop: `1px solid ${C.line}` }}>
                {/* Metrics row */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(130px, 1fr))", gap: 8, marginBottom: 14 }}>
                  {[
                    ["Spend", rpShort(c.spend), C.watch],
                    ["Results", nfmt(c.results), C.good],
                    ["Impressi", nfmt(c.impressions), C.muted],
                    ["CTR", c.ctr.toFixed(2)+"%", C.ink],
                    ["CPC", rpShort(c.cpc), C.muted],
                    ["Cost/Result", rpShort(c.cpr), C.watch],
                    ...(c.roas > 0 ? [["ROAS", c.roas.toFixed(2)+"x", c.roas >= 3 ? C.good : C.bad]] : []),
                    ...(c.atc > 0 ? [["Add to Cart", nfmt(c.atc), C.accent]] : []),
                    ...(c.reach > 0 ? [["Reach", nfmt(c.reach), C.muted]] : []),
                  ].map(([lbl, val, col], j) => (
                    <div key={j} style={{ background: C.panel2, borderRadius: 6, padding: "8px 10px" }}>
                      <div style={{ fontFamily: mono, fontSize: 9.5, color: C.dim, letterSpacing: 0.8, marginBottom: 3 }}>{lbl}</div>
                      <div style={{ fontFamily: mono, fontSize: 13, fontWeight: 700, color: col }}>{val}</div>
                    </div>
                  ))}
                </div>
                {/* Diagnosis items */}
                <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1, marginBottom: 8 }}>ANALISA · {obj.label.toUpperCase()}</div>
                {c.items.map((it, j) => (
                  <div key={j} style={{ borderLeft: `3px solid ${it.level==="bad" ? C.bad : it.level==="good" ? C.good : it.level==="watch" ? C.watch : C.muted}`, background: (it.level==="bad" ? C.bad : it.level==="good" ? C.good : it.level==="watch" ? C.watch : C.muted)+"0a", borderRadius: "0 6px 6px 0", padding: "10px 14px", marginBottom: 8 }}>
                    <div style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: it.level==="bad" ? C.bad : it.level==="good" ? C.good : it.level==="watch" ? C.watch : C.muted, marginBottom: 5 }}>{it.tag}</div>
                    <div style={{ fontFamily: sans, fontSize: 13, color: C.ink, lineHeight: 1.6 }}>{it.msg}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function AllChannelView({ active, metaSnap, tiktokAdsSnap }) {
  const shopeeSpend  = active ? active.ads.reduce((t,a) => t+a.spend, 0) : 0;
  const shopeeGmv    = active ? active.ads.reduce((t,a) => t+a.gmv, 0) : 0;
  const shopeeConv   = active ? active.ads.reduce((t,a) => t+(a.conv||0), 0) : 0;
  const shopeeRoas   = shopeeSpend > 0 ? shopeeGmv/shopeeSpend : 0;
  const metaSpend    = metaSnap?.summary?.totalSpend || 0;
  const metaResults  = metaSnap?.summary?.totalResult || 0;
  const metaRoas     = metaSnap?.summary?.blendedRoas || 0;
  const metaImpr     = metaSnap?.summary?.totalImpr || 0;
  const ttAdsSpend   = tiktokAdsSnap?.summary?.totalSpend || 0;
  const ttAdsGmv     = tiktokAdsSnap?.summary?.totalGmv || 0;
  const ttAdsOrders  = tiktokAdsSnap?.summary?.totalOrders || 0;
  const ttAdsRoi     = tiktokAdsSnap?.summary?.blendedRoi || 0;
  const totalSpend   = shopeeSpend + metaSpend + ttAdsSpend;
  const channels     = [
    { name: "Shopee Ads", spend: shopeeSpend, gmv: shopeeGmv, conv: shopeeConv, roas: shopeeRoas, color: C.accent },
    ...(metaSnap ? [{ name: "Meta Ads", spend: metaSpend, gmv: 0, conv: metaResults, roas: metaRoas, color: "#1877f2" }] : []),
    ...(tiktokAdsSnap ? [{ name: "TikTok GMV Max", spend: ttAdsSpend, gmv: ttAdsGmv, conv: ttAdsOrders, roas: ttAdsRoi, color: "#fe2c55" }] : []),
  ];

  return (
    <div>
      <SectionLabel>SEMUA CHANNEL · RINGKASAN SPEND</SectionLabel>
      <div style={S.kpiGrid}>
        <Kpi label="Total Spend Semua Channel" value={rpShort(totalSpend)} dir="cost" big accent={C.watch} />
        <Kpi label="Shopee Ads Spend" value={rpShort(shopeeSpend)} dir="cost" accent={C.accent} />
        {metaSnap && <Kpi label="Meta Ads Spend" value={rpShort(metaSpend)} dir="cost" accent={"#1877f2"} />}
        <Kpi label="Shopee ROAS" value={shopeeRoas.toFixed(2)+"x"} dir="rev" accent={shopeeRoas >= 5 ? C.good : C.watch} />
        {metaSnap && <Kpi label="Meta Blended ROAS" value={metaRoas.toFixed(2)+"x"} dir="rev" accent={metaRoas >= 3 ? C.good : C.watch} />}
        {metaSnap && <Kpi label="Meta Impressi" value={nfmt(metaImpr)} dir="neutral" accent={C.muted} />}
        {tiktokAdsSnap && <Kpi label="TikTok GMV Max ROI" value={ttAdsRoi.toFixed(2)+"x"} dir="rev" accent={ttAdsRoi >= 5 ? C.good : C.watch} />}
        {tiktokAdsSnap && <Kpi label="TikTok GMV" value={rpShort(ttAdsGmv)} dir="rev" accent={"#fe2c55"} />}
      </div>

      {/* Spend allocation bar */}
      {totalSpend > 0 && (
        <div style={{ marginTop: 16, background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: "14px 16px" }}>
          <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1, marginBottom: 12 }}>ALOKASI SPEND PER CHANNEL</div>
          {channels.map((ch, i) => {
            const pct = totalSpend ? ch.spend/totalSpend*100 : 0;
            return (
              <div key={i} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                  <span style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, color: ch.color }}>{ch.name}</span>
                  <span style={{ fontFamily: mono, fontSize: 13, color: C.ink }}>{rpShort(ch.spend)} <span style={{ color: C.muted }}>({pct.toFixed(0)}%)</span></span>
                </div>
                <div style={{ height: 8, background: C.panel2, borderRadius: 4, overflow: "hidden" }}>
                  <div style={{ width: pct+"%", height: "100%", background: ch.color, borderRadius: 4 }} />
                </div>
                <div style={{ fontFamily: mono, fontSize: 11, color: C.muted, marginTop: 4 }}>
                  ROAS {ch.roas.toFixed(2)}x · {ch.name === "Shopee Ads" ? `${nfmt(ch.conv)} konversi` : `${nfmt(ch.conv)} results`}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {!metaSnap && (
        <div style={{ ...S.invIntro, marginTop: 12 }}>
          Import <b>Meta Ads CSV</b> via tombol + Import Data untuk melihat perbandingan lintas channel.
        </div>
      )}
    </div>
  );
}


/* ============================================================================
   TIKTOK GMV MAX VIEW
   ========================================================================== */
function TikTokAdsView({ snap, campaignSnap, onImportCampaign }) {
  const s = snap.summary;
  const daily = snap.daily.filter(d => d.tgl && d.tgl !== "-");
  const [openCamp, setOpenCamp] = useState(null);

  // ROI trend — 7-day moving avg
  const trend = useMemo(() => {
    return daily.map((d, i) => {
      const window = daily.slice(Math.max(0, i-6), i+1);
      const avgRoi = window.reduce((t,w) => t+w.roi, 0) / window.length;
      return { ...d, avgRoi };
    });
  }, [daily]);

  // Weekly summary
  const weeks = useMemo(() => {
    const wks = {};
    daily.forEach(d => {
      const dt = new Date(d.tgl);
      if (isNaN(dt)) return;
      const wk = `W${Math.ceil(dt.getDate()/7)} (${dt.getDate() < 8 ? "1–7" : dt.getDate() < 15 ? "8–14" : dt.getDate() < 22 ? "15–21" : dt.getDate() < 29 ? "22–28" : "29–31"})`;
      if (!wks[wk]) wks[wk] = { spend:0, gmv:0, orders:0, days:0, roiSum:0 };
      wks[wk].spend += d.spend; wks[wk].gmv += d.gmv;
      wks[wk].orders += d.orders; wks[wk].days++;
      wks[wk].roiSum += d.roi;
    });
    return Object.entries(wks).map(([wk, v]) => ({
      wk, ...v, roi: v.spend > 0 ? v.gmv/v.spend : 0, avgRoi: v.roiSum/Math.max(v.days,1)
    }));
  }, [daily]);

  // Best/worst days
  const sorted = [...daily].filter(d => d.roi > 0).sort((a,b) => b.roi - a.roi);
  const best3  = sorted.slice(0,3);
  const worst3 = sorted.slice(-3).reverse();

  // Diagnosa — gabungan daily + campaign file
  const diagnosa = useMemo(() => {
    const items = [];
    const roi = s.blendedRoi;
    const roiVariance = daily.filter(d=>d.roi>0).reduce((t,d) => t + Math.pow(d.roi - s.avgRoi, 2), 0) / Math.max(daily.filter(d=>d.roi>0).length-1, 1);
    const roiStdDev = Math.sqrt(roiVariance);

    // ── Daily: ROI overall ──
    if (roi >= 8) {
      items.push({ level: "good", tag: "ROI EXCELLENT", msg: `Blended ROI ${roi.toFixed(2)}x — sangat efisien. Di atas 8x berarti setiap Rp 1 spend → Rp ${roi.toFixed(1)} GMV. Pertahankan budget dan pantau tanda diminishing return.` });
    } else if (roi >= 5) {
      items.push({ level: "good", tag: "ROI BAIK", msg: `Blended ROI ${roi.toFixed(2)}x — healthy (benchmark fashion 4–8x). Masih ada ruang scale 20–30% sambil pantau ROI tidak drop di bawah 4x.` });
    } else if (roi >= 3) {
      items.push({ level: "watch", tag: "ROI MODERAT", msg: `Blended ROI ${roi.toFixed(2)}x — acceptable tapi belum optimal. Target minimal 4–5x. Cek kelengkapan katalog dan apakah produk margin tinggi sudah aktif.` });
    } else {
      items.push({ level: "bad", tag: "ROI RENDAH", msg: `Blended ROI ${roi.toFixed(2)}x — di bawah threshold. Pertimbangkan kurangi budget GMV Max, alihkan ke manual. ROI rendah biasanya indikasi katalog perlu refresh atau harga tidak kompetitif.` });
    }

    // ── Daily: konsistensi ──
    if (roiStdDev > 2) {
      items.push({ level: "watch", tag: "ROI TIDAK STABIL", msg: `Deviasi ROI harian ${roiStdDev.toFixed(1)} — fluktuasi tinggi. Terbaik ${best3[0]?.roi.toFixed(1)}x, terburuk ${worst3[0]?.roi.toFixed(1)}x. Penyebab umum: flash sale, stok SKU habis di tengah bulan, atau algoritma dalam fase learning ulang.` });
    } else {
      items.push({ level: "good", tag: "ROI KONSISTEN", msg: `Deviasi ROI ${roiStdDev.toFixed(1)} — stabil. Algoritma GMV Max sudah matang dengan katalog dan budget sekarang.` });
    }

    if (s.avgCpa > 75000) {
      items.push({ level: "watch", tag: "CPA TINGGI", msg: `Rata-rata CPA Rp ${s.avgCpa.toLocaleString("id-ID")} — cek apakah GMV Max mendorong produk AOV tinggi (wajar) atau entry-level margin tipis (masalah). CPA ideal = di bawah 10% dari AOV.` });
    }

    // ── Cross-file: campaign insights ──
    if (campaignSnap?.campaigns?.length > 0) {
      const camps = campaignSnap.campaigns;
      const totalCampSpend = camps.reduce((t,c) => t+c.spend, 0);

      const winners  = camps.filter(c => c.roi >= roi * 1.2 && c.spend > 0);
      const draggers = camps.filter(c => c.roi > 0 && c.roi < roi * 0.6 && c.spend > totalCampSpend * 0.05);
      const topByGmv = [...camps].sort((a,b) => b.gmv - a.gmv)[0];

      if (winners.length > 0)
        items.push({ level: "good", tag: "KAMPANYE TERBAIK", msg: `${winners.length} kampanye ROI di atas rata-rata: ${winners.slice(0,3).map(c=>`"${c.name.slice(0,22)}" ${c.roi.toFixed(1)}x`).join(" · ")}. Prioritaskan stok produk di kampanye ini.` });

      if (draggers.length > 0) {
        const dragSpend = draggers.reduce((t,c) => t+c.spend, 0);
        items.push({ level: "bad", tag: "KAMPANYE DRAG ROI", msg: `${draggers.length} kampanye ROI jauh di bawah rata-rata: ${draggers.slice(0,2).map(c=>`"${c.name.slice(0,22)}" ${c.roi.toFixed(1)}x`).join(" · ")} — menyerap ${rpShort(dragSpend)} spend. Ketatkan target ROI atau pause untuk naikkan blended ROI.` });
      }

      if (topByGmv) {
        const topShare = topByGmv.gmv / (campaignSnap.summary.totalGmv || 1) * 100;
        if (topShare > 60)
          items.push({ level: "watch", tag: "KONSENTRASI GMV TINGGI", msg: `Kampanye "${topByGmv.name.slice(0,30)}" menghasilkan ${topShare.toFixed(0)}% GMV TikTok. Dependensi tinggi — kalau kampanye ini drop, impact besar. Kembangkan kampanye lain sebagai backup.` });
      }

      const lowCtr = camps.filter(c => c.ctr > 0 && c.ctr < 1 && c.spend > totalCampSpend * 0.05);
      if (lowCtr.length > 0)
        items.push({ level: "watch", tag: "CTR RENDAH", msg: `${lowCtr.length} kampanye CTR di bawah 1%: ${lowCtr.slice(0,2).map(c=>`"${c.name.slice(0,20)}" ${c.ctr.toFixed(2)}%`).join(", ")}. Refresh visual atau hook opening video.` });

      const lowCvr = camps.filter(c => c.cvr > 0 && c.cvr < 1 && c.spend > totalCampSpend * 0.05);
      if (lowCvr.length > 0)
        items.push({ level: "watch", tag: "CVR RENDAH", msg: `${lowCvr.length} kampanye CVR di bawah 1%: ${lowCvr.slice(0,2).map(c=>`"${c.name.slice(0,20)}" ${c.cvr.toFixed(2)}%`).join(", ")}. Masalah listing bukan iklan — cek harga, foto, ulasan produk.` });

      // Validasi konsistensi spend antar dua file
      const gapPct = campaignSnap.summary.totalSpend > 0 && s.totalSpend > 0
        ? Math.abs(campaignSnap.summary.totalSpend - s.totalSpend) / s.totalSpend * 100 : 0;
      if (gapPct > 15)
        items.push({ level: "watch", tag: "GAP DATA ANTAR FILE", msg: `Spend GMV Max harian (${rpShort(s.totalSpend)}) vs Campaign Report (${rpShort(campaignSnap.summary.totalSpend)}) selisih ${gapPct.toFixed(0)}%. Pastikan kedua file menggunakan periode yang sama.` });
    }

    items.push({ level: "neutral", tag: "KONTEKS GMV MAX", msg: "Lever optimasi GMV Max: (1) budget harian, (2) kelengkapan & kualitas katalog, (3) harga kompetitif, (4) stok cukup agar algoritma tidak switch ke produk lain. Tidak bisa dioptimasi di level creative atau targeting." });

    return items;
  }, [s, daily, best3, worst3, campaignSnap]);

  const maxGmv = Math.max(...daily.map(d => d.gmv), 1);

  return (
    <div>
      <SectionLabel>TIKTOK GMV MAX · {snap.periodStart} → {snap.periodEnd} · {daily.length} hari</SectionLabel>

      {/* KPI */}
      <div style={S.kpiGrid}>
        <Kpi label="Total Spend" value={rpShort(s.totalSpend)} dir="cost" accent={C.watch} />
        <Kpi label="Total GMV" value={rpShort(s.totalGmv)} dir="rev" accent={C.accent} />
        <Kpi label="Blended ROI" value={s.blendedRoi.toFixed(2)+"x"} dir="rev" big accent={s.blendedRoi >= 5 ? C.good : s.blendedRoi >= 3 ? C.watch : C.bad} />
        <Kpi label="Total Orders" value={nfmt(s.totalOrders)} dir="rev" accent={C.good} />
        <Kpi label="Avg CPA" value={rpShort(s.avgCpa)} dir="cost" accent={C.watch} />
        <Kpi label="Avg Daily ROI" value={s.avgRoi.toFixed(2)+"x"} dir="rev" accent={C.muted} />
      </div>

      {/* Dual-line chart: Biaya vs Pesanan harian */}
      {daily.length > 1 && (() => {
        const W = 700, H = 160, padL = 52, padR = 16, padT = 12, padB = 28;
        const maxSpend  = Math.max(...daily.map(d => d.spend), 1);
        const maxOrders = Math.max(...daily.map(d => d.orders), 1);
        const xOf = i => padL + (i / Math.max(daily.length - 1, 1)) * (W - padL - padR);
        const ySpend  = v => padT + (1 - v / maxSpend)  * (H - padT - padB);
        const yOrders = v => padT + (1 - v / maxOrders) * (H - padT - padB);
        const spendPts  = daily.map((d,i) => `${xOf(i)},${ySpend(d.spend)}`).join(" ");
        const orderPts  = daily.map((d,i) => `${xOf(i)},${yOrders(d.orders)}`).join(" ");
        const spendArea = `${xOf(0)},${H-padB} ` + spendPts + ` ${xOf(daily.length-1)},${H-padB}`;
        // Y-axis ticks (spend)
        const yTicks = [0, 0.5, 1].map(r => ({ v: maxSpend * r, y: ySpend(maxSpend * r) }));
        // Date ticks — show ~5 evenly
        const step = Math.max(1, Math.floor(daily.length / 5));
        const dateTicks = daily.filter((_, i) => i % step === 0 || i === daily.length - 1);
        return (
          <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, padding: "16px 18px", marginTop: 12 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontFamily: mono, fontSize: 10, color: C.muted, letterSpacing: 1.2 }}>BIAYA & PESANAN HARIAN</div>
              <div style={{ display: "flex", gap: 16 }}>
                {[["#1dd3b0", "Biaya"], ["#f5a623", "Pesanan SKU"]].map(([c, l]) => (
                  <div key={l} style={{ display: "flex", alignItems: "center", gap: 5, fontFamily: mono, fontSize: 10, color: C.muted }}>
                    <div style={{ width: 20, height: 2.5, background: c, borderRadius: 2 }} />{l}
                  </div>
                ))}
              </div>
            </div>
            <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", overflow: "visible" }}>
              {/* Grid lines */}
              {yTicks.map((t, i) => (
                <g key={i}>
                  <line x1={padL} x2={W - padR} y1={t.y} y2={t.y} stroke={C.line} strokeWidth="0.5" strokeDasharray="3,3" />
                  <text x={padL - 4} y={t.y + 3.5} textAnchor="end" fontSize="8" fill={C.dim} fontFamily="monospace">{rpShort(t.v)}</text>
                </g>
              ))}
              {/* Date ticks */}
              {dateTicks.map((d, i) => {
                const idx = daily.indexOf(d);
                return <text key={i} x={xOf(idx)} y={H - 4} textAnchor="middle" fontSize="8" fill={C.dim} fontFamily="monospace">{d.tgl?.slice(5)}</text>;
              })}
              {/* Spend area fill */}
              <polygon points={spendArea} fill="#1dd3b0" fillOpacity="0.08" />
              {/* Spend line */}
              <polyline points={spendPts} fill="none" stroke="#1dd3b0" strokeWidth="1.8" strokeLinejoin="round" strokeLinecap="round" />
              {/* Orders line (right axis scale) */}
              <polyline points={orderPts} fill="none" stroke="#f5a623" strokeWidth="1.8" strokeLinejoin="round" strokeLinecap="round" strokeDasharray="0" />
              {/* Right axis ticks (orders) */}
              {[0, 0.5, 1].map((r, i) => (
                <text key={i} x={W - padR + 4} y={yOrders(maxOrders * r) + 3.5} textAnchor="start" fontSize="8" fill={C.dim} fontFamily="monospace">{Math.round(maxOrders * r)}</text>
              ))}
              {/* Hover dots on data points */}
              {daily.map((d, i) => (
                <circle key={i} cx={xOf(i)} cy={ySpend(d.spend)} r="2.5" fill="#1dd3b0" fillOpacity="0.7" stroke="none">
                  <title>{`${d.tgl} · Biaya ${rpShort(d.spend)} · Pesanan ${d.orders} · ROI ${d.roi.toFixed(1)}x`}</title>
                </circle>
              ))}
              {daily.map((d, i) => (
                <circle key={`o${i}`} cx={xOf(i)} cy={yOrders(d.orders)} r="2.5" fill="#f5a623" fillOpacity="0.7" stroke="none">
                  <title>{`${d.tgl} · Pesanan ${d.orders} · ROI ${d.roi.toFixed(1)}x`}</title>
                </circle>
              ))}
            </svg>
          </div>
        );
      })()}

      {/* Weekly table */}
      <div style={{ marginTop: 12 }}>
        <SectionLabel>RINGKASAN MINGGUAN</SectionLabel>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead><tr>
              <th style={S.th}>Minggu</th>
              <th style={{ ...S.th, textAlign:"right" }}>Spend</th>
              <th style={{ ...S.th, textAlign:"right" }}>GMV</th>
              <th style={{ ...S.th, textAlign:"right" }}>Orders</th>
              <th style={{ ...S.th, textAlign:"right" }}>ROI</th>
              <th style={{ ...S.th, textAlign:"right" }}>CPA</th>
            </tr></thead>
            <tbody>
              {weeks.map((w, i) => (
                <tr key={i} style={S.tr}>
                  <td style={S.td}><span style={{ fontFamily: mono, fontSize: 12 }}>{w.wk}</span></td>
                  <td style={{ ...S.tdR, color: C.watch }}>{rpShort(w.spend)}</td>
                  <td style={{ ...S.tdR, color: C.accent }}>{rpShort(w.gmv)}</td>
                  <td style={S.tdR}>{nfmt(w.orders)}</td>
                  <td style={{ ...S.tdR, fontWeight: 700, color: w.roi >= 5 ? C.good : w.roi >= 3 ? C.watch : C.bad }}>{w.roi.toFixed(2)}x</td>
                  <td style={{ ...S.tdR, color: C.muted }}>{rpShort(w.spend/Math.max(w.orders,1))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Best/worst days */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 12 }}>
        <div style={{ background: C.panel, border: `1px solid ${C.good}33`, borderRadius: 8, padding: 14 }}>
          <div style={{ fontFamily: mono, fontSize: 10, color: C.good, letterSpacing: 1, marginBottom: 10 }}>🏆 HARI TERBAIK (by ROI)</div>
          {best3.map((d, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: i<2?`1px solid ${C.line}22`:"none" }}>
              <span style={{ fontFamily: mono, fontSize: 12 }}>{d.tgl}</span>
              <span style={{ fontFamily: mono, fontSize: 12, color: C.good, fontWeight: 700 }}>{d.roi.toFixed(2)}x</span>
              <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>{rpShort(d.gmv)}</span>
            </div>
          ))}
        </div>
        <div style={{ background: C.panel, border: `1px solid ${C.bad}33`, borderRadius: 8, padding: 14 }}>
          <div style={{ fontFamily: mono, fontSize: 10, color: C.bad, letterSpacing: 1, marginBottom: 10 }}>📉 HARI TERLEMAH (by ROI)</div>
          {worst3.map((d, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: i<2?`1px solid ${C.line}22`:"none" }}>
              <span style={{ fontFamily: mono, fontSize: 12 }}>{d.tgl}</span>
              <span style={{ fontFamily: mono, fontSize: 12, color: C.bad, fontWeight: 700 }}>{d.roi.toFixed(2)}x</span>
              <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>{rpShort(d.gmv)}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Per-campaign table */}
      {campaignSnap ? (
        <div style={{ marginTop: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <SectionLabel>PER KAMPANYE · {campaignSnap.summary.campaignCount} kampanye · {nfmt(campaignSnap.summary.totalOrders)} orders</SectionLabel>
          </div>
          <div style={S.tableWrap}>
                <table style={S.table}>
                  <thead><tr>
                    <th style={{ ...S.th, width: 18 }}></th>
                    <th style={S.th}>Kampanye</th>
                    <th style={{ ...S.th, textAlign: "right" }}>Spend</th>
                    <th style={{ ...S.th, textAlign: "right" }}>GMV</th>
                    <th style={{ ...S.th, textAlign: "right" }}>ROI</th>
                    <th style={{ ...S.th, textAlign: "right" }}>Orders</th>
                    <th style={{ ...S.th, textAlign: "right" }}>Impresi</th>
                    <th style={{ ...S.th, textAlign: "right" }}>CTR</th>
                    <th style={{ ...S.th, textAlign: "right" }}>CVR</th>
                    <th style={{ ...S.th, textAlign: "right" }}>CPA</th>
                  </tr></thead>
                  <tbody>
                    {campaignSnap.campaigns.map((c, i) => {
                      const roiCol = c.roi >= 7 ? C.good : c.roi >= 4 ? C.accent : c.roi >= 2 ? C.watch : C.bad;
                      const cpa = c.orders > 0 ? c.spend / c.orders : 0;
                      const isOpen = openCamp === i;
                      return (
                        <React.Fragment key={i}>
                          <tr style={{ ...S.tr, cursor: "pointer" }} onClick={() => setOpenCamp(isOpen ? null : i)}>
                            <td style={{ ...S.td, color: C.muted, fontFamily: mono, fontSize: 11 }}>{isOpen ? "▾" : "▸"}</td>
                            <td style={S.td}>
                              <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: C.ink }}>{c.name}</div>
                              <div style={{ fontFamily: sans, fontSize: 11, color: C.dim }}>{c.products.length} creative/produk</div>
                            </td>
                            <td style={{ ...S.tdR, color: C.watch }}>{rpShort(c.spend)}</td>
                            <td style={{ ...S.tdR, color: C.accent }}>{rpShort(c.gmv)}</td>
                            <td style={{ ...S.tdR, fontWeight: 700, color: roiCol }}>{c.roi.toFixed(2)}x</td>
                            <td style={S.tdR}>{nfmt(c.orders)}</td>
                            <td style={{ ...S.tdR, color: C.muted }}>{c.impr > 0 ? nfmt(c.impr) : "—"}</td>
                            <td style={{ ...S.tdR, color: c.ctr > 0 && c.ctr < 1 ? C.watch : C.ink }}>{c.ctr > 0 ? c.ctr.toFixed(2)+"%" : "—"}</td>
                            <td style={{ ...S.tdR, color: c.cvr > 0 && c.cvr < 1 ? C.watch : C.ink }}>{c.cvr > 0 ? c.cvr.toFixed(2)+"%" : "—"}</td>
                            <td style={{ ...S.tdR, color: C.muted }}>{rpShort(cpa)}</td>
                          </tr>
                          {isOpen && (
                            <tr style={S.diagRow}>
                              <td></td>
                              <td colSpan={9} style={{ ...S.diagCell, padding: "10px 12px" }}>
                                <div style={{ fontFamily: mono, fontSize: 9, color: C.muted, letterSpacing: 1.2, marginBottom: 8 }}>DETAIL CREATIVE / PRODUK · TOP 5</div>
                                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 8 }}>
                                  {[...c.products].sort((a,b) => b.gmv - a.gmv).slice(0,5).map((p, j) => {
                                    const pRoiCol = p.roi >= 7 ? C.good : p.roi >= 4 ? C.accent : p.roi >= 2 ? C.watch : C.bad;
                                    return (
                                      <div key={j} style={{ background: C.panel, border: `1px solid ${C.line}`, borderLeft: `3px solid ${pRoiCol}`, borderRadius: "0 6px 6px 0", padding: "8px 10px" }}>
                                        <div style={{ fontFamily: mono, fontSize: 10, color: C.dim, marginBottom: 4 }}>{p.prodId || `#${j+1}`}</div>
                                        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                                          <span style={{ fontFamily: mono, fontSize: 11, color: C.watch }}>Rp{rpShort(p.spend)}</span>
                                          <span style={{ fontFamily: mono, fontSize: 11, color: C.accent }}>{rpShort(p.gmv)}</span>
                                          <span style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: pRoiCol }}>{p.roi.toFixed(1)}x</span>
                                          {p.ctr > 0 && <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>CTR {p.ctr.toFixed(2)}%</span>}
                                          {p.cvr > 0 && <span style={{ fontFamily: mono, fontSize: 11, color: C.muted }}>CVR {p.cvr.toFixed(2)}%</span>}
                                          {p.orders > 0 && <span style={{ fontFamily: mono, fontSize: 11, color: C.good }}>{p.orders} order</span>}
                                        </div>
                                        {p.status && <div style={{ fontFamily: sans, fontSize: 10, color: C.dim, marginTop: 4 }}>{p.status}</div>}
                                      </div>
                                    );
                                  })}
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
        </div>
      ) : (
        <div style={{ marginTop: 16, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 8, padding: "14px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 3 }}>📊 Data Per Kampanye belum diimport</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.muted }}>TikTok Ads Manager → Pelaporan → Produk → Materi Iklan Produk → pilih periode → Ekspor</div>
          </div>
          <button onClick={onImportCampaign} style={{ fontFamily: mono, fontSize: 12, fontWeight: 700, color: "#fff", background: "#fe2c55", border: "none", borderRadius: 7, padding: "9px 16px", cursor: "pointer", flexShrink: 0, marginLeft: 16 }}>
            + Import Campaign
          </button>
        </div>
      )}


      {/* Diagnosa */}
      <div style={{ marginTop: 12 }}>
        <SectionLabel>DIAGNOSA &amp; SARAN</SectionLabel>
        {diagnosa.map((it, i) => (
          <div key={i} style={{ borderLeft: `3px solid ${it.level==="bad"?C.bad:it.level==="good"?C.good:it.level==="watch"?C.watch:C.muted}`, background: (it.level==="bad"?C.bad:it.level==="good"?C.good:it.level==="watch"?C.watch:C.muted)+"0a", borderRadius: "0 6px 6px 0", padding: "10px 14px", marginBottom: 8 }}>
            <div style={{ fontFamily: mono, fontSize: 11, fontWeight: 700, color: it.level==="bad"?C.bad:it.level==="good"?C.good:it.level==="watch"?C.watch:C.muted, marginBottom: 5 }}>{it.tag}</div>
            <div style={{ fontFamily: sans, fontSize: 13, color: C.ink, lineHeight: 1.6 }}>{it.msg}</div>
          </div>
        ))}
      </div>

    </div>
  );
}


/* ============================================================================
   LABA / RUGI TAB  —  P&L Simulator
   Auto-populated: Revenue, Beban Marketplace, Beban Ads, HPP (if COGS filled)
   Manual input: Fixed costs (gaji, sewa, dll), Beban Lain-lain
   ========================================================================== */
function PnLTab({ active, shopeeSnap, tiktokSnap, metaSnap, tiktokAdsSnap, cogsItems, pnlInputs, setPnlInputs, onGoToCogs }) {

  // ── Auto-populated values ──
  const revenue = useMemo(() => {
    const sh = shopeeSnap?.summary?.netGmv || 0;
    const tt = tiktokSnap?.summary?.netGmv || 0;
    return sh + tt;
  }, [shopeeSnap, tiktokSnap]);

  const bebanMarketplace = useMemo(() => {
    const sh = Math.abs(shopeeSnap?.summary?.totalFee || 0);
    const tt = Math.abs(tiktokSnap?.summary?.totalFee || 0);
    return sh + tt;
  }, [shopeeSnap, tiktokSnap]);

  const bebanAds = useMemo(() => {
    const shopee  = active ? active.ads.reduce((t,a) => t+a.spend, 0) : 0;
    const tiktok  = tiktokAdsSnap?.summary?.totalSpend || 0;
    const meta    = metaSnap?.summary?.totalSpend || 0;
    return shopee + tiktok + meta;
  }, [active, tiktokAdsSnap, metaSnap]);

  const hpp = useMemo(() => {
    if (!cogsItems.length || !active) return null;
    // Estimate HPP from ad revenue using blended COGS rate
    // Find avg COGS rate from items that have both hpp and harga
    const withPrice = cogsItems.filter(c => c.harga > 0 && c.total > 0);
    if (!withPrice.length) return null;
    const avgCogsRate = withPrice.reduce((t,c) => t + c.total/c.harga, 0) / withPrice.length;
    return revenue * avgCogsRate;
  }, [cogsItems, revenue]);

  // ── Manual inputs with IDB persistence ──
  async function setInput(key, val) {
    const v = parseFloat(val) || 0;
    const next = { ...pnlInputs, [key]: v };
    setPnlInputs(next);
    await idbPut("pnl_inputs", { key, value: v });
  }

  const g = (key, def = 0) => pnlInputs[key] ?? def;

  // Manual fixed cost items
  const fixedItems = [
    { key: "gaji",       label: "Beban Gaji & SDM",        placeholder: "e.g. 38000000" },
    { key: "sewa",       label: "Beban Sewa",               placeholder: "e.g. 3250000" },
    { key: "ekspedisi",  label: "Beban Ekspedisi & Logistik",placeholder: "e.g. 600000" },
    { key: "listrik",    label: "Beban Listrik & Internet",  placeholder: "e.g. 800000" },
    { key: "software",   label: "Beban Software & Tools",    placeholder: "e.g. 300000" },
    { key: "lainnya",    label: "Beban Operasional Lainnya", placeholder: "e.g. 1000000" },
  ];

  const bebanLainItems = [
    { key: "bunga",      label: "Beban Bunga / Cicilan",    placeholder: "e.g. 1500000" },
    { key: "nonop",      label: "Beban Lain-lain",          placeholder: "e.g. 500000" },
  ];

  const totalFixedManual = fixedItems.reduce((t, it) => t + g(it.key), 0);
  const totalBebanLain   = bebanLainItems.reduce((t, it) => t + g(it.key), 0);
  const totalBebanOps    = bebanMarketplace + bebanAds + totalFixedManual;
  const labaKotor        = revenue - (hpp || 0);
  const labaOperasional  = labaKotor - totalBebanOps;
  const labaBersih       = labaOperasional - totalBebanLain;
  const marginBersih     = revenue > 0 ? labaBersih/revenue*100 : 0;
  const marginKotor      = revenue > 0 ? labaKotor/revenue*100 : 0;
  const hppRate          = revenue > 0 && hpp ? hpp/revenue*100 : null;

  // ── Row helpers ──
  const Row = ({ label, value, bold, color, indent, sub, pct }) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline",
      padding: "9px 16px", borderBottom: `1px solid ${C.line}`,
      background: bold ? C.panel2 : "transparent" }}>
      <div style={{ paddingLeft: indent ? 20 : 0 }}>
        <span style={{ fontFamily: sans, fontSize: bold ? 13.5 : 13, fontWeight: bold ? 800 : 500,
          color: bold ? C.ink : C.ink }}>{label}</span>
        {sub && <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>{sub}</div>}
      </div>
      <div style={{ display: "flex", gap: 16, alignItems: "baseline" }}>
        {pct != null && <span style={{ fontFamily: sans, fontSize: 11.5, color: C.muted, minWidth: 50, textAlign: "right" }}>{pct.toFixed(1)}%</span>}
        <span style={{ fontFamily: sans, fontSize: bold ? 15 : 13.5, fontWeight: bold ? 800 : 600,
          color: color || C.ink, minWidth: 120, textAlign: "right" }}>
          {value >= 0 ? rp(value) : `(${rp(Math.abs(value))})`}
        </span>
      </div>
    </div>
  );

  const InputRow = ({ item }) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center",
      padding: "8px 16px", borderBottom: `1px solid ${C.line}` }}>
      <span style={{ fontFamily: sans, fontSize: 13, color: C.ink, paddingLeft: 20 }}>{item.label}</span>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <span style={{ fontFamily: sans, fontSize: 11.5, color: C.dim }}>Rp</span>
        <input
          type="number"
          defaultValue={g(item.key) || ""}
          placeholder={item.placeholder}
          onBlur={e => setInput(item.key, e.target.value)}
          style={{ width: 140, textAlign: "right", fontFamily: sans, fontSize: 13, fontWeight: 600,
            color: C.ink, background: C.panel2, border: `1px solid ${C.line}`,
            borderRadius: 6, padding: "6px 10px", outline: "none" }}
        />
      </div>
    </div>
  );

  const SectionHead = ({ label }) => (
    <div style={{ padding: "10px 16px 6px", background: C.panel2, borderBottom: `1px solid ${C.line}` }}>
      <span style={{ fontFamily: sans, fontSize: 11, fontWeight: 700, letterSpacing: 0.8,
        color: C.muted, textTransform: "uppercase" }}>{label}</span>
    </div>
  );

  const AutoBadge = () => (
    <span style={{ fontFamily: sans, fontSize: 10, fontWeight: 700, color: C.good,
      background: C.good+"15", border: `1px solid ${C.good}33`,
      padding: "1px 6px", borderRadius: 10, marginLeft: 6 }}>auto</span>
  );

  const ManualBadge = () => (
    <span style={{ fontFamily: sans, fontSize: 10, fontWeight: 700, color: C.watch,
      background: C.watch+"15", border: `1px solid ${C.watch}33`,
      padding: "1px 6px", borderRadius: 10, marginLeft: 6 }}>manual</span>
  );

  const period = shopeeSnap?.periodStart
    ? `${shopeeSnap.periodStart} → ${shopeeSnap.periodEnd}`
    : active ? `${active.periodStart} → ${active.periodEnd}` : "—";

  return (
    <div>
      <SectionLabel>SIMULASI LABA / RUGI · {period}</SectionLabel>
      <DataDisclaimer />
      <div style={S.kpiGrid}>
        <Kpi label="Revenue Bersih" value={rpShort(revenue)} dir="rev" accent={C.accent} big />
        <Kpi label="Laba Kotor" value={rpShort(labaKotor)} dir={labaKotor>=0?"rev":"cost"} accent={labaKotor>=0?C.good:C.bad} />
        <Kpi label="Margin Kotor" value={marginKotor.toFixed(1)+"%"} dir={marginKotor>=30?"rev":"cost"} accent={marginKotor>=40?C.good:marginKotor>=20?C.watch:C.bad} />
        <Kpi label="Laba Operasional" value={rpShort(labaOperasional)} dir={labaOperasional>=0?"rev":"cost"} accent={labaOperasional>=0?C.good:C.bad} />
        <Kpi label="Laba Bersih" value={rpShort(labaBersih)} dir={labaBersih>=0?"rev":"cost"} big accent={labaBersih>=0?C.good:C.bad} />
        <Kpi label="Margin Bersih" value={marginBersih.toFixed(1)+"%"} dir={marginBersih>=5?"rev":"cost"} accent={marginBersih>=10?C.good:marginBersih>=3?C.watch:C.bad} />
      </div>

      {/* Data source info */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12, marginBottom: 4 }}>
        {[
          shopeeSnap && { label: "Shopee income", ok: true },
          tiktokSnap && { label: "TikTok income", ok: true },
          active     && { label: "Shopee Ads", ok: true },
          tiktokAdsSnap && { label: "TikTok Ads", ok: true },
          metaSnap   && { label: "Meta Ads", ok: true },
          cogsItems.length > 0 ? { label: `COGS (${cogsItems.length} SKU)`, ok: true } : { label: "COGS belum diisi", ok: false },
        ].filter(Boolean).map((s, i) => (
          <span key={i} style={{ fontFamily: sans, fontSize: 11, fontWeight: 600,
            color: s.ok ? C.good : C.watch,
            background: (s.ok ? C.good : C.watch)+"12",
            border: `1px solid ${(s.ok ? C.good : C.watch)}33`,
            padding: "2px 8px", borderRadius: 10 }}>
            {s.ok ? "✓" : "⚠"} {s.label}
          </span>
        ))}
      </div>

      {/* P&L Table */}
      <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12,
        overflow: "hidden", marginTop: 12, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>

        {/* PENDAPATAN */}
        <SectionHead label="Pendapatan" />
        <Row label={<>Revenue Bersih (Net GMV)<AutoBadge /></>}
          sub={!shopeeSnap && !tiktokSnap ? "Import file penghasilan untuk data otomatis" : null}
          value={revenue} pct={100} />

        {/* HPP */}
        <SectionHead label="Harga Pokok Penjualan" />
        {hpp != null
          ? <Row label={<>HPP (estimasi dari COGS template)<AutoBadge /></>}
              sub={`Rata-rata COGS rate × revenue · ${hppRate?.toFixed(1)}% dari revenue`}
              value={-hpp} indent pct={hppRate ? -hppRate : null} />
          : <div style={{ padding: "12px 16px", display: "flex", justifyContent: "space-between",
              alignItems: "center", borderBottom: `1px solid ${C.line}` }}>
              <div>
                <span style={{ fontFamily: sans, fontSize: 13, color: C.muted, paddingLeft: 20 }}>HPP</span>
                <span style={{ fontFamily: sans, fontSize: 11, color: C.watch, marginLeft: 8 }}>⚠ COGS belum diisi</span>
              </div>
              <button onClick={onGoToCogs} style={{ fontFamily: sans, fontSize: 12, fontWeight: 700,
                color: C.accent, background: C.accent+"10", border: `1px solid ${C.accent}33`,
                padding: "5px 12px", borderRadius: 6, cursor: "pointer" }}>
                Isi COGS →
              </button>
            </div>
        }

        {/* LABA KOTOR */}
        <Row label="LABA KOTOR" value={labaKotor} bold
          color={labaKotor >= 0 ? C.good : C.bad}
          pct={marginKotor} />

        {/* BEBAN OPERASIONAL */}
        <SectionHead label="Beban Operasional" />
        <Row label={<>Beban Marketplace<AutoBadge /></>}
          sub={[shopeeSnap&&"Shopee", tiktokSnap&&"TikTok"].filter(Boolean).join(" + ") || "Import income file"}
          value={-bebanMarketplace} indent pct={revenue?-bebanMarketplace/revenue*100:null} />
        <Row label={<>Beban Iklan & Ads<AutoBadge /></>}
          sub={[active&&"Shopee", tiktokAdsSnap&&"TikTok GMV Max", metaSnap&&"Meta"].filter(Boolean).join(" + ") || "Import file iklan"}
          value={-bebanAds} indent pct={revenue?-bebanAds/revenue*100:null} />
        {fixedItems.map(item => (
          <InputRow key={item.key} item={item} />
        ))}
        <Row label={<>Total Beban Operasional</>}
          value={-totalBebanOps} bold
          pct={revenue?-totalBebanOps/revenue*100:null} />

        {/* LABA OPERASIONAL */}
        <Row label="LABA OPERASIONAL" value={labaOperasional} bold
          color={labaOperasional >= 0 ? C.good : C.bad}
          pct={revenue?labaOperasional/revenue*100:null} />

        {/* BEBAN LAIN-LAIN */}
        <SectionHead label="Beban Lain-lain" />
        {bebanLainItems.map(item => (
          <InputRow key={item.key} item={item} />
        ))}
        {totalBebanLain > 0 && (
          <Row label="Total Beban Lain-lain" value={-totalBebanLain}
            pct={revenue?-totalBebanLain/revenue*100:null} />
        )}

        {/* LABA BERSIH */}
        <Row label="LABA BERSIH" value={labaBersih} bold
          color={labaBersih >= 0 ? C.good : C.bad}
          pct={marginBersih} />
      </div>

      {/* Insight */}
      <div style={{ marginTop: 16 }}>
        <SectionLabel>INSIGHT</SectionLabel>
        {[
          revenue === 0 && { level: "watch", msg: "Import file penghasilan Shopee/TikTok untuk populate revenue otomatis." },
          !cogsItems.length && { level: "watch", msg: "HPP belum bisa dihitung — isi COGS template dulu di tab COGS/Margin." },
          hppRate && hppRate > 50 && { level: "bad", msg: `HPP ${hppRate.toFixed(0)}% dari revenue — sangat tinggi. Untuk fashion, idealnya di bawah 40%. Cek pricing atau efisiensi produksi.` },
          marginKotor > 0 && marginKotor < 30 && { level: "watch", msg: `Margin kotor ${marginKotor.toFixed(1)}% — tipis untuk fashion. Minimum sehat adalah 40%+ untuk bisa nutup beban operasional dan tetap profit.` },
          labaBersih < 0 && revenue > 0 && { level: "bad", msg: `Posisi rugi Rp ${rpShort(Math.abs(labaBersih))} bulan ini. Dua lever utama: naikkan revenue atau potong beban operasional (terutama yang fixed).` },
          labaBersih > 0 && marginBersih < 5 && { level: "watch", msg: `Laba bersih tipis ${marginBersih.toFixed(1)}%. Bisnis profitable tapi tidak punya buffer. Fokus pada scale channel dengan ROAS tertinggi.` },
          marginBersih >= 10 && { level: "good", msg: `Margin bersih ${marginBersih.toFixed(1)}% — sehat untuk skala ini. Reinvestasi ke channel terbaik atau bangun cash buffer minimum 2 bulan fixed cost.` },
          bebanAds > 0 && revenue > 0 && bebanAds/revenue > 0.25 && { level: "watch", msg: `Beban ads ${(bebanAds/revenue*100).toFixed(0)}% dari revenue — terlalu besar. Benchmark sehat: 10–18% untuk fashion marketplace. Evaluasi channel dengan ROAS terendah.` },
        ].filter(Boolean).map((it, i) => it && (
          <div key={i} style={{ borderLeft: `3px solid ${it.level==="bad"?C.bad:it.level==="good"?C.good:C.watch}`,
            background: (it.level==="bad"?C.bad:it.level==="good"?C.good:C.watch)+"0a",
            borderRadius: "0 8px 8px 0", padding: "10px 14px", marginBottom: 8,
            fontFamily: sans, fontSize: 13, color: C.ink, lineHeight: 1.6 }}>
            {it.msg}
          </div>
        ))}
      </div>
    </div>
  );
}

function QuadTag({ q }) {
  const ql = QUAD[q] || QUAD.dog;
  return (
    <span style={{ ...S.skuQuad, color: ql.color, borderColor: ql.color + "55", background: ql.color + "15" }}>
      {ql.label}
    </span>
  );
}

function OvInput({ val, ph, onSet }) {
  return <input type="number" defaultValue={val ?? ""} placeholder={ph}
    onBlur={(e) => onSet(e.target.value)} style={{ ...S.marginInput, width: 58 }} />;
}
function FcDisclaimerCashcow({ rows }) {
  const cc = rows.filter((r) => r.quadrant === "cashcow");
  if (!cc.length) return null;
  return <div style={S.fcAllocNote}>⚠ {cc.length} Cash Cow di rencana ini punya CVR rendah  -  target revenue-nya butuh perbaikan konversi dulu, kalau cuma tambah spend bakal boros.</div>;
}

// SVG scatter: x=share (log-ish via sqrt), y=cvr
function BCGScatter({ classified, onPick, openCode }) {
  const W = 680, H = 380, pad = 44;
  const maxShare = Math.max(...classified.map((p) => p.share), 1);
  const maxCvr = Math.max(...classified.map((p) => p.blendedCvr || p.cvr), 1) * 1.1;
  const shareCut = (() => {
    const s = classified.map((p) => p.share).sort((a, b) => a - b);
    const meanShare = 100 / Math.max(classified.length, 1);
    return Math.max(s[Math.floor(s.length / 2)] || 0, meanShare * 1.5, 2);
  })();
  const cvrCut = (() => {
    const c = classified.map((p) => p.blendedCvr || p.cvr).filter((x) => x > 0).sort((a, b) => a - b);
    return Math.max(c[Math.floor(c.length / 2)] || 0, 0.8);
  })();
  const sx = (share) => pad + (Math.sqrt(share) / Math.sqrt(maxShare)) * (W - pad * 2);
  const sy = (cvr) => H - pad - (cvr / maxCvr) * (H - pad * 2);
  const cutX = sx(shareCut), cutY = sy(cvrCut);

  // Tooltip state
  const [hovered, setHovered] = React.useState(null);

  return (
    <div style={S.scatterWrap}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto" }}>
        {/* quadrant backgrounds */}
        <rect x={cutX} y={pad} width={W - pad - cutX} height={cutY - pad} fill={QUAD.star.color} opacity="0.06" />
        <rect x={cutX} y={cutY} width={W - pad - cutX} height={H - pad - cutY} fill={QUAD.cashcow.color} opacity="0.06" />
        <rect x={pad} y={pad} width={cutX - pad} height={cutY - pad} fill={QUAD.question.color} opacity="0.06" />
        <rect x={pad} y={cutY} width={cutX - pad} height={H - pad - cutY} fill={QUAD.dog.color} opacity="0.06" />
        {/* median lines */}
        <line x1={cutX} y1={pad} x2={cutX} y2={H - pad} stroke={C.line} strokeDasharray="4 4" />
        <line x1={pad} y1={cutY} x2={W - pad} y2={cutY} stroke={C.line} strokeDasharray="4 4" />
        {/* axes */}
        <line x1={pad} y1={H - pad} x2={W - pad} y2={H - pad} stroke={C.line} />
        <line x1={pad} y1={pad} x2={pad} y2={H - pad} stroke={C.line} />
        {/* quadrant labels */}
        <text x={W - pad - 6} y={pad + 14} textAnchor="end" fill={QUAD.star.color} fontSize="11" fontFamily={mono}>STAR</text>
        <text x={W - pad - 6} y={H - pad - 6} textAnchor="end" fill={QUAD.cashcow.color} fontSize="11" fontFamily={mono}>CASH COW</text>
        <text x={pad + 6} y={pad + 14} fill={QUAD.question.color} fontSize="11" fontFamily={mono}>QUESTION</text>
        <text x={pad + 6} y={H - pad - 6} fill={QUAD.dog.color} fontSize="11" fontFamily={mono}>DOG</text>
        {/* axis titles */}
        <text x={W / 2} y={H - 10} textAnchor="middle" fill={C.muted} fontSize="11" fontFamily={mono}>revenue share →</text>
        <text x={14} y={H / 2} textAnchor="middle" fill={C.muted} fontSize="11" fontFamily={mono} transform={`rotate(-90 14 ${H / 2})`}>CVR produk →</text>
        {/* points — use blendedCvr for position */}
        {classified.map((p) => {
          const cvr = p.blendedCvr || p.cvr;
          const r = Math.max(4, Math.min(16, Math.sqrt(p.netSales || p.sales) / 700));
          const q = QUAD[p.quadrant];
          const on = openCode === p.code;
          const hov = hovered === p.code;
          const x = sx(p.share), y = sy(cvr);
          return (
            <g key={p.code}
              style={{ cursor: "pointer" }}
              onClick={() => {
                onPick(p.code);
                // Scroll to SKU card
                setTimeout(() => {
                  const el = document.getElementById("sku-" + p.code);
                  if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                }, 60);
              }}
              onMouseEnter={() => setHovered(p.code)}
              onMouseLeave={() => setHovered(null)}>
              <circle cx={x} cy={y} r={on ? r + 4 : hov ? r + 2 : r}
                fill={q.color} opacity={on ? 1 : hov ? 0.85 : 0.6}
                stroke={on ? "#fff" : hov ? q.color : q.color}
                strokeWidth={on ? 2.5 : hov ? 1.5 : 0.5} />
              {/* Name label on hover or selected */}
              {(hov || on) && (() => {
                const name = p.name.length > 22 ? p.name.slice(0,22)+"…" : p.name;
                const lx = x + r + 6;
                const anchor = lx + name.length * 5.5 > W - 10 ? "end" : "start";
                const lxAdj = anchor === "end" ? x - r - 6 : lx;
                return (
                  <g>
                    <rect x={lxAdj - 3} y={y - 9} width={name.length * 5.8 + 8} height={14}
                      fill={C.panel} opacity="0.88" rx="3" />
                    <text x={lxAdj} y={y + 2} textAnchor={anchor === "end" ? "end" : "start"}
                      fontSize="9.5" fontFamily={mono} fill={q.color} fontWeight="700">
                      {name}
                    </text>
                  </g>
                );
              })()}
            </g>
          );
        })}
      </svg>
    </div>
  );
}

/* ---------- Small components ---------------------------------------------- */
function SectionLabel({ children, inline }) {
  return <div style={{ ...S.sectionLabel, ...(inline ? { margin: 0 } : {}) }}>{children}</div>;
}
function Kpi({ label, value, sub, dir, big, accent }) {
  let subColor = C.muted;
  if (sub && typeof sub === "object") {
    const favorable = dir === "cost" ? !sub.up : sub.up;
    subColor = dir === "neutral" ? C.muted : favorable ? C.good : C.bad;
  }
  const ac = accent || C.line;
  const valStr = String(value || "");
  // Adaptive font size: shorter values get bigger font
  const fontSize = valStr.length <= 6 ? 22 : valStr.length <= 9 ? 18 : valStr.length <= 12 ? 15 : 13;
  return (
    <div style={{ ...S.kpi, ...(big ? S.kpiBig : {}), borderTop: `3px solid ${ac}`, borderColor: C.line, borderTopColor: ac }}>
      <div style={S.kpiLabel}>{label}</div>
      <div style={{ ...S.kpiVal, fontSize, color: accent || C.ink, wordBreak: "break-all", lineHeight: 1.15 }}>{value}</div>
      {sub && <div style={{ ...S.kpiSub, color: subColor }}>{sub.text} vs prev</div>}
    </div>
  );
}

/* ---------- Theme --------------------------------------------------------- */
const C = {
  bg:     "#f5f6fa",   // page background — very light grey
  panel:  "#ffffff",   // card surface
  panel2: "#f0f1f5",   // secondary surface / table header
  line:   "#e2e4ea",   // borders
  ink:    "#1a1d23",   // primary text
  muted:  "#6b7280",   // secondary text
  dim:    "#9ca3af",   // tertiary / placeholder
  good:   "#16a34a",   // green
  bad:    "#dc2626",   // red
  watch:  "#d97706",   // amber
  accent: "#2563eb",   // blue — primary action
};
const mono = "'Plus Jakarta Sans','SF Mono','JetBrains Mono',ui-monospace,monospace";
const sans = "'Plus Jakarta Sans',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif";

const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
  * { box-sizing: border-box; }
  body { margin: 0; background: ${C.bg}; }
  ::-webkit-scrollbar { height: 6px; width: 6px; }
  ::-webkit-scrollbar-track { background: ${C.panel2}; }
  ::-webkit-scrollbar-thumb { background: ${C.line}; border-radius: 6px; }
  input::placeholder { color: ${C.dim}; }
  tbody tr:hover { background: ${C.panel2}; }
  button:hover { opacity: 0.88; }
  @media (prefers-reduced-motion: no-preference) {
    .dd-fade { animation: ddf .2s ease; }
  }
  @keyframes ddf { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; } }
`;
const S = {
  root: { minHeight: "100vh", background: C.bg, color: C.ink, fontFamily: sans, fontSize: 14 },

  // ── Topbar ──
  topbar: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 28px", borderBottom: `1px solid ${C.line}`, position: "sticky", top: 0, background: C.panel, zIndex: 10, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" },
  brand: { display: "flex", gap: 12, alignItems: "center" },
  brandMark: { fontSize: 24, color: C.accent, lineHeight: 1 },
  brandName: { fontFamily: sans, fontWeight: 800, letterSpacing: 0.5, fontSize: 15, color: C.ink },
  brandSub: { fontSize: 11, color: C.muted, letterSpacing: 0.2 },
  topActions: { display: "flex", gap: 8, alignItems: "center" },
  storeTag: { fontFamily: sans, fontSize: 11.5, fontWeight: 600, color: C.muted, border: `1px solid ${C.line}`, background: C.panel2, padding: "5px 10px", borderRadius: 6 },
  importBtn: { background: C.accent, color: "#fff", border: "none", padding: "9px 16px", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: sans, boxShadow: "0 1px 3px rgba(37,99,235,0.3)" },
  importBtn2: { background: "transparent", color: C.accent, border: `1px solid ${C.accent}`, padding: "9px 16px", borderRadius: 8, fontWeight: 600, fontSize: 13, cursor: "pointer", fontFamily: sans },
  importBtnBig: { background: C.accent, color: "#fff", border: "none", padding: "12px 24px", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer", marginTop: 18, boxShadow: "0 2px 6px rgba(37,99,235,0.25)" },
  importBtn2Big: { background: "transparent", color: C.accent, border: `1px solid ${C.accent}`, padding: "12px 24px", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer", marginTop: 18 },
  resetBtn: { background: "transparent", color: C.muted, border: `1px solid ${C.line}`, padding: "8px 12px", borderRadius: 8, fontWeight: 600, fontSize: 12, cursor: "pointer", fontFamily: "inherit" },
  importMenu: { position: "absolute", top: "calc(100% + 8px)", right: 0, background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: "8px", minWidth: 270, zIndex: 50, boxShadow: "0 8px 24px rgba(0,0,0,0.12), 0 2px 8px rgba(0,0,0,0.06)" },
  importMenuItem: { display: "flex", alignItems: "center", gap: 12, width: "100%", background: "transparent", border: "none", borderRadius: 8, padding: "10px 10px", cursor: "pointer", textAlign: "left" },

  // ── Source strip ──
  srcStrip: { display: "flex", gap: 12, alignItems: "center", padding: "8px 28px", borderBottom: `1px solid ${C.line}`, background: C.panel2, flexWrap: "wrap" },
  srcChip: { display: "inline-flex", alignItems: "center", gap: 6, fontFamily: sans, fontSize: 11.5, color: C.muted, fontWeight: 500 },
  srcDot: { width: 7, height: 7, borderRadius: 7 },
  srcFull: { fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.good, border: `1px solid ${C.good}44`, background: C.good + "0f", padding: "3px 9px", borderRadius: 20 },

  // ── Empty states ──
  emptyCards: { display: "flex", gap: 14, justifyContent: "center", flexWrap: "wrap", marginTop: 24 },
  emptyCard: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: 20, width: 260, textAlign: "left", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  emptyCardTitle: { fontFamily: sans, fontSize: 13.5, fontWeight: 700, marginBottom: 8, color: C.ink },
  emptyCardDesc: { fontSize: 12.5, color: C.muted, lineHeight: 1.55 },

  // ── BCG / Product ──
  bcgIntro: { fontSize: 13, color: C.muted, lineHeight: 1.55, marginBottom: 16, maxWidth: 760 },
  invIntro: { fontSize: 13, color: C.muted, lineHeight: 1.55, marginBottom: 16, maxWidth: 800, borderLeft: `3px solid ${C.accent}33`, paddingLeft: 14, background: C.accent + "05", borderRadius: "0 6px 6px 0", padding: "10px 14px" },
  scatterWrap: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" },
  quadLegend: { display: "flex", gap: 18, flexWrap: "wrap", marginTop: 12 },
  quadLegendItem: { display: "inline-flex", alignItems: "center", gap: 7, fontSize: 12.5 },
  quadDot: { width: 10, height: 10, borderRadius: 10 },

  // ── SKU cards ──
  skuCard: { background: C.panel, border: `1px solid ${C.line}`, borderLeft: "3px solid", borderRadius: 10, marginBottom: 8, boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  skuHead: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "13px 16px", cursor: "pointer", gap: 12, flexWrap: "wrap" },
  skuHeadL: { display: "flex", alignItems: "center", gap: 10, minWidth: 0 },
  skuQuad: { fontFamily: sans, fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, padding: "3px 9px", borderRadius: 20, border: "1px solid", whiteSpace: "nowrap" },
  skuName: { fontWeight: 700, fontSize: 13.5, color: C.ink },
  skuHeadR: { display: "flex", alignItems: "center", gap: 14, fontFamily: sans, fontSize: 12, color: C.muted, whiteSpace: "nowrap" },
  skuStat: { fontFamily: sans, fontSize: 12, color: C.muted },
  skuBody: { padding: "4px 16px 14px", borderTop: `1px solid ${C.line}` },
  skuDiag: { borderLeft: "3px solid", paddingLeft: 12, padding: "8px 12px", marginTop: 9, borderRadius: "0 6px 6px 0" },
  skuDiagTitle: { fontFamily: sans, fontSize: 11.5, fontWeight: 700, marginBottom: 4 },
  skuDiagMsg: { fontSize: 13, color: C.ink, lineHeight: 1.6 },

  // ── Stock ──
  stockRow: { display: "flex", alignItems: "center", gap: 12, padding: "10px 0 6px", flexWrap: "wrap" },
  stockLabel: { fontFamily: sans, fontSize: 11, letterSpacing: 0.5, color: C.muted, fontWeight: 600, textTransform: "uppercase", whiteSpace: "nowrap" },
  stockInput: { width: 90, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 8, color: C.ink, padding: "7px 10px", fontFamily: sans, fontSize: 13, outline: "none" },
  stockDio: { fontFamily: sans, fontSize: 12, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" },

  // ── Forecast ──
  fcModeRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-end", gap: 16, flexWrap: "wrap", marginTop: 8 },
  fcToggle: { display: "inline-flex", background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 10, padding: 3 },
  fcToggleBtn: { background: "none", border: "none", color: C.muted, padding: "8px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", borderRadius: 8, fontFamily: sans },
  fcToggleOn: { background: C.panel, color: C.accent, boxShadow: "0 1px 3px rgba(0,0,0,0.1)" },
  fcInputWrap: { display: "flex", flexDirection: "column", gap: 5, minWidth: 220 },
  fcInputLabel: { fontFamily: sans, fontSize: 11, letterSpacing: 0.5, color: C.muted, fontWeight: 600, textTransform: "uppercase" },
  fcDisclaimer: { fontSize: 12.5, color: C.muted, lineHeight: 1.55, margin: "16px 0", maxWidth: 800, borderLeft: `3px solid ${C.line}`, paddingLeft: 12 },
  fcKpiRow: { display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(150px,1fr))", gap: 10, marginBottom: 8 },
  fcAllocNote: { fontFamily: sans, fontSize: 12, color: C.watch, fontWeight: 600, marginBottom: 10 },
  fcFootnote: { fontSize: 11.5, color: C.muted, lineHeight: 1.5, marginTop: 10, fontStyle: "italic" },

  // ── Snapshot strip ──
  snapStrip: { display: "flex", gap: 8, alignItems: "center", padding: "10px 28px", borderBottom: `1px solid ${C.line}`, background: C.panel2, overflowX: "auto" },
  snapStripLabel: { fontFamily: sans, fontSize: 10.5, fontWeight: 600, color: C.dim, letterSpacing: 1, marginRight: 4, textTransform: "uppercase" },
  snapChip: { fontFamily: sans, fontSize: 12, fontWeight: 500, background: C.panel, color: C.muted, border: `1px solid ${C.line}`, padding: "5px 12px", borderRadius: 20, cursor: "pointer", whiteSpace: "nowrap", boxShadow: "0 1px 2px rgba(0,0,0,0.04)" },
  snapChipActive: { background: C.accent, color: "#fff", borderColor: C.accent, boxShadow: "0 2px 6px rgba(37,99,235,0.25)" },

  // ── Tabs ──
  tabs: { display: "flex", gap: 2, padding: "0 28px", borderBottom: `1px solid ${C.line}`, background: C.panel, overflowX: "auto" },
  tab: { fontFamily: sans, fontWeight: 500, fontSize: 13.5, color: C.muted, background: "none", border: "none", borderBottom: "2px solid transparent", padding: "13px 14px", cursor: "pointer", whiteSpace: "nowrap", marginBottom: -1 },
  tabActive: { color: C.accent, borderBottomColor: C.accent, fontWeight: 700 },
  main: { padding: "28px 28px 80px", maxWidth: 1200, margin: "0 auto" },

  // ── Section labels ──
  sectionLabel: { fontFamily: sans, fontSize: 11, fontWeight: 700, letterSpacing: 1, color: C.muted, textTransform: "uppercase", margin: "30px 0 14px", paddingBottom: 8, borderBottom: `1px solid ${C.line}` },

  // ── KPI cards ──
  kpiGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(170px,1fr))", gap: 12 },
  kpi: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: "16px 18px", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  kpiBig: { borderColor: C.accent + "33", background: C.accent + "04" },
  kpiLabel: { fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.muted, letterSpacing: 0.5, marginBottom: 8, textTransform: "uppercase" },
  kpiVal: { fontFamily: sans, fontSize: 18, fontWeight: 800, color: C.ink },
  kpiSub: { fontFamily: sans, fontSize: 11, marginTop: 5, color: C.muted },

  // ── CM cards ──
  cmRow: { display: "flex", gap: 12, flexWrap: "wrap", alignItems: "stretch" },
  cmCard: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: "16px 18px", minWidth: 180, flex: 1, boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  cmCardFinal: { background: C.accent + "06", borderColor: C.accent + "33" },
  cmLabel: { fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.muted, letterSpacing: 0.3, marginBottom: 8 },
  cmVal: { fontFamily: sans, fontSize: 20, fontWeight: 800, color: C.ink },
  cmEmpty: { background: C.panel2, border: `1px dashed ${C.line}`, borderRadius: 10, padding: "16px 18px", color: C.muted, fontSize: 13, width: "100%" },
  cmNote: { fontFamily: sans, fontSize: 11.5, fontWeight: 600, color: C.watch, width: "100%", marginTop: 4 },

  // ── Action plan ──
  planGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 12 },
  planCol: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  planHead: { fontFamily: sans, fontSize: 12, fontWeight: 800, letterSpacing: 0.5, marginBottom: 12 },
  planEmpty: { color: C.dim, fontFamily: sans, fontSize: 12 },
  planItem: { padding: "9px 0", borderTop: `1px solid ${C.line}` },
  planAd: { fontSize: 13, fontWeight: 600, marginBottom: 3, color: C.ink },
  planMsg: { fontSize: 12, color: C.muted, lineHeight: 1.5 },

  // ── Performance table ──
  perfBar: { display: "flex", justifyContent: "space-between", alignItems: "center", margin: "28px 0 14px", paddingBottom: 8, borderBottom: `1px solid ${C.line}` },
  search: { background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 8, color: C.ink, padding: "8px 13px", fontSize: 13, fontFamily: sans, outline: "none", minWidth: 200 },

  tableWrap: { overflowX: "auto", border: `1px solid ${C.line}`, borderRadius: 12, boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { textAlign: "left", fontFamily: sans, fontSize: 11, fontWeight: 700, letterSpacing: 0.5, color: C.muted, textTransform: "uppercase", padding: "12px 14px", background: C.panel2, borderBottom: `1px solid ${C.line}`, whiteSpace: "nowrap", position: "sticky", top: 0 },
  tr: { borderBottom: `1px solid ${C.line}` },
  td: { padding: "11px 14px", verticalAlign: "top", color: C.ink },
  tdR: { padding: "11px 14px", textAlign: "right", fontFamily: sans, fontWeight: 600, whiteSpace: "nowrap", color: C.ink },
  adName: { fontWeight: 700, fontSize: 13, color: C.ink },
  adMeta: { fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 3 },
  biddingTag: { fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.muted, border: `1px solid ${C.line}`, background: C.panel2, padding: "2px 8px", borderRadius: 20, whiteSpace: "nowrap" },
  pill: { fontFamily: sans, fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, padding: "3px 9px", borderRadius: 20, border: "1px solid", whiteSpace: "nowrap" },
  muted: { color: C.muted, fontFamily: sans, fontSize: 11.5 },

  // ── Threshold panel ──
  thPanel: { background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 12, padding: "14px 18px", marginTop: 8 },
  thHead: { display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 14 },
  thHeadLabel: { fontFamily: sans, fontSize: 11.5, fontWeight: 700, letterSpacing: 0.5, color: C.muted, textTransform: "uppercase" },
  thToggle: { fontFamily: sans, fontSize: 12, fontWeight: 500, color: C.muted, display: "flex", alignItems: "center", gap: 6, cursor: "pointer" },
  thFields: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 10 },
  thField: { display: "flex", flexDirection: "column", gap: 5 },
  thLabel: { fontFamily: sans, fontSize: 11, fontWeight: 600, letterSpacing: 0.3, color: C.muted, textTransform: "uppercase" },
  thInputWrap: { display: "flex", alignItems: "center", background: C.panel, border: `1px solid ${C.line}`, borderRadius: 8, padding: "0 10px" },
  thInput: { background: "none", border: "none", color: C.ink, fontFamily: sans, fontSize: 14, fontWeight: 600, padding: "8px 0", width: "100%", outline: "none" },
  thSuffix: { fontFamily: sans, fontSize: 12, color: C.muted, marginLeft: 4 },

  // ── Diagnosa ──
  diagRow: { background: C.panel2 },
  diagCell: { padding: "14px 18px 18px" },
  diagLabel: { fontFamily: sans, fontSize: 11, fontWeight: 700, letterSpacing: 0.5, color: C.dim, marginBottom: 12, textTransform: "uppercase" },
  diagItem: { borderLeft: "3px solid", paddingLeft: 12, marginBottom: 10, display: "flex", flexDirection: "column", gap: 3 },
  diagTag: { fontFamily: sans, fontSize: 11, fontWeight: 700, letterSpacing: 0.3 },
  diagMsg: { fontSize: 13, color: C.ink, lineHeight: 1.6 },

  // ── Strategy ──
  stratIntro: { fontSize: 13, color: C.muted, lineHeight: 1.55, marginBottom: 16, maxWidth: 760 },
  stratGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 12 },
  stratCard: { background: C.panel, border: `1px solid ${C.line}`, borderTop: "3px solid", borderRadius: 12, padding: 18, boxShadow: "0 1px 3px rgba(0,0,0,0.05)" },
  stratTitle: { fontFamily: sans, fontSize: 13, fontWeight: 800, marginBottom: 10, color: C.ink },
  stratWhen: { fontSize: 13, color: C.ink, marginBottom: 8, lineHeight: 1.55 },
  stratAction: { fontSize: 13, color: C.muted, lineHeight: 1.55, marginBottom: 12 },
  stratAds: { borderTop: `1px solid ${C.line}`, paddingTop: 10, display: "flex", flexDirection: "column", gap: 6 },
  stratAdRow: { display: "flex", justifyContent: "space-between", fontSize: 12.5, gap: 8 },

  // ── COGS ──
  cogsHelp: { fontSize: 13, color: C.muted, lineHeight: 1.55, marginBottom: 14, maxWidth: 760 },
  bulkRow: { display: "flex", gap: 8, marginBottom: 14 },
  smallBtn: { background: C.panel2, color: C.ink, border: `1px solid ${C.line}`, borderRadius: 8, padding: "8px 14px", fontSize: 13, cursor: "pointer", fontFamily: sans, fontWeight: 600 },
  marginInput: { width: 68, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 6, color: C.ink, padding: "6px 8px", fontFamily: sans, fontSize: 13, fontWeight: 600, textAlign: "right", outline: "none" },

  // ── Empty ──
  empty: { textAlign: "center", padding: "80px 20px", maxWidth: 520, margin: "0 auto" },
  emptyMark: { fontSize: 48, opacity: 0.5 },
  emptyTitle: { fontFamily: sans, fontWeight: 800, letterSpacing: 0.5, marginTop: 14, fontSize: 20, color: C.ink },
  emptyText: { color: C.muted, lineHeight: 1.65, fontSize: 14, marginTop: 10 },

  // ── Dialog ──
  dlgOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, backdropFilter: "blur(2px)" },
  dlgBox: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 14, padding: "24px 26px", maxWidth: 400, width: "90%", boxShadow: "0 20px 60px rgba(0,0,0,0.15)" },
  dlgMsg: { fontSize: 14, color: C.ink, lineHeight: 1.6, marginBottom: 20 },
  dlgBtns: { display: "flex", gap: 10, justifyContent: "flex-end" },
  dlgCancel: { background: C.panel2, color: C.muted, border: `1px solid ${C.line}`, padding: "9px 18px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  dlgOk: { background: C.bad, color: "#fff", border: "none", padding: "9px 18px", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  toast: { position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: C.ink, color: "#fff", padding: "12px 20px", borderRadius: 10, fontFamily: sans, fontSize: 13, fontWeight: 500, zIndex: 50, boxShadow: "0 4px 20px rgba(0,0,0,0.2)", whiteSpace: "nowrap" },
};
export { DEFAULT_SEASONAL };
// end
