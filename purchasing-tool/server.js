const express = require("express");
const multer  = require("multer");
const xlsx    = require("xlsx");
const path    = require("path");
const fs      = require("fs");
const os      = require("os");

const app  = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, "public")));
app.use(express.json({ limit: "20mb" }));

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = path.join(__dirname, "data", "scans");
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => { cb(null, `scan_${Date.now()}.png`); },
});
const upload = multer({ storage });

const EXCEL_PATH = path.join(__dirname, "data", "monitoring.xlsx");

// ── Months we look for as monitoring sheets ───────────────────────────────────
// Handles: "MAY2026", "May 2026", "JANUARY 2026", "jan", etc.
const MONTH_PATTERNS = [
  "JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE",
  "JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER",
  "JAN","FEB","MAR","APR","JUN",
  "JUL","AUG","SEP","OCT","NOV","DEC",
];

function isMonthSheet(name) {
  const u = name.toUpperCase().replace(/\s/g, "");
  return MONTH_PATTERNS.some(m => u.includes(m));
}

// ── Read workbook ─────────────────────────────────────────────────────────────
function getWorkbook() {
  if (!fs.existsSync(EXCEL_PATH)) createBlankWorkbook();
  return xlsx.readFile(EXCEL_PATH);
}

function createBlankWorkbook() {
  const wb = xlsx.utils.book_new();
  const headers = [
    "PR DATE","PR DATE RECEIVED","PR NO.","REQUESTING DEPT.",
    "PO DATE","PO NUMBER","END USER/S","SUPPLIER'S NAME",
    "ITEM CODE","ITEM DESCRIPTION","SPECIFICATIONS","QTY","UoM",
    "UNIT PRICE","AMOUNT","TOTAL AMOUNT","PAYMENT TERMS","PR REQUIRED DATE",
    "DATE DELIVERED","REMARKS","PURCHASE ORDER STATUS",
    "ITEMS/SERVICES","ORIGINAL PRICE","TOTAL COST SAVINGS"
  ];
  ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"].forEach(m => {
    xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([headers]), m);
  });
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ["Month","No. of PRs","Processed PO (Served)","Processed PR (For PO)",
     "Processed PO (Waiting for Delivery)","Cancelled PO","Total PO Created","Remarks"]
  ]), "Summary");
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ["SUPPLIER'S NAME","TOTAL AMOUNT","ORIGINAL PRICE","TOTAL COST SAVINGS"]
  ]), "Total cost savings");
  xlsx.writeFile(wb, EXCEL_PATH);
}

// ── Parse a monitoring sheet into clean data rows ─────────────────────────────
// The actual format has:
//   rows 1-4: title/summary block (skip)
//   row 5: column headers
//   row 6+: data (with merged cells spanning multiple item rows)
function parseMonthSheet(ws) {
  const ref = ws["!ref"];
  if (!ref) return [];

  const raw = xlsx.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: false,          // format dates as strings
    dateNF: "yyyy-mm-dd",
  });

  if (raw.length < 5) return [];

  // Find the header row — look for "PR DATE" or "PR NO." in any row
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(raw.length, 10); i++) {
    const row = raw[i];
    if (row && row.some(c => c && String(c).toUpperCase().includes("PR DATE"))) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) return [];

  const headers = raw[headerRowIdx].map(h => h ? String(h).trim().toUpperCase() : "");

  // Column index lookup
  const col = name => headers.findIndex(h => h === name.toUpperCase());

  const rows = [];
  for (let i = headerRowIdx + 1; i < raw.length; i++) {
    const r = raw[i];
    if (!r || r.every(v => v === null || v === "")) continue;

    // Skip rows that are just formula errors
    if (r.every(v => v === null || String(v).startsWith("#"))) continue;

    const get = (name) => {
      const idx = col(name);
      if (idx === -1) return "";
      const val = r[idx];
      if (val === null || val === undefined) return "";
      if (String(val).startsWith("#")) return ""; // formula error
      return String(val).trim();
    };

    // Must have at least one of these to be a real data row
    const hasContent = get("PO NUMBER") || get("PR NO.") || get("ITEM DESCRIPTION") || get("SUPPLIER'S NAME");
    if (!hasContent) continue;

    const parseNum = v => {
      const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
      return isNaN(n) ? 0 : n;
    };

    rows.push({
      "PR DATE":               get("PR DATE"),
      "PR DATE RECEIVED":      get("PR DATE RECEIVED"),
      "PR NO.":                get("PR NO."),
      "REQUESTING DEPT.":      get("REQUESTING DEPT."),
      "PO DATE":               get("PO DATE"),
      "PO NUMBER":             get("PO NUMBER"),
      "END USER/S":            get("END USER/S"),
      "SUPPLIER'S NAME":       get("SUPPLIER'S NAME"),
      "ITEM CODE":             get("ITEM CODE"),
      "ITEM DESCRIPTION":      get("ITEM DESCRIPTION"),
      "SPECIFICATIONS":        get("SPECIFICATIONS"),
      "QTY":                   get("QTY"),
      "UoM":                   get("UOM") || get("UoM"),
      "UNIT PRICE":            parseNum(get("UNIT PRICE")),
      "AMOUNT":                parseNum(get("AMOUNT")),
      "TOTAL AMOUNT":          parseNum(get("TOTAL AMOUNT")),
      "PAYMENT TERMS":         get("PAYMENT TERMS"),
      "PR REQUIRED DATE":      get("PR REQUIRED DATE"),
      "DATE DELIVERED":        get("DATE DELIVERED"),
      "REMARKS":               get("REMARKS"),
      "PURCHASE ORDER STATUS": get("PURCHASE ORDER STATUS"),
      "ITEMS/SERVICES":        get("ITEMS/SERVICES"),
      "ORIGINAL PRICE":        parseNum(get("ORIGINAL PRICE")),
      "TOTAL COST SAVINGS":    parseNum(get("TOTAL COST SAVINGS")),
    });
  }
  return rows;
}

// ── Get all month sheets and their parsed rows ─────────────────────────────────
function getAllMonthData() {
  const wb = getWorkbook();
  const result = {};
  wb.SheetNames.forEach(name => {
    if (isMonthSheet(name)) {
      result[name] = parseMonthSheet(wb.Sheets[name]);
    }
  });
  return result;
}

// ── API: sheet names ──────────────────────────────────────────────────────────
app.get("/api/sheets", (req, res) => {
  try {
    const wb = getWorkbook();
    res.json({ sheets: wb.SheetNames });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── API: get one sheet as JSON ────────────────────────────────────────────────
app.get("/api/sheet/:name", (req, res) => {
  try {
    const wb = getWorkbook();
    const ws = wb.Sheets[req.params.name];
    if (!ws) return res.status(404).json({ error: "Sheet not found" });
    res.json({ data: parseMonthSheet(ws) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── API: look up PO number ────────────────────────────────────────────────────
app.get("/api/po/:poNumber", (req, res) => {
  try {
    const poNumber = req.params.poNumber.trim().toUpperCase();
    const allData  = getAllMonthData();
    for (const [sheet, rows] of Object.entries(allData)) {
      const match = rows.find(r =>
        String(r["PO NUMBER"] || "").trim().toUpperCase() === poNumber
      );
      if (match) return res.json({ found: true, sheet, row: match });
    }
    res.json({ found: false });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── API: add row to a sheet ───────────────────────────────────────────────────
app.post("/api/sheet/:name/add", (req, res) => {
  try {
    const wb = getWorkbook();
    const name = req.params.name;

    // Find the sheet — match by name fragment
    const sheetName = wb.SheetNames.find(s =>
      s.toUpperCase().includes(name.toUpperCase()) || s.toUpperCase() === name.toUpperCase()
    );
    if (!sheetName) return res.status(404).json({ error: "Sheet not found: " + name });

    const ws   = wb.Sheets[sheetName];
    const rows = parseMonthSheet(ws);
    rows.push(req.body);

    // Rebuild sheet from scratch preserving header structure
    const headers = Object.keys(req.body);
    const newData  = [headers, ...rows.map(r => headers.map(h => r[h] || ""))];
    wb.Sheets[sheetName] = xlsx.utils.aoa_to_sheet(newData);

    xlsx.writeFile(wb, EXCEL_PATH);
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Dashboard cache ────────────────────────────────────────────────────────────
const CACHE_PATH = path.join(__dirname, "data", "dashboard-cache.json");

function readCache() {
  try {
    if (!fs.existsSync(CACHE_PATH)) return null;
    return JSON.parse(fs.readFileSync(CACHE_PATH, "utf8"));
  } catch(e) { return null; }
}

function writeCache(data) {
  try {
    fs.writeFileSync(CACHE_PATH, JSON.stringify(data), "utf8");
  } catch(e) { console.warn("Could not write dashboard cache:", e.message); }
}

function buildDashboardData() {
  const allData = getAllMonthData();

  const MONTH_ORDER = [
    "JAN","FEB","MAR","APR","MAY","JUN",
    "JUL","AUG","SEP","OCT","NOV","DEC"
  ];

  function toMonthKey(name) {
    const u = name.toUpperCase().replace(/\s/g, "");
    const fullNames = {
      "JANUARY":"JAN","FEBRUARY":"FEB","MARCH":"MAR","APRIL":"APR",
      "JUNE":"JUN","JULY":"JUL","AUGUST":"AUG","SEPTEMBER":"SEP",
      "OCTOBER":"OCT","NOVEMBER":"NOV","DECEMBER":"DEC",
    };
    for (const [full, short] of Object.entries(fullNames)) {
      if (u.includes(full)) return short;
    }
    for (const m of MONTH_ORDER) {
      if (u.includes(m)) return m;
    }
    return name.substring(0, 3).toUpperCase();
  }

  const monthMap    = {};
  const supplierMap = {};
  const deptMap     = {};
  const statusMap   = {};
  const allRows     = [];

  MONTH_ORDER.forEach(m => monthMap[m] = 0);

  for (const [sheetName, rows] of Object.entries(allData)) {
    const monthKey = toMonthKey(sheetName);
    rows.forEach(r => {
      const amt        = r["TOTAL AMOUNT"] > 0 ? r["TOTAL AMOUNT"] : (r["AMOUNT"] || 0);
      const isHeaderRow = !!(r["PO NUMBER"] || r["PR NO."]);
      const spendAmt   = isHeaderRow ? amt : 0;

      if (monthMap[monthKey] !== undefined) monthMap[monthKey] += spendAmt;

      const supplier = r["SUPPLIER'S NAME"] || "";
      if (supplier && isHeaderRow) supplierMap[supplier] = (supplierMap[supplier] || 0) + spendAmt;

      const dept = r["REQUESTING DEPT."] || "";
      if (dept && isHeaderRow) deptMap[dept] = (deptMap[dept] || 0) + spendAmt;

      const status = r["PURCHASE ORDER STATUS"] || "Unknown";
      if (status) statusMap[status] = (statusMap[status] || 0) + 1;

      allRows.push({ ...r, _sheet: sheetName });
    });
  }

  const monthlySpending = MONTH_ORDER.map(m => ({ month: m, total: monthMap[m] }));

  // Cost savings sheet
  const wb = getWorkbook();
  let savingsData = [];
  const savingsSheet = wb.SheetNames.find(s =>
    s.toLowerCase().includes("cost") || s.toLowerCase().includes("saving")
  );
  if (savingsSheet) {
    const raw = xlsx.utils.sheet_to_json(wb.Sheets[savingsSheet], { header:1, defval:null, raw:false });
    const hIdx = raw.findIndex(r => r && r.some(c => c && String(c).toUpperCase().includes("SUPPLIER")));
    if (hIdx >= 0) {
      const heads = raw[hIdx].map(h => h ? String(h).trim().toUpperCase() : "");
      for (let i = hIdx + 1; i < raw.length; i++) {
        const r = raw[i];
        if (!r || !r[0] || String(r[0]).startsWith("#")) continue;
        const get = name => {
          const idx = heads.indexOf(name.toUpperCase());
          const v = idx >= 0 ? r[idx] : null;
          if (v === null || String(v).startsWith("#")) return 0;
          return parseFloat(String(v).replace(/[^0-9.-]/g, "")) || 0;
        };
        savingsData.push({
          "SUPPLIER'S NAME":    String(r[0] || ""),
          "TOTAL AMOUNT":       get("TOTAL AMOUNT"),
          "ORIGINAL PRICE":     get("ORIGINAL PRICE"),
          "TOTAL COST SAVINGS": get("TOTAL COST SAVINGS"),
        });
      }
    }
  }
  if (!savingsData.length) {
    const sMap = {};
    allRows.forEach(r => {
      const s = r["SUPPLIER'S NAME"] || "Unknown";
      if (!sMap[s]) sMap[s] = { total:0, original:0 };
      sMap[s].total    += r["TOTAL AMOUNT"] || r["AMOUNT"] || 0;
      sMap[s].original += r["ORIGINAL PRICE"] || 0;
    });
    savingsData = Object.entries(sMap).map(([name, v]) => ({
      "SUPPLIER'S NAME":    name,
      "TOTAL AMOUNT":       v.total,
      "ORIGINAL PRICE":     v.original,
      "TOTAL COST SAVINGS": Math.max(0, v.original - v.total),
    }));
  }

  const normStatus = { Served:0, "For Delivery":0, Pending:0, Cancelled:0 };
  Object.entries(statusMap).forEach(([k, v]) => {
    const u = k.toLowerCase();
    if (u.includes("served") || u.includes("complete")) normStatus["Served"] += v;
    else if (u.includes("deliver")) normStatus["For Delivery"] += v;
    else if (u.includes("cancel"))  normStatus["Cancelled"] += v;
    else normStatus["Pending"] += v;
  });

  return {
    monthlySpending,
    supplierBreakdown: Object.entries(supplierMap)
      .map(([name, total]) => ({ name, total }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 10),
    deptSpending: Object.entries(deptMap)
      .map(([name, total]) => ({ name, total }))
      .sort((a, b) => b.total - a.total),
    statusCounts: normStatus,
    savingsData,
    rawRows: allRows.length,
    sheetsDetected: Object.keys(allData),
    cachedAt: Date.now(),
  };
}

// ── API: dashboard — serves cache instantly, rebuilds on ?refresh=1 ───────────
app.get("/api/dashboard", (req, res) => {
  try {
    const forceRefresh = req.query.refresh === "1";

    if (!forceRefresh) {
      const cached = readCache();
      if (cached) {
        console.log(`[Dashboard] Serving cache (${cached.rawRows} rows, saved ${new Date(cached.cachedAt).toLocaleTimeString()})`);
        return res.json({ ...cached, fromCache: true });
      }
      console.log("[Dashboard] No cache found — building from Excel...");
    } else {
      console.log("[Dashboard] Force refresh — rebuilding from Excel...");
    }

    const data = buildDashboardData();
    writeCache(data);
    console.log(`[Dashboard] Built and cached (${data.rawRows} rows)`);
    res.json({ ...data, fromCache: false });

  } catch(e) {
    console.error("Dashboard error:", e);
    // Try to serve stale cache if build fails
    const cached = readCache();
    if (cached) {
      console.warn("[Dashboard] Build failed, serving stale cache");
      return res.json({ ...cached, fromCache: true, stale: true });
    }
    res.status(500).json({ error: e.message });
  }
});

// ── API: upload scan image ────────────────────────────────────────────────────
app.post("/api/upload-scan", upload.single("image"), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file" });
  res.json({ success: true, path: req.file.path });
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, "0.0.0.0", () => {
  const ifaces = os.networkInterfaces();
  let ip = "localhost";
  Object.values(ifaces).forEach(iface =>
    iface.forEach(d => { if (d.family === "IPv4" && !d.internal) ip = d.address; })
  );
  console.log(`\n✅ Purchasing Tool running!`);
  console.log(`   Local:   http://localhost:${PORT}`);
  console.log(`   Network: http://${ip}:${PORT}  ← use this on phones\n`);

  // Pre-build dashboard cache in the background so first visitor doesn't wait.
  // If a cache already exists and is less than 60 minutes old, skip the rebuild.
  const existing = readCache();
  const cacheAge = existing ? (Date.now() - (existing.cachedAt || 0)) : Infinity;
  const sixtyMin = 60 * 60 * 1000;

  if (existing && cacheAge < sixtyMin) {
    const mins = Math.floor(cacheAge / 60000);
    console.log(`[Dashboard] Cache is ${mins}m old — skipping rebuild. Hit Refresh in the browser to force one.\n`);
  } else {
    console.log("[Dashboard] Building dashboard cache from monitoring.xlsx...");
    // Run async so the server starts immediately and isn't blocked
    setImmediate(() => {
      try {
        const data = buildDashboardData();
        writeCache(data);
        console.log(`[Dashboard] Cache ready — ${data.rawRows} rows from ${data.sheetsDetected.join(", ")}\n`);
      } catch(e) {
        console.warn(`[Dashboard] Cache build failed: ${e.message} — will retry on first request.\n`);
      }
    });
  }
});

// ── API: Export canvass sheet ─────────────────────────────────────────────────
// Copies TEMP.xlsx, fills in the data, saves as a new file with timestamp,
// and streams it back to the browser as a download.
app.post("/api/canvass/export", async (req, res) => {
  const xl   = require("xlsx");
  const xlsxPkg = require("xlsx");
  const openpyxl = null; // not available — use xlsx only

  // We'll use a Python script to do the openpyxl work since it handles
  // merged cells and row insertion far better than the xlsx npm package
  const { execFile } = require("child_process");
  const os2 = require("os");
  const { v4: uuidv4 } = require("crypto");

  const tmpJson = path.join(os2.tmpdir(), `canvass_${Date.now()}.json`);
  const outFile = path.join(__dirname, "data", `Canvass_${Date.now()}.xlsx`);

  try {
    fs.writeFileSync(tmpJson, JSON.stringify(req.body));

    await new Promise((resolve, reject) => {
      execFile("python3", [
        path.join(__dirname, "canvass_fill.py"),
        tmpJson,
        path.join(__dirname, "data", "TEMP.xlsx"),
        outFile,
      ], (err, stdout, stderr) => {
        if (err) { console.error("Python error:", stderr); return reject(new Error(stderr || err.message)); }
        resolve();
      });
    });

    fs.unlinkSync(tmpJson);

    res.download(outFile, `Canvass_${Date.now()}.xlsx`, err => {
      // Clean up after download
      try { fs.unlinkSync(outFile); } catch(e) {}
    });
  } catch(e) {
    try { fs.unlinkSync(tmpJson); } catch(_) {}
    res.status(500).json({ error: e.message });
  }
});
