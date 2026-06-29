const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const os = require("os");

const app = express();
const PORT = 3000;

// ── Middleware ──────────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, "public")));
app.use(express.json({ limit: "20mb" }));

// Multer: save uploaded scanned images to /data/scans/
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = path.join(__dirname, "data", "scans");
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    cb(null, `scan_${Date.now()}.png`);
  },
});
const upload = multer({ storage });

// ── Helper: get path to monitoring Excel ──────────────────────────────────
const EXCEL_PATH = path.join(__dirname, "data", "monitoring.xlsx");

function getWorkbook() {
  if (!fs.existsSync(EXCEL_PATH)) {
    // Create a blank workbook with correct sheets if none exists
    const wb = xlsx.utils.book_new();
    const monHeaders = [
      "PR DATE", "PR DATE RECEIVED", "PR NO.", "REQUESTING DEPT.",
      "PO DATE", "PO NUMBER", "END USER/S", "SUPPLIER'S NAME",
      "ITEM CODE", "ITEM DESCRIPTION", "SPECIFICATIONS", "QTY", "UoM",
      "UNIT PRICE", "AMOUNT", "TOTAL AMOUNT", "PR REQUIRED DATE",
      "DATE DELIVERED", "REMARKS", "PURCHASE ORDER STATUS",
      "ITEMS/SERVICES", "ORIGINAL PRICE", "TOTAL COST SAVINGS"
    ];
    const summaryHeaders = [
      "MONTH", "NO. OF PR ENDORSED", "PROCESSED PO (SERVED)",
      "PROCESSED PR (FOR PO)", "PROCESSED PO (WAITING FOR DELIVERY)",
      "CANCELLED PO", "TOTAL PO CREATED", "REMARKS"
    ];
    const savingsHeaders = [
      "SUPPLIER'S NAME", "TOTAL AMOUNT", "ORIGINAL PRICE", "TOTAL COST SAVINGS"
    ];

    // Create one monitoring sheet per month
    const months = [
      "JAN","FEB","MAR","APR","MAY","JUN",
      "JUL","AUG","SEP","OCT","NOV","DEC"
    ];
    months.forEach(m => {
      const ws = xlsx.utils.aoa_to_sheet([monHeaders]);
      xlsx.utils.book_append_sheet(wb, ws, m);
    });

    const summaryWs = xlsx.utils.aoa_to_sheet([summaryHeaders]);
    xlsx.utils.book_append_sheet(wb, summaryWs, "SUMMARY");

    const savingsWs = xlsx.utils.aoa_to_sheet([savingsHeaders]);
    xlsx.utils.book_append_sheet(wb, savingsWs, "COST SAVINGS");

    xlsx.writeFile(wb, EXCEL_PATH);
  }
  return xlsx.readFile(EXCEL_PATH);
}

// ── API: Get all sheet names ───────────────────────────────────────────────
app.get("/api/sheets", (req, res) => {
  try {
    const wb = getWorkbook();
    res.json({ sheets: wb.SheetNames });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: Get sheet data as JSON ────────────────────────────────────────────
app.get("/api/sheet/:name", (req, res) => {
  try {
    const wb = getWorkbook();
    const ws = wb.Sheets[req.params.name];
    if (!ws) return res.status(404).json({ error: "Sheet not found" });
    const data = xlsx.utils.sheet_to_json(ws, { defval: "" });
    res.json({ data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: Look up PO number across all monthly sheets ──────────────────────
app.get("/api/po/:poNumber", (req, res) => {
  try {
    const wb = getWorkbook();
    const poNumber = req.params.poNumber.trim().toUpperCase();
    const months = [
      "JAN","FEB","MAR","APR","MAY","JUN",
      "JUL","AUG","SEP","OCT","NOV","DEC"
    ];

    let found = null;
    for (const month of months) {
      const ws = wb.Sheets[month];
      if (!ws) continue;
      const rows = xlsx.utils.sheet_to_json(ws, { defval: "" });
      const match = rows.find(
        r => String(r["PO NUMBER"] || "").trim().toUpperCase() === poNumber
      );
      if (match) {
        found = { sheet: month, row: match };
        break;
      }
    }

    if (found) {
      res.json({ found: true, ...found });
    } else {
      res.json({ found: false });
    }
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: Add new row to a monthly sheet ───────────────────────────────────
app.post("/api/sheet/:name/add", (req, res) => {
  try {
    const wb = getWorkbook();
    const sheetName = req.params.name;
    const ws = wb.Sheets[sheetName];
    if (!ws) return res.status(404).json({ error: "Sheet not found" });

    const rows = xlsx.utils.sheet_to_json(ws, { defval: "" });
    rows.push(req.body);
    const newWs = xlsx.utils.json_to_sheet(rows);
    wb.Sheets[sheetName] = newWs;
    xlsx.writeFile(wb, EXCEL_PATH);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: Update existing row by PO number ─────────────────────────────────
app.put("/api/po/:poNumber", (req, res) => {
  try {
    const wb = getWorkbook();
    const poNumber = req.params.poNumber.trim().toUpperCase();
    const months = [
      "JAN","FEB","MAR","APR","MAY","JUN",
      "JUL","AUG","SEP","OCT","NOV","DEC"
    ];

    let updated = false;
    for (const month of months) {
      const ws = wb.Sheets[month];
      if (!ws) continue;
      const rows = xlsx.utils.sheet_to_json(ws, { defval: "" });
      const idx = rows.findIndex(
        r => String(r["PO NUMBER"] || "").trim().toUpperCase() === poNumber
      );
      if (idx !== -1) {
        rows[idx] = { ...rows[idx], ...req.body };
        wb.Sheets[month] = xlsx.utils.json_to_sheet(rows);
        updated = true;
        break;
      }
    }

    xlsx.writeFile(wb, EXCEL_PATH);
    res.json({ success: updated });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: Upload scanned image (for server-side processing if needed) ───────
app.post("/api/upload-scan", upload.single("image"), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file" });
  res.json({ success: true, path: req.file.path });
});

// ── API: Dashboard summary data ───────────────────────────────────────────
app.get("/api/dashboard", (req, res) => {
  try {
    const wb = getWorkbook();
    const months = [
      "JAN","FEB","MAR","APR","MAY","JUN",
      "JUL","AUG","SEP","OCT","NOV","DEC"
    ];

    const monthlySpending = [];
    const supplierMap = {};
    const deptMap = {};
    const statusCounts = { Served: 0, Pending: 0, Cancelled: 0, "For Delivery": 0 };

    months.forEach(month => {
      const ws = wb.Sheets[month];
      if (!ws) { monthlySpending.push({ month, total: 0 }); return; }
      const rows = xlsx.utils.sheet_to_json(ws, { defval: "" });
      let monthTotal = 0;

      rows.forEach(r => {
        const amt = parseFloat(r["TOTAL AMOUNT"]) || 0;
        monthTotal += amt;

        const supplier = r["SUPPLIER'S NAME"] || "Unknown";
        supplierMap[supplier] = (supplierMap[supplier] || 0) + amt;

        const dept = r["REQUESTING DEPT."] || "Unknown";
        deptMap[dept] = (deptMap[dept] || 0) + amt;

        const status = String(r["PURCHASE ORDER STATUS"] || "").trim();
        if (status.toLowerCase().includes("served")) statusCounts["Served"]++;
        else if (status.toLowerCase().includes("cancel")) statusCounts["Cancelled"]++;
        else if (status.toLowerCase().includes("deliver")) statusCounts["For Delivery"]++;
        else if (status) statusCounts["Pending"]++;
      });

      monthlySpending.push({ month, total: monthTotal });
    });

    // Cost savings
    const savingsWs = wb.Sheets["COST SAVINGS"];
    const savingsData = savingsWs
      ? xlsx.utils.sheet_to_json(savingsWs, { defval: "" })
      : [];

    res.json({
      monthlySpending,
      supplierBreakdown: Object.entries(supplierMap)
        .map(([name, total]) => ({ name, total }))
        .sort((a, b) => b.total - a.total)
        .slice(0, 10),
      deptSpending: Object.entries(deptMap)
        .map(([name, total]) => ({ name, total }))
        .sort((a, b) => b.total - a.total),
      statusCounts,
      savingsData,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: Serve config (API keys) to frontend ──────────────────────────────
app.get("/api/config", (req, res) => {
  try {
    const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, "data", "config.json"), "utf8"));
    console.log("[config] /api/config requested");
    console.log("[config] GEMINI_API_KEY present:", !!cfg.GEMINI_API_KEY);
    res.json({ GEMINI_API_KEY: cfg.GEMINI_API_KEY || "" });
  } catch(e) {
    console.error("[config] Failed to read config.json:", e.message);
    res.status(500).json({ GEMINI_API_KEY: "" });
  }
});
// ── Start ──────────────────────────────────────────────────────────────────
app.listen(PORT, "0.0.0.0", () => {
  const interfaces = os.networkInterfaces();
  let localIP = "localhost";
  Object.values(interfaces).forEach(iface => {
    iface.forEach(details => {
      if (details.family === "IPv4" && !details.internal) {
        localIP = details.address;
      }
    });
  });
  console.log(`\n✅ Purchasing Tool running!`);
  console.log(`   Local:   http://localhost:${PORT}`);
  console.log(`   Network: http://${localIP}:${PORT}  ← use this on phones\n`);
});
