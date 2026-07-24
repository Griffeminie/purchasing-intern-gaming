/**
 * server.js — Purchasing Tool backend
 * Added: PO file upload, serve, delete routes
 */

// ── Startup timing instrumentation (temporary, remove once diagnosed) ──────
const __bootStart = Date.now();
let __lastMark = __bootStart;
function bootLog(label) {
  const now = Date.now();
  console.log(`[BOOT] ${label} — step: ${now - __lastMark}ms | total: ${now - __bootStart}ms`);
  __lastMark = now;
}
bootLog("process start (server.js begin)");

const express    = require("express");
const multer     = require("multer");
const xlsx       = require("xlsx");
const path       = require("path");
const fs         = require("fs");
const os         = require("os");
bootLog("core requires (express, multer, xlsx, path, fs, os)");

// ── Cross-platform python command resolver ────────────────────────────────────
const { execFileSync } = require("child_process");
function resolvePythonCmd() {
  for (const cmd of ["python3", "python"]) {
    const t0 = Date.now();
    try {
      execFileSync(cmd, ["--version"], { stdio: "ignore" });
      console.log(`[BOOT]   tried "${cmd}" -> OK in ${Date.now() - t0}ms`);
      return cmd;
    } catch (e) {
      console.log(`[BOOT]   tried "${cmd}" -> FAILED in ${Date.now() - t0}ms`);
    }
  }
  throw new Error("No Python interpreter found (tried python3, python)");
}
const PYTHON_CMD = resolvePythonCmd();
bootLog("resolvePythonCmd() total");
console.log(`[Python] Using command: ${PYTHON_CMD}`);

const {
  insertRow, updateRow, updateRowFile, deleteRow,
  getByMonth, getMonths, findByPoNumber, buildDashboard, db,
  getAllSuppliers, importSuppliers, insertSupplier, updateSupplier,
  deleteSupplier, supplierCount,
} = require("./db");
bootLog('require("./db")  <- SQLite open/schema/migrations happen here');
const { buildMonthWorkbook, buildYearWorkbook } = require("./xlsx-template");
bootLog('require("./xlsx-template")');

const app  = express();
const PORT = 3000;

const session = require('express-session');
const bcrypt  = require('bcrypt');
const users   = require('./data/users.js'); // your bcrypt-hashed users object

// ── Session (must come before any routes that need req.session) ──────────────
app.use(session({
  secret: 'change-this-to-a-random-long-string', // move to config.json / env var
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }, // 8 hours
}));

function requireAuth(req, res, next) {
  if (req.session && req.session.user) return next();

  const ALWAYS_ALLOWED = ['/', '/index.html', '/api/login'];
  const STATIC_ASSET   = /\.(css|js|png|jpe?g|svg|ico|gif|woff2?)$/i;

  if (ALWAYS_ALLOWED.includes(req.path) || STATIC_ASSET.test(req.path)) {
    return next();
  }
  if (req.path.startsWith('/api/')) {
    return res.status(401).json({ error: 'Not authenticated' });
  }
  // Any other page (dashboard.html, monitoring.html, etc.) -> bounce home,
  // remembering where they were trying to go
  return res.redirect(`/?next=${encodeURIComponent(req.path)}`);
}
app.use(requireAuth);

// ── NOW it's safe to serve static files and parse JSON ────────────────────────
app.use(express.static(path.join(__dirname, "public")));
app.use(express.json({ limit: "20mb" }));

app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};
  const hash = users[username];
  if (hash && await bcrypt.compare(password, hash)) {
    req.session.user = username;
    return res.json({ success: true });
  }
  res.status(401).json({ success: false, error: 'Invalid username or password' });
});

app.post('/api/logout', (req, res) => {
  req.session.destroy(() => res.json({ success: true }));
});

// ── Canvass workbook directory ────────────────────────────────────────────────
const CANVASS_DIR = "X:\\NEW PURCHASING TEAM\\CANVASS SHEET";

// ── Gemini config (server-side only — never sent to the browser) ────────────
const GEMINI_MODEL = "gemini-2.5-flash-lite";
function getGeminiKey() {
  try {
    const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, "data", "config.json"), "utf8"));
    return cfg.GEMINI_API_KEY || "";
  } catch (e) {
    return "";
  }
}

const EXTRACT_PROMPT = `
You are extracting information from a supplier quotation document.
Return ONLY valid JSON with EXACTLY this structure — no markdown, no backticks, no explanations:

{
  "supplierName": "",
  "location": "",
  "contactPerson": "",
  "contactNumber": "",
  "items": [
    { "itemNo": "", "description": "", "qty": "", "unit": "", "unitPrice": "", "total": "" }
  ],
  "deliveryCharge": "",
  "discount": "",
  "grandTotal": "",
  "creditTerms": "",
  "vat": "",
  "availability": "",
  "deliveryAddress": ""
}

Rules:
- Extract every line item. Preserve text exactly as written.
- supplierName: the company/business name issuing the quotation.
- location: the SUPPLIER's own business address/city, usually printed in the letterhead near the supplier name or logo. This is different from deliveryAddress (the customer's delivery address, usually printed lower in the document).
- contactPerson: the name of the person to contact (signatory, sales rep, prepared-by, etc.), if printed on the document.
- contactNumber: a phone/mobile number associated with the supplier or contact person, if printed on the document.
- Do not invent values. If a value cannot be found, use an empty string "" — never guess or fabricate a name, address, or number.
- NEVER perform arithmetic or derive a number that is not itself printed on the document. This applies especially to "vat", "grandTotal", "deliveryCharge", and "discount" — if the document only says a word or phrase like "VAT INCLUSIVE", "STANDARD RATE", "12%", or "FREE" with no specific peso figure next to it, output that exact phrase as plain text. Do NOT calculate what 12% of some other number would be, do NOT back-compute a VAT amount from a grand total, and do NOT fill in a number that "should" be there. Only output a numeric value for these fields if that exact number is visibly printed on the document.
- Output plain human-readable text only. Never include special/control tokens such as <bos>, <eos>, <pad>, <unk>, <start_of_turn>, <end_of_turn> or any other bracketed tokens anywhere in a value.
- Return ONLY JSON. No markdown. No explanations. No code blocks.
- itemNo: always output an empty string "" for this field, no matter what. Some documents have a leading number column that looks like an item index but is actually something else (a quantity, a product code, a reference ID) depending on the layout — you cannot reliably tell them apart, so don't try. Item numbering is assigned separately by the system, not extracted by you.
`;

const ACCOUNTABILITY_PROMPT = `
Extract all rows from this accountability form into a JSON array. Each object must have these exact keys:
- "store_name": the store name exactly as written
- "m3_code": the M3 code or store code if visible (null if not present)
- "accountable_person": the full name of the accountable/custodian employee

Return ONLY valid JSON. No markdown, no code fences, no explanation. Example:
[{"store_name":"CSI Dagupan 10016","m3_code":"10016","accountable_person":"Janette Ferrer"}]
`;

const MATCH_PROMPT = `
You are given items from MULTIPLE supplier quotations for the SAME canvass request.
Different suppliers describe the same physical product differently
(e.g. "BATTERY EVEREADY AA", "BATTERY AA ENERGIZER", "GP ULTRA AA (4 PCS) ALKALINE BATTERY"
are all the same product: AA batteries).

Group items that refer to the SAME physical product/spec across suppliers, even if brand,
wording, or order differs. Do NOT group genuinely different products or specs
(e.g. different paper sizes, different column counts on notebooks).

For each group, pick ONE short, generic, brand-neutral canonical name (no brand names,
no supplier item codes), e.g. "BATTERY AA", "BOND PAPER SHORT 70GSM", "WHITE FOLDER LONG",
"COLUMNAR NOTEBOOK 8 COLUMN".

Return ONLY valid JSON, no markdown, no explanations, in exactly this structure:
[
  { "canonicalName": "BATTERY AA", "itemIds": ["s1i3","s2i7","s3i1"] }
]

Every item id in the input must appear in exactly one group. If an item has no match,
put it alone in its own group with a clean canonical name.

Items:
`;

async function callGeminiRaw(payload) {
  const key = getGeminiKey();
  if (!key) throw new Error("Gemini API key not configured on server");
  const resp = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${key}`,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) }
  );
  const data = await resp.json();
  if (data.error) throw new Error(data.error.message);
  return data;
}

// ── PO PDF upload storage ─────────────────────────────────────────────────────
const poFilesDir = path.join(__dirname, "data", "po_files");
fs.mkdirSync(poFilesDir, { recursive: true });

const poStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, poFilesDir),
  filename: (req, file, cb) => {
    // Keep original name but prefix with rowId for uniqueness
    const rowId = req.params.id || "unknown";
    const ext   = path.extname(file.originalname) || ".pdf";
    const base  = path.basename(file.originalname, ext)
                      .replace(/[^a-zA-Z0-9._-]/g, "_")
                      .substring(0, 80);
    cb(null, `PO_${rowId}_${base}${ext}`);
  },
});
const poUpload = multer({
  storage: poStorage,
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 MB
  fileFilter: (req, file, cb) => {
    const ok = file.mimetype === "application/pdf" ||
               file.originalname.toLowerCase().endsWith(".pdf");
    if (!ok) return cb(new Error("Only PDF files are allowed"));
    cb(null, true);
  },
});

// ── Scan image upload storage ─────────────────────────────────────────────────
const scanStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = path.join(__dirname, "data", "scans");
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => { cb(null, `scan_${Date.now()}.png`); },
});
const upload = multer({ storage: scanStorage });

// ── Excel headers for export ──────────────────────────────────────────────────
const HEADERS = [
  "PR DATE", "PR DATE RECEIVED", "PR NO.", "REQUESTING DEPT.",
  "PO DATE", "PO NUMBER", "END USER/S", "SUPPLIER'S NAME",
  "ITEM CODE", "ITEM DESCRIPTION", "SPECIFICATIONS", "QTY", "UoM",
  "UNIT PRICE", "AMOUNT", "TOTAL AMOUNT", "PAYMENT TERMS",
  "PR REQUIRED DATE", "DATE DELIVERED", "REMARKS",
  "PURCHASE ORDER STATUS", "ITEMS/SERVICES", "ORIGINAL PRICE",
  "TOTAL COST SAVINGS",
];

const MONTH_ORDER = [
  "JAN","FEB","MAR","APR","MAY","JUN",
  "JUL","AUG","SEP","OCT","NOV","DEC",
];

function rowsToSheet(rows) {
  const data = [
    HEADERS,
    ...rows.map(r => HEADERS.map(h => r[h] ?? "")),
  ];
  return xlsx.utils.aoa_to_sheet(data);
}

function sendWorkbook(res, wb, filename) {
  const buf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buf);
}

// Same as sendWorkbook, but for an exceljs Workbook (used by the template-based
// monitoring exports in xlsx-template.js, since exceljs preserves formatting
// that the plain "xlsx" package above can't).
async function sendExcelJSWorkbook(res, wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  res.setHeader("Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(Buffer.from(buf));
}

// ═══════════════════════════════════════════════════════════════════════════════
// API ROUTES
// ═══════════════════════════════════════════════════════════════════════════════

// ── Config: tell frontend whether AI is ready — key itself never leaves the server ──
app.get("/api/config", (req, res) => {
  res.json({ aiReady: !!getGeminiKey() });
});

// ── Gemini: extract one quotation ─────────────────────────────────────────────
app.post("/api/gemini/extract", async (req, res) => {
  try {
    const { mimeType, base64 } = req.body;
    if (!mimeType || !base64) return res.status(400).json({ error: "mimeType and base64 required" });
    const data = await callGeminiRaw({
      contents: [{ parts: [{ text: EXTRACT_PROMPT }, { inline_data: { mime_type: mimeType, data: base64 } }] }],
    });
    res.json(data);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Gemini: extract one accountability form (store/M3 code/accountable person) ──
app.post("/api/gemini/extract-accountability", async (req, res) => {
  try {
    const { mimeType, base64 } = req.body;
    if (!mimeType || !base64) return res.status(400).json({ error: "mimeType and base64 required" });
    const data = await callGeminiRaw({
      contents: [{ parts: [{ text: ACCOUNTABILITY_PROMPT }, { inline_data: { mime_type: mimeType, data: base64 } }] }],
    });
    res.json(data);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Gemini: match items across suppliers ──────────────────────────────────────
app.post("/api/gemini/match", async (req, res) => {
  try {
    const { items } = req.body;
    if (!items) return res.status(400).json({ error: "items required" });
    const data = await callGeminiRaw({
      contents: [{ parts: [{ text: MATCH_PROMPT + JSON.stringify(items, null, 2) }] }],
    });
    res.json(data);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Suppliers — persisted in SQLite (mirrors the monitoring data pattern) ─────
app.get("/api/suppliers", (req, res) => {
  try {
    const suppliers = getAllSuppliers();
    console.log(`[Suppliers] Loaded ${suppliers.length} records from SQLite`);
    res.json({ suppliers });
  } catch (e) {
    console.error("[Suppliers] ERROR reading from SQLite:", e.message);
    res.status(500).json({ suppliers: [], error: e.message });
  }
});

// ── Public health check — no sheet data, just confirms server is up ──────────
app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    loggedIn: !!(req.session && req.session.user),
    username: req.session.user || null,
  });
});

// ── List existing department canvass workbooks ────────────────────────────────
// GET /api/canvass/books
// Returns the .xlsx files sitting in CANVASS_DIR so the frontend can offer
// them as a dropdown instead of the user typing a path.
app.get("/api/canvass/books", (req, res) => {
  try {
    if (!fs.existsSync(CANVASS_DIR)) {
      console.warn(`[Canvass Books] Directory not found: ${CANVASS_DIR}`);
      return res.json({ books: [], error: `Directory not found: ${CANVASS_DIR}` });
    }
    const files = fs.readdirSync(CANVASS_DIR)
      .filter(f => f.toLowerCase().endsWith(".xlsx") && !f.startsWith("~$")) // skip Excel lock files
      .sort((a, b) => a.localeCompare(b));
    res.json({ books: files });
  } catch (e) {
    console.error("[Canvass Books] ERROR:", e.message);
    res.status(500).json({ books: [], error: e.message });
  }
});

// Bulk import — replaces the whole supplier list. Accepts { suppliers: [...] }
// or a raw array, same shape the front-end's "Import JSON" flow already parses.
app.post("/api/suppliers/import", (req, res) => {
  try {
    const body = req.body || {};
    const list = Array.isArray(body) ? body : body.suppliers;
    if (!Array.isArray(list)) {
      return res.status(400).json({ error: 'Expected a "suppliers" array (or a top-level array) of records.' });
    }
    const count = importSuppliers(list);
    console.log(`[Suppliers] Imported ${count} records into SQLite`);
    res.json({ success: true, count });
  } catch (e) {
    console.error("[Suppliers] ERROR importing:", e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/supplier", (req, res) => {
  try {
    const id = insertSupplier(req.body);
    res.json({ success: true, id });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put("/api/supplier/:id", (req, res) => {
  try {
    updateSupplier(parseInt(req.params.id), req.body);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete("/api/supplier/:id", (req, res) => {
  try {
    deleteSupplier(parseInt(req.params.id));
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});
app.get("/api/sheets", (req, res) => {
  try {
    res.json({ sheets: getMonths() });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/sheet/:name", (req, res) => {
  try {
    const month = req.params.name.toUpperCase().trim();
    res.json({ data: getByMonth(month) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/po/:poNumber", (req, res) => {
  try {
    const row = findByPoNumber(req.params.poNumber.trim());
    if (row) res.json({ found: true, sheet: row.month, row });
    else     res.json({ found: false });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/api/sheet/:name/add", (req, res) => {
  try {
    const raw   = req.params.name.toUpperCase();
    const month = MONTH_ORDER.find(m => raw.includes(m)) || raw.substring(0, 3);
    const id    = insertRow(req.body, month);
    res.json({ success: true, id });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Last Purchase lookup (for canvass "Add Last Purchase" per item) ─────────
// GET /api/purchase-history/search?q=some text
// Searches item description, supplier name, and PO number across ALL months.
// Returns only the 4 fields the canvass lookup popup needs.
app.get("/api/purchase-history/search", (req, res) => {
  try {
    const q = (req.query.q || "").trim();
    let rows;
    if (q) {
      rows = db.prepare(`
        SELECT po_date, supplier_name, po_number, unit_price
        FROM purchase_orders
        WHERE item_description LIKE ? OR supplier_name LIKE ? OR po_number LIKE ?
        ORDER BY po_date DESC
        LIMIT 50
      `).all(`%${q}%`, `%${q}%`, `%${q}%`);
    } else {
      rows = db.prepare(`
        SELECT po_date, supplier_name, po_number, unit_price
        FROM purchase_orders
        ORDER BY po_date DESC
        LIMIT 50
      `).all();
    }
    res.json({
      results: rows.map(r => ({
        poDate:    r.po_date || "",
        supplier:  r.supplier_name || "",
        poNumber:  r.po_number || "",
        unitPrice: r.unit_price ?? "",
      })),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.put("/api/row/:id", (req, res) => {
  try {
    const id    = parseInt(req.params.id);
    const month = (req.body.month || "").toUpperCase() || "JAN";
    updateRow(id, req.body, month);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete("/api/row/:id", (req, res) => {
  try {
    const id  = parseInt(req.params.id);
    // Also delete the linked PDF if it exists
    const row = db.prepare("SELECT po_file FROM purchase_orders WHERE id = ?").get(id);
    if (row && row.po_file) {
      const filePath = path.join(poFilesDir, row.po_file);
      try { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); } catch(e) {}
    }
    deleteRow(id);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── PO FILE: Upload PDF and link to a row ─────────────────────────────────────
// POST /api/row/:id/file  (multipart field name: "po_file")
app.post("/api/row/:id/file", poUpload.single("po_file"), (req, res) => {
  try {
    const id = parseInt(req.params.id);
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    // If this row already had a file, delete the old one
    const existing = db.prepare("SELECT po_file FROM purchase_orders WHERE id = ?").get(id);
    if (existing && existing.po_file) {
      const old = path.join(poFilesDir, existing.po_file);
      try { if (fs.existsSync(old)) fs.unlinkSync(old); } catch(e) {}
    }

    updateRowFile(id, req.file.filename);
    console.log(`[PO File] Row ${id} → ${req.file.filename}`);
    res.json({ success: true, filename: req.file.filename });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── PO FILE: View/download a PDF ──────────────────────────────────────────────
// GET /api/row/:id/file
app.get("/api/row/:id/file", (req, res) => {
  try {
    const id  = parseInt(req.params.id);
    const row = db.prepare("SELECT po_file FROM purchase_orders WHERE id = ?").get(id);
    if (!row || !row.po_file) return res.status(404).json({ error: "No file attached" });

    const filePath = path.join(poFilesDir, row.po_file);
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: "File not found on disk" });

    // Send inline so the browser can display it in an iframe
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${row.po_file}"`);
    fs.createReadStream(filePath).pipe(res);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── PO FILE: Remove a PDF from a row ─────────────────────────────────────────
// DELETE /api/row/:id/file
app.delete("/api/row/:id/file", (req, res) => {
  try {
    const id  = parseInt(req.params.id);
    const row = db.prepare("SELECT po_file FROM purchase_orders WHERE id = ?").get(id);
    if (row && row.po_file) {
      const filePath = path.join(poFilesDir, row.po_file);
      try { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); } catch(e) {}
    }
    updateRowFile(id, "");
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Dashboard ─────────────────────────────────────────────────────────────────
app.get("/api/dashboard", (req, res) => {
  try {
    res.json({ ...buildDashboard(), fromCache: false });
  } catch (e) {
    console.error("Dashboard error:", e);
    res.status(500).json({ error: e.message });
  }
});

// ── Export: one month (uses data/monitoring_template.xlsx as the layout) ──────
app.get("/api/export/month/:month", async (req, res) => {
  try {
    const month = req.params.month.toUpperCase().trim();
    const rows  = getByMonth(month);
    if (!rows.length) return res.status(404).json({ error: `No data for ${month}` });

    const year = new Date().getFullYear();
    const wb   = await buildMonthWorkbook(month, year, rows);
    await sendExcelJSWorkbook(res, wb, `PO_Monitoring_${month}_${year}.xlsx`);
  } catch (e) {
    console.error("Month export error:", e);
    res.status(500).json({ error: e.message });
  }
});

// ── Export: full year (uses data/monitoring_template.xlsx per month sheet) ────
app.get("/api/export/year", async (req, res) => {
  try {
    const year = new Date().getFullYear();
    const monthsWithRows = [];
    for (const month of MONTH_ORDER) {
      const rows = getByMonth(month);
      if (rows.length) monthsWithRows.push({ month, rows });
    }
    if (!monthsWithRows.length) return res.status(404).json({ error: "No data to export." });

    const wb = await buildYearWorkbook(monthsWithRows, year);

    const summary = wb.addWorksheet("Summary");
    summary.addRow([
      "Month", "No. of PRs", "Processed PO (Served)", "Processed PR (For PO)",
      "Processed PO (Waiting for Delivery)", "Cancelled PO",
      "Total PO Created", "Remarks",
    ]);
    summary.getRow(1).font = { bold: true };

    await sendExcelJSWorkbook(res, wb, `PO_Monitoring_Full_${year}.xlsx`);
  } catch (e) {
    console.error("Year export error:", e);
    res.status(500).json({ error: e.message });
  }
});

// ── Export: filtered ──────────────────────────────────────────────────────────
app.get("/api/export/filtered", (req, res) => {
  try {
    const { month, supplier, dept, status } = req.query;
    const conditions = [];
    const params     = [];
    if (month)    { conditions.push("UPPER(month) = UPPER(?)");         params.push(month); }
    if (supplier) { conditions.push("supplier_name LIKE ?");            params.push(`%${supplier}%`); }
    if (dept)     { conditions.push("requesting_dept LIKE ?");          params.push(`%${dept}%`); }
    if (status)   { conditions.push("UPPER(po_status) LIKE UPPER(?)"); params.push(`%${status}%`); }

    const where  = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";
    const rows   = db.prepare(`SELECT * FROM purchase_orders ${where} ORDER BY id ASC`).all(...params);
    if (!rows.length) return res.status(404).json({ error: "No rows match." });

    const apiRows = rows.map(r => ({
      "PR DATE": r.pr_date, "PR DATE RECEIVED": r.pr_date_received,
      "PR NO.": r.pr_no, "REQUESTING DEPT.": r.requesting_dept,
      "PO DATE": r.po_date, "PO NUMBER": r.po_number,
      "END USER/S": r.end_user, "SUPPLIER'S NAME": r.supplier_name,
      "ITEM CODE": r.item_code, "ITEM DESCRIPTION": r.item_description,
      "SPECIFICATIONS": r.specifications, "QTY": r.qty, "UoM": r.uom,
      "UNIT PRICE": r.unit_price, "AMOUNT": r.amount, "TOTAL AMOUNT": r.total_amount,
      "PAYMENT TERMS": r.payment_terms, "PR REQUIRED DATE": r.pr_required_date,
      "DATE DELIVERED": r.date_delivered, "REMARKS": r.remarks,
      "PURCHASE ORDER STATUS": r.po_status, "ITEMS/SERVICES": r.items_services,
      "ORIGINAL PRICE": r.original_price, "TOTAL COST SAVINGS": r.total_cost_savings,
    }));

    const wb = xlsx.utils.book_new();
    const sheetName = [month,supplier,dept,status].filter(Boolean).join("_") || "Filtered";
    xlsx.utils.book_append_sheet(wb, rowsToSheet(apiRows), sheetName.substring(0,31));
    sendWorkbook(res, wb, `PO_Export_${sheetName}_${Date.now()}.xlsx`);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Scan image upload ─────────────────────────────────────────────────────────
app.post("/api/upload-scan", upload.single("image"), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file" });
  res.json({ success: true, path: req.file.path });
});

// ── Canvass export (Python) ───────────────────────────────────────────────────
// Two modes, chosen by whether the request includes "targetBook":
//   - No targetBook: build a standalone .xlsx and send it back for download
//     (original behaviour).
//   - targetBook set: it's an absolute path to an existing department
//     workbook (can live anywhere on disk / a shared folder, not just
//     data/) — the new canvass gets appended to it as the LAST sheet and
//     the file is updated IN PLACE. Nothing is downloaded in this mode.
app.post("/api/canvass/export", async (req, res) => {
  const { execFile } = require("child_process");
  const { targetBook, sheetName, ...canvassData } = req.body;

  const appendMode = !!(targetBook && String(targetBook).trim());
  let resolvedTarget = null;

  if (appendMode) {
    const fileName = String(targetBook).trim();
    // Only a bare filename is accepted from the client — always resolved
    // against CANVASS_DIR server-side, never an arbitrary path, and
    // blocked from escaping that folder (no slashes/backslashes/"..").
    if (/[\\/]/.test(fileName) || fileName.includes("..")) {
      return res.status(400).json({ error: "Invalid file name." });
    }
    resolvedTarget = path.join(CANVASS_DIR, fileName);
    if (!fs.existsSync(resolvedTarget)) {
      return res.status(404).json({ error: `Workbook not found: ${resolvedTarget}` });
    }
  }

  const tmpJson = path.join(os.tmpdir(), `canvass_${Date.now()}.json`);
  // Scratch file used ONLY to build the single new sheet — never the same
  // path as resolvedTarget, so the department book is never touched until
  // canvass_fill.py explicitly opens it to append + save in place.
  const scratchOut = appendMode
    ? path.join(os.tmpdir(), `canvass_scratch_${Date.now()}.xlsx`)
    : path.join(__dirname, "data", `Canvass_${Date.now()}.xlsx`);

  try {
    fs.writeFileSync(tmpJson, JSON.stringify(canvassData));

    const args = [
      path.join(__dirname, "canvass_fill.py"),
      tmpJson,
      path.join(__dirname, "data", "TEMP.xlsx"),
      scratchOut,
    ];
    if (appendMode) {
      args.push(resolvedTarget);
      args.push(sheetName || canvassData.title || "Canvass");
    }

    await new Promise((resolve, reject) => {
      execFile(PYTHON_CMD, args, (err, stdout, stderr) => {
        if (err) { return reject(new Error(stderr || err.message)); }
        resolve();
      });
    });
    fs.unlinkSync(tmpJson);

    if (appendMode) {
      try { fs.unlinkSync(scratchOut); } catch(e) {}
      console.log(`[Canvass] Appended sheet "${sheetName || canvassData.title}" to ${resolvedTarget}`);
      res.json({ success: true, appendedTo: resolvedTarget, sheet: sheetName || canvassData.title || "Canvass" });
    } else {
      res.download(scratchOut, `Canvass_${Date.now()}.xlsx`, err => {
        try { fs.unlinkSync(scratchOut); } catch(e) {}
      });
    }
  } catch (e) {
    try { fs.unlinkSync(tmpJson); } catch(_) {}
    try { if (appendMode) fs.unlinkSync(scratchOut); } catch(_) {}
    res.status(500).json({ error: e.message });
  }
});



// ═══════════════════════════════════════════════════════════════════════════════
// PO FOLDER AUTO-SCAN
// POST /api/po-files/scan
// Body: { folderPath: "C:\\Users\\ITSU\\Desktop\\PO Files" }
//
// Logic:
//   1. Read all .pdf files in the given folder
//   2. Extract PO number from filename using pattern: PO#[NUMBER]
//   3. Look up that PO number in the database
//   4. If found: copy the file to data/po_files/ and link it to the row
//   5. Return a report of matched / unmatched files
// ═══════════════════════════════════════════════════════════════════════════════
app.post("/api/po-files/scan", async (req, res) => {
  const { folderPath } = req.body;
  if (!folderPath) return res.status(400).json({ error: "folderPath is required" });

  // Normalise path — handle both forward and back slashes
  const folder = path.resolve(folderPath.trim());

  if (!fs.existsSync(folder)) {
    return res.status(404).json({ error: `Folder not found: ${folder}` });
  }

  let files;
  try {
    files = fs.readdirSync(folder).filter(f => f.toLowerCase().endsWith(".pdf"));
  } catch (e) {
    return res.status(500).json({ error: `Cannot read folder: ${e.message}` });
  }

  if (!files.length) {
    return res.json({ matched: [], unmatched: [], total: 0, message: "No PDF files found in that folder." });
  }

  // Extract PO number from filename
  // Handles: PO#400085166, PO# 400085166, PO #400085166, PO-400085166
  function extractPoNumber(filename) {
    // Try PO# pattern first (matches your files: "CCI - PO#400085166 PRICOM")
    let m = filename.match(/PO[#\-\s#]+(\d{6,})/i);
    if (m) return m[1];
    // Fallback: any sequence of 6+ digits after "PO"
    m = filename.match(/PO[^0-9]*(\d{6,})/i);
    if (m) return m[1];
    return null;
  }

  const matched   = [];
  const unmatched = [];

  for (const filename of files) {
    const poNum = extractPoNumber(filename);

    if (!poNum) {
      unmatched.push({ filename, reason: "Could not extract PO number from filename" });
      continue;
    }

    // Look up in DB — try exact match first, then partial
    let row = db.prepare(
      "SELECT id, po_number, supplier_name, po_file FROM purchase_orders WHERE TRIM(po_number) = ?"
    ).get(poNum);

    // If not found, try stripping leading zeros or matching the tail
    if (!row) {
      row = db.prepare(
        "SELECT id, po_number, supplier_name, po_file FROM purchase_orders WHERE TRIM(po_number) LIKE ?"
      ).get(`%${poNum}%`);
    }

    if (!row) {
      unmatched.push({ filename, poExtracted: poNum, reason: "No matching PO number found in database" });
      continue;
    }

    // Copy the file to our managed po_files folder (don't delete the original)
    const srcPath  = path.join(folder, filename);
    const destName = `PO_${row.id}_${filename.replace(/[^a-zA-Z0-9._\-# ]/g, "_")}`;
    const destPath = path.join(poFilesDir, destName);

    try {
      fs.copyFileSync(srcPath, destPath);
    } catch (e) {
      unmatched.push({ filename, poExtracted: poNum, reason: `File copy failed: ${e.message}` });
      continue;
    }

    // If row already had a different file, delete the old one
    if (row.po_file && row.po_file !== destName) {
      const old = path.join(poFilesDir, row.po_file);
      try { if (fs.existsSync(old)) fs.unlinkSync(old); } catch(e) {}
    }

    // Link to DB row
    db.prepare("UPDATE purchase_orders SET po_file = ? WHERE id = ?").run(destName, row.id);

    matched.push({
      filename,
      poExtracted: poNum,
      dbPoNumber:  row.po_number,
      supplier:    row.supplier_name,
      rowId:       row.id,
    });
  }

  console.log(`[Scan] Folder: ${folder} | PDFs: ${files.length} | Matched: ${matched.length} | Unmatched: ${unmatched.length}`);

  res.json({
    total:     files.length,
    matched:   matched.length,
    unmatched: unmatched.length,
    details:   { matched, unmatched },
    folderPath: folder,
  });
});


// ── Scanner: extract PO PDFs via Python ──────────────────────────────────────
// POST /api/scanner/extract
// Accepts multipart upload of one or more PDFs
const scanPdfStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = path.join(__dirname, "data", "scanner_uploads");
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    cb(null, `scan_${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g,"_")}`);
  },
});
const scanUpload = multer({
  storage: scanPdfStorage,
  fileFilter: (req, file, cb) => {
    const ok = file.mimetype === "application/pdf" || file.originalname.toLowerCase().endsWith(".pdf");
    cb(null, ok);
  },
});

app.post("/api/scanner/extract", scanUpload.array("pdfs", 20), async (req, res) => {
  if (!req.files || !req.files.length)
    return res.status(400).json({ error: "No PDF files uploaded" });

  const { execFile } = require("child_process");
  const filePaths = req.files.map(f => f.path);

  try {
    const result = await new Promise((resolve, reject) => {
      const py = execFile(
        PYTHON_CMD,
        [path.join(__dirname, "scanner_api.py")],
        { maxBuffer: 10 * 1024 * 1024 },
        (err, stdout, stderr) => {
          if (err) return reject(new Error(stderr || err.message));
          try { resolve(JSON.parse(stdout)); }
          catch(e) { reject(new Error("Invalid JSON from scanner_api.py: " + stdout.substring(0, 200))); }
        }
      );
      py.stdin.write(JSON.stringify(filePaths));
      py.stdin.end();
    });

    // Clean up temp uploads after extraction
    filePaths.forEach(p => { try { fs.unlinkSync(p); } catch(e) {} });

    res.json(result);
  } catch(e) {
    filePaths.forEach(p => { try { fs.unlinkSync(p); } catch(e) {} });
    res.status(500).json({ error: e.message });
  }
});

// POST /api/scanner/save — save confirmed rows into SQLite
app.post("/api/scanner/save", (req, res) => {
  try {
    const { rows, month } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: "No rows to save" });
    const targetMonth = (month || new Date().toLocaleString("en-US",{month:"short"}).toUpperCase());
    let saved = 0;
    for (const row of rows) {
      insertRow(row, targetMonth);
      saved++;
    }
    res.json({ success: true, saved });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, "0.0.0.0", () => {
  bootLog("app.listen callback fired — SERVER READY");
  const ifaces = os.networkInterfaces();
  let ip = "localhost";
  Object.values(ifaces).forEach(iface =>
    iface.forEach(d => { if (d.family === "IPv4" && !d.internal) ip = d.address; })
  );
  const count = db.prepare("SELECT COUNT(*) AS cnt FROM purchase_orders").get();
  console.log(`\nMainframe initialized`);
  console.log(`Local:   http://localhost:${PORT}`);
  console.log(`Network: http://${ip}:${PORT}\n`);
  console.log(`Database: ${count.cnt} rows in SQLite`);
  console.log(`PO Files: ${poFilesDir}\n`);
});