/**
 * migrate.js  — fixed for unmerged Excel format
 * Run: node migrate.js --clear
 */

const xlsx  = require("xlsx");
const path  = require("path");
const fs    = require("fs");
const { db, insertMany } = require("./db");

const EXCEL_PATH  = path.join(__dirname, "data", "monitoring.xlsx");
const DRY_RUN     = process.argv.includes("--dry-run");
const CLEAR_FIRST = process.argv.includes("--clear");

const MONTH_PATTERNS = [
  "JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE",
  "JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER",
  "JAN","FEB","MAR","APR","MAY","JUN",
  "JUL","AUG","SEP","OCT","NOV","DEC",
];

const FULL_TO_SHORT = {
  "JANUARY":"JAN","FEBRUARY":"FEB","MARCH":"MAR","APRIL":"APR",
  "JUNE":"JUN","JULY":"JUL","AUGUST":"AUG","SEPTEMBER":"SEP",
  "OCTOBER":"OCT","NOVEMBER":"NOV","DECEMBER":"DEC",
};

function isMonthSheet(name) {
  const u = name.toUpperCase().replace(/[\s0-9]/g, "");
  return MONTH_PATTERNS.some(m => u.includes(m));
}

function toMonthKey(name) {
  const u = name.toUpperCase().replace(/[\s0-9]/g, "");
  for (const [full, short] of Object.entries(FULL_TO_SHORT)) {
    if (u.includes(full)) return short;
  }
  if (u.includes("MAY")) return "MAY";
  for (const m of ["JAN","FEB","MAR","APR","JUN",
                    "JUL","AUG","SEP","OCT","NOV","DEC"]) {
    if (u.includes(m)) return m;
  }
  return name.substring(0, 3).toUpperCase();
}

// ── Convert Excel serial date to readable string ──────────────────────────────
function excelDateToString(val) {
  if (!val) return "";
  // Already a string date
  if (typeof val === "string" && val.trim() && !val.trim().match(/^\d+$/)) {
    return val.trim();
  }
  const num = Number(val);
  if (!num || isNaN(num)) return String(val || "").trim();
  // Excel serial: days since Jan 1 1900 (with leap year bug)
  try {
    const date = xlsx.SSF.parse_date_code(num);
    if (date) {
      const y = date.y;
      const m = String(date.m).padStart(2, "0");
      const d = String(date.d).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
  } catch(e) {}
  return String(val).trim();
}

// ── Parse a monthly sheet ─────────────────────────────────────────────────────
function parseMonthSheet(ws) {
  if (!ws["!ref"]) return [];

  // Read raw=true so PO numbers don't become scientific notation
  const raw = xlsx.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: true,       // ← key fix: keeps 4000084421 as-is
  });

  if (raw.length < 2) return [];

  // Find header row — look for "PR NO." or "PR DATE"
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const row = raw[i];
    if (!row) continue;
    const rowStr = row.map(c => String(c || "").toUpperCase()).join("|");
    if (rowStr.includes("PR DATE") && rowStr.includes("PR NO")) {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx === -1) {
    console.log("    ⚠️  Could not find header row");
    return [];
  }

  // Build header map — normalize header names
  const headers = raw[headerRowIdx].map(h => String(h || "").trim().toUpperCase());

  // Flexible column finder — handles "PO Number", "PO NUMBER", "PO No.", etc.
  function findCol(patterns) {
    for (const pattern of patterns) {
      const idx = headers.findIndex(h => h.includes(pattern.toUpperCase()));
      if (idx !== -1) return idx;
    }
    return -1;
  }

  // Map all columns flexibly
  const colMap = {
    pr_date:            findCol(["PR DATE"]),
    pr_date_received:   findCol(["PR DATE RECEIVED", "DATE RECEIVED"]),
    pr_no:              findCol(["PR NO"]),
    requesting_dept:    findCol(["REQUESTING DEPT", "DEPT"]),
    po_date:            findCol(["PO DATE"]),
    po_number:          findCol(["PO NUMBER", "PO NO", "PO NUM"]),
    end_user:           findCol(["END USER"]),
    supplier_name:      findCol(["SUPPLIER"]),
    item_code:          findCol(["ITEM CODE"]),
    item_description:   findCol(["ITEM DESCRIPTION", "DESCRIPTION"]),
    specifications:     findCol(["SPECIFICATION"]),
    qty:                findCol(["QTY", "QUANTITY"]),
    uom:                findCol(["UOM", "UoM", "UNIT OF MEASURE"]),
    unit_price:         findCol(["UNIT PRICE"]),
    amount:             findCol(["AMOUNT"]),
    total_amount:       findCol(["TOTAL AMOUNT"]),
    payment_terms:      findCol(["PAYMENT TERMS", "PAYMENT"]),
    pr_required_date:   findCol(["PR REQUIRED DATE", "REQUIRED DATE"]),
    date_delivered:     findCol(["DATE DELIVERED", "DELIVERED"]),
    remarks:            findCol(["REMARKS"]),
    po_status:          findCol(["PURCHASE ORDER STATUS", "PO STATUS", "STATUS"]),
    items_services:     findCol(["ITEMS/SERVICES", "ITEMS", "SERVICES"]),
    original_price:     findCol(["ORIGINAL PRICE"]),
    total_cost_savings: findCol(["TOTAL COST SAVINGS", "COST SAVINGS"]),
  };

  console.log(`    Header row: ${headerRowIdx}, PO Number col: ${colMap.po_number}, PR No col: ${colMap.pr_no}`);

  const rows = [];

  for (let i = headerRowIdx + 1; i < raw.length; i++) {
    const r = raw[i];
    if (!r) continue;

    // Get value from a column index
    const get = (idx) => {
      if (idx === -1) return "";
      const val = r[idx];
      if (val === null || val === undefined) return "";
      const s = String(val).trim();
      if (s.startsWith("#")) return "";  // Excel error cells
      if (s === " " || s === "") return "";
      return s;
    };

    // Get date value (converts serial numbers)
    const getDate = (idx) => {
      if (idx === -1) return "";
      const val = r[idx];
      if (val === null || val === undefined) return "";
      return excelDateToString(val);
    };

    // Get number value
    const getNum = (idx) => {
      if (idx === -1) return 0;
      const val = r[idx];
      if (!val) return 0;
      const n = parseFloat(String(val).replace(/[^0-9.-]/g, ""));
      return isNaN(n) ? 0 : n;
    };

    // PO Number — keep as full integer string, no scientific notation
    const poRaw = colMap.po_number !== -1 ? r[colMap.po_number] : null;
    const poNumber = poRaw !== null && poRaw !== undefined
      ? String(Math.round(Number(poRaw)))  // converts 4000084421.0 → "4000084421"
      : "";

    // PR No
    const prNo = get(colMap.pr_no);

    // Skip completely empty rows
    const hasContent = poNumber || prNo || get(colMap.item_description) || get(colMap.supplier_name);
    if (!hasContent) continue;

    // Skip summary/header-repeat rows
    const firstCell = String(r[0] || "").toUpperCase();
    if (firstCell.includes("PR DATE") || firstCell.includes("TOTAL") || firstCell.includes("GRAND")) continue;

    rows.push({
      "PR DATE":               getDate(colMap.pr_date),
      "PR DATE RECEIVED":      getDate(colMap.pr_date_received),
      "PR NO.":                prNo,
      "REQUESTING DEPT.":      get(colMap.requesting_dept),
      "PO DATE":               getDate(colMap.po_date),
      "PO NUMBER":             poNumber,
      "END USER/S":            get(colMap.end_user),
      "SUPPLIER'S NAME":       get(colMap.supplier_name),
      "ITEM CODE":             get(colMap.item_code),
      "ITEM DESCRIPTION":      get(colMap.item_description),
      "SPECIFICATIONS":        get(colMap.specifications),
      "QTY":                   get(colMap.qty),
      "UoM":                   get(colMap.uom),
      "UNIT PRICE":            getNum(colMap.unit_price),
      "AMOUNT":                getNum(colMap.amount),
      "TOTAL AMOUNT":          getNum(colMap.total_amount),
      "PAYMENT TERMS":         get(colMap.payment_terms),
      "PR REQUIRED DATE":      getDate(colMap.pr_required_date),
      "DATE DELIVERED":        getDate(colMap.date_delivered),
      "REMARKS":               get(colMap.remarks),
      "PURCHASE ORDER STATUS": get(colMap.po_status),
      "ITEMS/SERVICES":        get(colMap.items_services),
      "ORIGINAL PRICE":        0,   // not imported — savings derived from negative items
      "TOTAL COST SAVINGS":    0,   // not imported — calculated from discounts in system
    });
  }

  return rows;
}

// ── Main ──────────────────────────────────────────────────────────────────────
function main() {
  console.log("\n  Purchasing Tool -- Excel to SQLite Migration");
  console.log("-".repeat(50));

  if (!fs.existsSync(EXCEL_PATH)) {
    console.error("  Excel file not found: " + EXCEL_PATH);
    process.exit(1);
  }

  if (DRY_RUN) console.log("  DRY RUN -- nothing will be written\n");

  if (CLEAR_FIRST && !DRY_RUN) {
    db.prepare("DELETE FROM purchase_orders").run();
    try { db.prepare("DELETE FROM sqlite_sequence WHERE name='purchase_orders'").run(); } catch(e) {}
    console.log("  Cleared existing database rows\n");
  }

  console.log("  Reading: " + EXCEL_PATH);
  const wb = xlsx.readFile(EXCEL_PATH, { cellDates: false });

  const monthSheets = wb.SheetNames.filter(isMonthSheet);
  console.log("  Found " + monthSheets.length + " monthly sheet(s): " + monthSheets.join(", ") + "\n");

  let totalInserted = 0;
  let totalSkipped  = 0;
  const toInsert    = [];

  for (const sheetName of monthSheets) {
    const month = toMonthKey(sheetName);
    const ws    = wb.Sheets[sheetName];
    const rows  = parseMonthSheet(ws);

    console.log("  " + sheetName.padEnd(20) + " -> " + month + "  (" + rows.length + " rows parsed)");

    for (const row of rows) {
      if (!CLEAR_FIRST && row["PO NUMBER"]) {
        const exists = db.prepare(
          "SELECT 1 FROM purchase_orders WHERE TRIM(po_number) = TRIM(?)"
        ).get(row["PO NUMBER"]);
        if (exists) { totalSkipped++; continue; }
      }
      toInsert.push({ row, month });
      totalInserted++;
    }
  }

  console.log("\n  Summary:");
  console.log("    Rows to insert : " + totalInserted);
  console.log("    Rows skipped   : " + totalSkipped + " (duplicates)");

  if (!DRY_RUN && toInsert.length > 0) {
    console.log("\n  Writing to database...");
    insertMany(toInsert);
    console.log("  Done! " + totalInserted + " rows imported.");
    const count = db.prepare("SELECT COUNT(*) AS cnt FROM purchase_orders").get();
    console.log("  Total rows in database: " + count.cnt);
  } else if (DRY_RUN) {
    console.log("\n  Dry run complete. Run without --dry-run to import.");
  }

  console.log("");
}

main();