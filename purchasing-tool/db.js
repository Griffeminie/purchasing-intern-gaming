/**
 * db.js
 * Database setup and all query functions.
 * Place this file next to server.js inside purchasing-tool/
 *
 * SQLite file will be created at:  data/purchasing.db
 */

const Database = require("better-sqlite3");
const path     = require("path");
const fs       = require("fs");

const DB_PATH = path.join(__dirname, "data", "purchasing.db");
fs.mkdirSync(path.join(__dirname, "data"), { recursive: true });

const db = new Database(DB_PATH);

db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

// ── Schema ────────────────────────────────────────────────────────────────────
db.exec(`
  CREATE TABLE IF NOT EXISTS purchase_orders (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    month               TEXT    NOT NULL,
    pr_date             TEXT,
    pr_date_received    TEXT,
    pr_no               TEXT,
    requesting_dept     TEXT,
    po_date             TEXT,
    po_number           TEXT,
    end_user            TEXT,
    supplier_name       TEXT,
    item_code           TEXT,
    item_description    TEXT,
    specifications      TEXT,
    qty                 TEXT,
    uom                 TEXT,
    unit_price          REAL    DEFAULT 0,
    amount              REAL    DEFAULT 0,
    total_amount        REAL    DEFAULT 0,
    payment_terms       TEXT,
    pr_required_date    TEXT,
    date_delivered      TEXT,
    remarks             TEXT,
    po_status           TEXT,
    items_services      TEXT,
    original_price      REAL    DEFAULT 0,
    total_cost_savings  REAL    DEFAULT 0,
    po_file             TEXT    DEFAULT '',
    created_at          TEXT    DEFAULT (datetime('now', 'localtime'))
  );

  CREATE INDEX IF NOT EXISTS idx_po_number   ON purchase_orders (po_number);
  CREATE INDEX IF NOT EXISTS idx_month       ON purchase_orders (month);
  CREATE INDEX IF NOT EXISTS idx_supplier    ON purchase_orders (supplier_name);
  CREATE INDEX IF NOT EXISTS idx_dept        ON purchase_orders (requesting_dept);
  CREATE INDEX IF NOT EXISTS idx_po_status   ON purchase_orders (po_status);
`);

// ── Add po_file column to existing databases (safe to run every time) ─────────
try {
  db.exec(`ALTER TABLE purchase_orders ADD COLUMN po_file TEXT DEFAULT ''`);
  console.log("[db] Added po_file column to existing database.");
} catch(e) {
  // Column already exists — this is expected on fresh installs
}

// ── Suppliers schema ────────────────────────────────────────────────────────────
db.exec(`
  CREATE TABLE IF NOT EXISTS suppliers (
    id                INTEGER PRIMARY KEY AUTOINCREMENT
  );
`);

// Add any columns that don't exist yet — safe whether the table is brand new
// (created just above with only `id`) or already existed with a different
// shape from an earlier version of this app.
const SUPPLIER_COLUMNS = [
  ["no",              "TEXT"],
  ["company",         "TEXT"],
  ["category",        "TEXT"],
  ["category2",       "TEXT"],
  ["old_name",        "TEXT"],
  ["current_name",    "TEXT"],
  ["contact_name",    "TEXT"],
  ["address",         "TEXT"],
  ["region",          "TEXT"],
  ["contact_details", "TEXT"],
  ["email",           "TEXT"],
  ["email2",          "TEXT"],
  ["status",          "TEXT"],
  ["created_at",      "TEXT"],
];
for (const [col, type] of SUPPLIER_COLUMNS) {
  try {
    db.exec(`ALTER TABLE suppliers ADD COLUMN ${col} ${type}`);
    console.log(`[db] Added "${col}" column to suppliers table.`);
  } catch (e) {
    // Column already exists — expected on every run after the first
  }
}

// Indexes — wrapped individually so one failure can't block the others or crash startup
for (const [name, col] of [
  ["idx_supplier_company",  "company"],
  ["idx_supplier_region",   "region"],
  ["idx_supplier_category", "category"],
  ["idx_supplier_status",   "status"],
]) {
  try {
    db.exec(`CREATE INDEX IF NOT EXISTS ${name} ON suppliers (${col})`);
  } catch (e) {
    console.warn(`[db] Could not create index ${name}:`, e.message);
  }
}


// ── Helper: Excel row → DB row ────────────────────────────────────────────────
function excelRowToDb(row, month) {
  return {
    month,
    pr_date:            row["PR DATE"]              || "",
    pr_date_received:   row["PR DATE RECEIVED"]     || "",
    pr_no:              row["PR NO."]               || "",
    requesting_dept:    row["REQUESTING DEPT."]     || "",
    po_date:            row["PO DATE"]              || "",
    po_number:          row["PO NUMBER"]            || "",
    end_user:           row["END USER/S"]           || "",
    supplier_name:      row["SUPPLIER'S NAME"]      || "",
    item_code:          row["ITEM CODE"]            || "",
    item_description:   row["ITEM DESCRIPTION"]     || "",
    specifications:     row["SPECIFICATIONS"]       || "",
    qty:                row["QTY"]                  || "",
    uom:                row["UoM"] || row["UOM"]    || "",
    unit_price:         Number(row["UNIT PRICE"])   || 0,
    amount:             Number(row["AMOUNT"])        || 0,
    total_amount:       Number(row["TOTAL AMOUNT"]) || 0,
    payment_terms:      row["PAYMENT TERMS"]        || "",
    pr_required_date:   row["PR REQUIRED DATE"]     || "",
    date_delivered:     row["DATE DELIVERED"]       || "",
    remarks:            row["REMARKS"]              || "",
    po_status:          row["PURCHASE ORDER STATUS"] || "",
    items_services:     row["ITEMS/SERVICES"]       || "",
    original_price:     0,  // savings derived from negative unit price items
    total_cost_savings: 0,  // not stored — calculated live from discounts
    po_file:            row["PO FILE"] || "",
  };
}

// ── Helper: DB row → API object ───────────────────────────────────────────────
function dbRowToApi(row) {
  return {
    id:                       row.id,
    month:                    row.month,
    "PR DATE":                row.pr_date,
    "PR DATE RECEIVED":       row.pr_date_received,
    "PR NO.":                 row.pr_no,
    "REQUESTING DEPT.":       row.requesting_dept,
    "PO DATE":                row.po_date,
    "PO NUMBER":              row.po_number,
    "END USER/S":             row.end_user,
    "SUPPLIER'S NAME":        row.supplier_name,
    "ITEM CODE":              row.item_code,
    "ITEM DESCRIPTION":       row.item_description,
    "SPECIFICATIONS":         row.specifications,
    "QTY":                    row.qty,
    "UoM":                    row.uom,
    "UNIT PRICE":             row.unit_price,
    "AMOUNT":                 row.amount,
    "TOTAL AMOUNT":           row.total_amount,
    "PAYMENT TERMS":          row.payment_terms,
    "PR REQUIRED DATE":       row.pr_required_date,
    "DATE DELIVERED":         row.date_delivered,
    "REMARKS":                row.remarks,
    "PURCHASE ORDER STATUS":  row.po_status,
    "ITEMS/SERVICES":         row.items_services,
    "ORIGINAL PRICE":         row.original_price,
    "TOTAL COST SAVINGS":     row.total_cost_savings,
    "PO FILE":                row.po_file || "",
    "created_at":             row.created_at,
  };
}

// ── Helper: supplier API row → DB row ─────────────────────────────────────────
function supplierRowToDb(row, i) {
  return {
    no:              String(row["NO"] ?? (i != null ? i + 1 : "")),
    company:         row["COMPANY"]         || "",
    category:        row["CATEGORY"]        || "",
    category2:       row["CATEGORY2"]       || "",
    old_name:        row["OLD_NAME"]        || "",
    current_name:    row["CURRENT_NAME"]    || "",
    contact_name:    row["CONTACT_NAME"]    || "",
    address:         row["ADDRESS"]         || "",
    region:          row["REGION"]          || "",
    contact_details: row["CONTACT_DETAILS"] || "",
    email:           row["EMAIL"]           || "",
    email2:          row["EMAIL2"]          || "",
    status:          row["STATUS"]          || "",
  };
}

// ── Helper: merge two "contact details" strings without losing existing numbers ─
// Splits on common separators, dedupes by the digits-only form so "0917 123 4567"
// and "09171234567" count as the same number, keeps existing order, and appends
// any genuinely new numbers found in the incoming data.
function mergeContactNumbers(oldRaw, newRaw) {
  const parseNums = (raw) =>
    String(raw || "")
      .split(/[\s'\/,]+/)
      .map((n) => n.trim())
      .filter(Boolean);

  const oldNums = parseNums(oldRaw);
  const newNums = parseNums(newRaw);
  const seen = new Set(oldNums.map((n) => n.replace(/\D/g, "")));
  const merged = [...oldNums];

  for (const n of newNums) {
    const digits = n.replace(/\D/g, "");
    if (digits && !seen.has(digits)) {
      merged.push(n);
      seen.add(digits);
    }
  }
  return merged.join(" / ");
}

// ── Helper: supplier DB row → API object ──────────────────────────────────────
function dbRowToSupplierApi(row) {
  return {
    id:              row.id,
    NO:              row.no,
    COMPANY:         row.company,
    CATEGORY:        row.category,
    CATEGORY2:       row.category2,
    OLD_NAME:        row.old_name,
    CURRENT_NAME:    row.current_name,
    CONTACT_NAME:    row.contact_name,
    ADDRESS:         row.address,
    REGION:          row.region,
    CONTACT_DETAILS: row.contact_details,
    EMAIL:           row.email,
    EMAIL2:          row.email2,
    STATUS:          row.status,
    created_at:      row.created_at,
  };
}

// ── Prepared statements ───────────────────────────────────────────────────────
const stmts = {

  insert: db.prepare(`
    INSERT INTO purchase_orders (
      month, pr_date, pr_date_received, pr_no, requesting_dept,
      po_date, po_number, end_user, supplier_name, item_code,
      item_description, specifications, qty, uom, unit_price,
      amount, total_amount, payment_terms, pr_required_date,
      date_delivered, remarks, po_status, items_services,
      original_price, total_cost_savings, po_file
    ) VALUES (
      @month, @pr_date, @pr_date_received, @pr_no, @requesting_dept,
      @po_date, @po_number, @end_user, @supplier_name, @item_code,
      @item_description, @specifications, @qty, @uom, @unit_price,
      @amount, @total_amount, @payment_terms, @pr_required_date,
      @date_delivered, @remarks, @po_status, @items_services,
      @original_price, @total_cost_savings, @po_file
    )
  `),

  update: db.prepare(`
    UPDATE purchase_orders SET
      month = @month, pr_date = @pr_date,
      pr_date_received = @pr_date_received, pr_no = @pr_no,
      requesting_dept = @requesting_dept, po_date = @po_date,
      po_number = @po_number, end_user = @end_user,
      supplier_name = @supplier_name, item_code = @item_code,
      item_description = @item_description, specifications = @specifications,
      qty = @qty, uom = @uom, unit_price = @unit_price,
      amount = @amount, total_amount = @total_amount,
      payment_terms = @payment_terms, pr_required_date = @pr_required_date,
      date_delivered = @date_delivered, remarks = @remarks,
      po_status = @po_status, items_services = @items_services,
      original_price = @original_price, total_cost_savings = @total_cost_savings
    WHERE id = @id
  `),

  updateFile: db.prepare(`
    UPDATE purchase_orders SET po_file = ? WHERE id = ?
  `),

  delete: db.prepare(`DELETE FROM purchase_orders WHERE id = ?`),

  byMonth: db.prepare(`
    SELECT * FROM purchase_orders WHERE month = ? ORDER BY id ASC
  `),

  months: db.prepare(`
    SELECT DISTINCT month FROM purchase_orders ORDER BY
      CASE month
        WHEN 'JAN' THEN 1  WHEN 'FEB' THEN 2  WHEN 'MAR' THEN 3
        WHEN 'APR' THEN 4  WHEN 'MAY' THEN 5  WHEN 'JUN' THEN 6
        WHEN 'JUL' THEN 7  WHEN 'AUG' THEN 8  WHEN 'SEP' THEN 9
        WHEN 'OCT' THEN 10 WHEN 'NOV' THEN 11 WHEN 'DEC' THEN 12
        ELSE 13
      END
  `),

  byPoNumber: db.prepare(`
    SELECT * FROM purchase_orders
    WHERE UPPER(TRIM(po_number)) = UPPER(TRIM(?))
    LIMIT 1
  `),

  // All totals use qty * unit_price (calculated), not stored Excel values
  // Negative qty*unit_price = discount/savings items

  monthlySpending: db.prepare(`
    SELECT month,
           SUM(CASE WHEN CAST(qty AS REAL) * unit_price > 0
                    THEN CAST(qty AS REAL) * unit_price
                    ELSE 0 END) AS total,
           SUM(CASE WHEN CAST(qty AS REAL) * unit_price < 0
                    THEN ABS(CAST(qty AS REAL) * unit_price)
                    ELSE 0 END) AS savings
    FROM purchase_orders
    GROUP BY month
  `),

  supplierBreakdown: db.prepare(`
    SELECT supplier_name AS name,
           SUM(CASE WHEN CAST(qty AS REAL) * unit_price > 0
                    THEN CAST(qty AS REAL) * unit_price
                    ELSE 0 END) AS total
    FROM purchase_orders
    WHERE supplier_name != ''
    GROUP BY supplier_name
    ORDER BY total DESC
    LIMIT 10
  `),

  deptSpending: db.prepare(`
    SELECT requesting_dept AS name,
           SUM(CASE WHEN CAST(qty AS REAL) * unit_price > 0
                    THEN CAST(qty AS REAL) * unit_price
                    ELSE 0 END) AS total
    FROM purchase_orders
    WHERE requesting_dept != ''
    GROUP BY requesting_dept
    ORDER BY total DESC
  `),

  statusCounts: db.prepare(`
    SELECT po_status, COUNT(DISTINCT po_number) AS cnt
    FROM purchase_orders
    WHERE po_status != '' AND po_number != ''
    GROUP BY po_status
  `),

  savingsData: db.prepare(`
    SELECT supplier_name AS "SUPPLIER'S NAME",
           SUM(CASE WHEN CAST(qty AS REAL) * unit_price > 0
                    THEN CAST(qty AS REAL) * unit_price ELSE 0 END) AS "TOTAL AMOUNT",
           SUM(CASE WHEN CAST(qty AS REAL) * unit_price < 0
                    THEN ABS(CAST(qty AS REAL) * unit_price) ELSE 0 END) AS "TOTAL COST SAVINGS"
    FROM purchase_orders
    WHERE supplier_name != ''
    GROUP BY supplier_name
    ORDER BY "TOTAL COST SAVINGS" DESC
  `),

  rowCount: db.prepare(`SELECT COUNT(*) AS cnt FROM purchase_orders`),

  // ── Suppliers ──────────────────────────────────────────────────────────────
  supplierInsert: db.prepare(`
    INSERT INTO suppliers (
      no, company, category, category2, old_name, current_name,
      contact_name, address, region, contact_details, email, email2, status,
      created_at
    ) VALUES (
      @no, @company, @category, @category2, @old_name, @current_name,
      @contact_name, @address, @region, @contact_details, @email, @email2, @status,
      datetime('now', 'localtime')
    )
  `),

  supplierUpdate: db.prepare(`
    UPDATE suppliers SET
      no = @no, company = @company, category = @category, category2 = @category2,
      old_name = @old_name, current_name = @current_name, contact_name = @contact_name,
      address = @address, region = @region, contact_details = @contact_details,
      email = @email, email2 = @email2, status = @status
    WHERE id = @id
  `),

  supplierDelete: db.prepare(`DELETE FROM suppliers WHERE id = ?`),
  supplierDeleteAll: db.prepare(`DELETE FROM suppliers`),
  supplierAll: db.prepare(`SELECT * FROM suppliers ORDER BY CAST(no AS INTEGER) ASC, id ASC`),
  supplierCount: db.prepare(`SELECT COUNT(*) AS cnt FROM suppliers`),
  // Match key for merge/upsert — company name, case/whitespace-insensitive
  supplierFindByCompany: db.prepare(`
    SELECT * FROM suppliers WHERE UPPER(TRIM(company)) = ? LIMIT 1
  `),
};

// ── Exported query functions ──────────────────────────────────────────────────

function insertRow(apiRow, month) {
  const row = excelRowToDb(apiRow, month);
  const info = stmts.insert.run(row);
  return info.lastInsertRowid;
}

function insertMany(rows) {
  const run = db.transaction((list) => {
    for (const { row, month } of list) {
      stmts.insert.run(excelRowToDb(row, month));
    }
  });
  run(rows);
}

function updateRow(id, apiRow, month) {
  const row = { ...excelRowToDb(apiRow, month), id };
  stmts.update.run(row);
}

function updateRowFile(id, filename) {
  stmts.updateFile.run(filename, id);
}

function deleteRow(id) {
  stmts.delete.run(id);
}

function getByMonth(month) {
  return stmts.byMonth.all(month.toUpperCase()).map(dbRowToApi);
}

function getMonths() {
  return stmts.months.all().map(r => r.month);
}

function findByPoNumber(poNumber) {
  const row = stmts.byPoNumber.get(poNumber);
  return row ? dbRowToApi(row) : null;
}

function buildDashboard() {
  const MONTH_ORDER = ["JAN","FEB","MAR","APR","MAY","JUN",
                       "JUL","AUG","SEP","OCT","NOV","DEC"];

  const spendMap   = {};
  const savingsMap = {};
  MONTH_ORDER.forEach(m => { spendMap[m] = 0; savingsMap[m] = 0; });
  stmts.monthlySpending.all().forEach(r => {
    if (spendMap[r.month] !== undefined) {
      spendMap[r.month]   = r.total   || 0;
      savingsMap[r.month] = r.savings || 0;
    }
  });
  const monthlySpending = MONTH_ORDER.map(m => ({
    month:   m,
    total:   spendMap[m],
    savings: savingsMap[m],
  }));

  const normStatus = { Served: 0, "For Delivery": 0, Pending: 0, Cancelled: 0 };
  stmts.statusCounts.all().forEach(r => {
    const u = (r.po_status || "").toLowerCase();
    if      (u.includes("served") || u.includes("complete")) normStatus["Served"]       += r.cnt;
    else if (u.includes("deliver"))                          normStatus["For Delivery"] += r.cnt;
    else if (u.includes("cancel"))                           normStatus["Cancelled"]    += r.cnt;
    else                                                     normStatus["Pending"]      += r.cnt;
  });

  return {
    monthlySpending,
    supplierBreakdown: stmts.supplierBreakdown.all(),
    deptSpending:      stmts.deptSpending.all(),
    statusCounts:      normStatus,
    savingsData:       stmts.savingsData.all(),
    rawRows:           stmts.rowCount.get().cnt,
    sheetsDetected:    getMonths(),
    cachedAt:          Date.now(),
  };
}

// ── Supplier query functions ───────────────────────────────────────────────────

function getAllSuppliers() {
  return stmts.supplierAll.all().map(dbRowToSupplierApi);
}

// Bulk replace — used by the "Import JSON" flow and the migrate-suppliers.js script.
// Wipes the table and inserts the given list in one transaction (atomic: all or nothing).
function importSuppliers(list) {
  const run = db.transaction((rows) => {
    stmts.supplierDeleteAll.run();
    try { db.prepare("DELETE FROM sqlite_sequence WHERE name='suppliers'").run(); } catch(e) {}
    rows.forEach((row, i) => stmts.supplierInsert.run(supplierRowToDb(row, i)));
  });
  run(list);
  return list.length;
}

// Merge/upsert — used by the "Update List" button. Unlike importSuppliers()
// (which wipes and replaces the whole table), this preserves every existing
// row: it matches incoming records to existing ones by COMPANY name, adds
// any brand-new contact numbers onto the existing ones (never drops old
// numbers), updates other fields only when the incoming value is non-empty
// and actually different, and inserts genuinely new suppliers with the next
// available NO. Existing suppliers not present in the incoming list are left
// untouched (never deleted).
function upsertSuppliers(list) {
  let inserted = 0, updated = 0, unchanged = 0, skipped = 0;

  const run = db.transaction((rows) => {
    let nextNo = stmts.supplierAll.all().reduce((max, r) => {
      const n = parseInt(r.no, 10);
      return isNaN(n) ? max : Math.max(max, n);
    }, 0);

    const FIELDS = [
      "category", "category2", "old_name", "current_name", "contact_name",
      "address", "region", "email", "email2", "status",
    ];

    rows.forEach((row, i) => {
      const incoming = supplierRowToDb(row, i);
      const companyKey = (incoming.company || incoming.current_name || "").trim();
      if (!companyKey) { skipped++; return; }

      const existing = stmts.supplierFindByCompany.get(companyKey.toUpperCase());

      if (!existing) {
        nextNo += 1;
        incoming.no = incoming.no && incoming.no.trim() ? incoming.no : String(nextNo);
        stmts.supplierInsert.run(incoming);
        inserted++;
        return;
      }

      const merged = { ...existing };
      let changed = false;

      for (const f of FIELDS) {
        if (incoming[f] && incoming[f] !== existing[f]) {
          merged[f] = incoming[f];
          changed = true;
        }
      }

      const mergedContacts = mergeContactNumbers(existing.contact_details, incoming.contact_details);
      if (mergedContacts !== (existing.contact_details || "")) {
        merged.contact_details = mergedContacts;
        changed = true;
      }

      if (changed) {
        stmts.supplierUpdate.run({ ...merged, id: existing.id });
        updated++;
      } else {
        unchanged++;
      }
    });
  });

  run(list);
  return { inserted, updated, unchanged, skipped, total: list.length };
}

function insertSupplier(row) {
  const dbRow = supplierRowToDb(row);
  const info = stmts.supplierInsert.run(dbRow);
  return info.lastInsertRowid;
}

function updateSupplier(id, row) {
  const dbRow = { ...supplierRowToDb(row), id };
  stmts.supplierUpdate.run(dbRow);
}

function deleteSupplier(id) {
  stmts.supplierDelete.run(id);
}

function supplierCount() {
  return stmts.supplierCount.get().cnt;
}

module.exports = { db, insertRow, insertMany, updateRow, updateRowFile,
                   deleteRow, getByMonth, getMonths, findByPoNumber, buildDashboard,
                   getAllSuppliers, importSuppliers, upsertSuppliers, insertSupplier,
                   updateSupplier, deleteSupplier, supplierCount };