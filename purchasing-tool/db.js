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

module.exports = { db, insertRow, insertMany, updateRow, updateRowFile,
                   deleteRow, getByMonth, getMonths, findByPoNumber, buildDashboard };