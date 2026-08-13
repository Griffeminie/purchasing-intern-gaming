/**
 * xlsx-template.js
 * Builds month / year monitoring exports from data/monitoring_template.xlsx,
 * preserving its title, formatting, and summary formulas — instead of the old
 * plain header+rows export.
 *
 * The template (as uploaded) had two issues that would break or corrupt every
 * export generated from it, so they're fixed here rather than copied forward:
 *   1. Cell A4 was `=[1]JAN2026!A4` — a formula linking to an *external*
 *      workbook that isn't shipped with this app. Re-saving a file with a
 *      dangling external link like that risks Excel showing a "we found a
 *      problem with this file" repair prompt. It's replaced with a plain
 *      date value (the 1st of the exported month) — the same thing the
 *      formula evaluated to.
 *   2. Cells B16:B17 were merged with nothing in them — but row 16/17 falls
 *      inside the data area (row 7 onward). Left merged, writing a value to
 *      row 17's column B (REQUESTING DEPT.) would silently fail. The merge
 *      is dropped before data is written.
 *
 * Everything else in the template — the title, the summary/count boxes in
 * columns C/F/N-P/U-X, fonts, fills, column widths — is carried over as-is.
 * Those summary formulas (COUNTIF/COUNTA/SUBTOTAL) recalculate automatically
 * when the file is opened in Excel; if a stat looks stale, pressing F9 forces
 * a recalculation.
 *
 * Per-PO layout: rows sharing the same PO NUMBER are grouped together and
 * merged into one cell spanning that PO's line items for most columns —
<<<<<<< HEAD
 * see MERGE_HEADERS below for the exact list. Amount rolls up as a sum;
 * Original Price / Total Cost Savings use manually-entered values if a PO
 * has any, otherwise they're derived from the sign of each line's Amount
 * (positive lines → Original Price, negative "discount" lines → Total Cost
 * Savings) — matching the dashboard's own totalSpend/totalSavings
 * convention, since scanner-generated POs record a discount as its own
 * negative-amount line item rather than filling those two fields directly.
 * Total Amount is a formula (Original Price minus Total Cost Savings). PR
 * Required Date shows the group's Date Delivered value. Everything not in
 * MERGE_HEADERS (item code/description, specs, qty, uom, unit price, PR
 * Date Received, PO Date) stays per line item, same as a normal export.
=======
 * see MERGE_HEADERS below for the exact list. Amount stays per-line-item
 * (it's just Qty × Unit Price, already computed per row — merging/summing
 * it would hide each item's own figure). Original Price / Total Cost
 * Savings use manually-entered values if a PO has any, otherwise they're
 * derived from the sign of each line's Amount (positive lines → Original
 * Price, negative "discount" lines → Total Cost Savings) — matching the
 * dashboard's own totalSpend/totalSavings convention, since scanner-
 * generated POs record a discount as its own negative-amount line item
 * rather than filling those two fields directly. Total Amount is written
 * as a plain computed value (Original Price minus Total Cost Savings) —
 * not an Excel formula, since the number is already fully known at export
 * time. PR Required Date shows the group's own PR Required Date value.
 * Everything not in MERGE_HEADERS (item code/description, specs, qty, uom,
 * unit price, amount, PR Date Received, PO Date) stays per line item, same
 * as a normal export.
>>>>>>> c16602719a7c662c5a52c25710cffccb199f9ed6
 */

const path    = require("path");
const ExcelJS = require("exceljs");

const TEMPLATE_PATH = path.join(__dirname, "data", "monitoring_template.xlsx");

// Column A → X, in the exact order the template's row-6 headers use.
// (Matches server.js's HEADERS array; index i maps 1:1 to column i+1.)
const HEADERS = [
  "PR DATE", "PR DATE RECEIVED", "PR NO.", "REQUESTING DEPT.",
  "PO DATE", "PO NUMBER", "END USER/S", "SUPPLIER'S NAME",
  "ITEM CODE", "ITEM DESCRIPTION", "SPECIFICATIONS", "QTY", "UoM",
  "UNIT PRICE", "AMOUNT", "TOTAL AMOUNT", "PAYMENT TERMS",
  "PR REQUIRED DATE", "DATE DELIVERED", "REMARKS",
  "PURCHASE ORDER STATUS", "ITEMS/SERVICES", "ORIGINAL PRICE",
  "TOTAL COST SAVINGS",
];

const MONTH_NAMES = {
  JAN: "January", FEB: "February", MAR: "March",     APR: "April",
  MAY: "May",     JUN: "June",     JUL: "July",       AUG: "August",
  SEP: "September", OCT: "October", NOV: "November",  DEC: "December",
};

const DATA_START_ROW = 7; // row 6 is the header row in the template

const COL = {};
HEADERS.forEach((h, i) => { COL[h] = i + 1; }); // 1-based column index

// Columns merged into one cell per PO (spanning that PO's line items).
// Everything NOT in this list stays per-line-item: PR DATE RECEIVED, PO
<<<<<<< HEAD
// DATE, ITEM CODE, ITEM DESCRIPTION, SPECIFICATIONS, QTY, UoM, UNIT PRICE.
const MERGE_HEADERS = [
  "PR DATE", "PR NO.", "REQUESTING DEPT.", "PO NUMBER", "END USER/S",
  "SUPPLIER'S NAME", "AMOUNT", "TOTAL AMOUNT", "PAYMENT TERMS",
=======
// DATE, ITEM CODE, ITEM DESCRIPTION, SPECIFICATIONS, QTY, UoM, UNIT PRICE,
// AMOUNT (Qty × Unit Price is already a per-item figure — merging it would
// hide each item's own number behind a single summed cell).
const MERGE_HEADERS = [
  "PR DATE", "PR NO.", "REQUESTING DEPT.", "PO NUMBER", "END USER/S",
  "SUPPLIER'S NAME", "TOTAL AMOUNT", "PAYMENT TERMS",
>>>>>>> c16602719a7c662c5a52c25710cffccb199f9ed6
  "PR REQUIRED DATE", "DATE DELIVERED", "REMARKS",
  "PURCHASE ORDER STATUS", "ITEMS/SERVICES", "ORIGINAL PRICE",
  "TOTAL COST SAVINGS",
];

<<<<<<< HEAD
// Of the merged columns, AMOUNT rolls up as a plain SUM across the group's
// line items (each item has its own Amount — merging must total them, not
// just show the first item's number and silently drop the rest). Original
// Price and Total Cost Savings are handled separately below, since they
// need sign-aware splitting, not a plain sum.
const SUM_HEADERS = ["AMOUNT"];

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
=======
// (No columns need a plain-sum roll-up anymore now that AMOUNT stays
// per-line-item — Original Price / Total Cost Savings below use their own
// sign-aware split instead of a plain sum.)
const SUM_HEADERS = [];
>>>>>>> c16602719a7c662c5a52c25710cffccb199f9ed6

// Group rows by PO NUMBER, preserving first-seen order of each PO and the
// original relative order of rows within a PO. Blank/missing PO numbers are
// never grouped together — each such row is its own single-row group —
// since two unrelated blank rows aren't the same PO.
function groupRowsByPoNumber(rows) {
  const groups = [];
  const indexByPo = new Map();

  rows.forEach(row => {
    const po = String(row["PO NUMBER"] ?? "").trim();
    if (!po) { groups.push([row]); return; }

    if (indexByPo.has(po)) {
      groups[indexByPo.get(po)].push(row);
    } else {
      indexByPo.set(po, groups.length);
      groups.push([row]);
    }
  });

  return groups;
}

function sum(rows, key) {
  return rows.reduce((acc, r) => acc + (parseFloat(r[key]) || 0), 0);
}

// Sum of only the positive values of `key` across the group.
function positiveSum(rows, key) {
  return rows.reduce((acc, r) => {
    const v = parseFloat(r[key]) || 0;
    return acc + (v > 0 ? v : 0);
  }, 0);
}

// Sum of the absolute value of only the negative values of `key` across
// the group (i.e. how much was discounted, as a positive number).
function negativeSumAbs(rows, key) {
  return rows.reduce((acc, r) => {
    const v = parseFloat(r[key]) || 0;
    return acc + (v < 0 ? -v : 0);
  }, 0);
}

// ── Load the template once per export request ─────────────────────────────
async function loadTemplateWorkbook() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_PATH);
  const sheet = wb.worksheets[0];
  if (!sheet) throw new Error("Template has no worksheets: " + TEMPLATE_PATH);
  return sheet;
}

// ── Clone the template worksheet (values, styles, merges, widths, heights,
//    freeze panes) into a new sheet inside destWorkbook ────────────────────
function cloneTemplateSheet(templateSheet, destWorkbook, sheetName) {
  const dest = destWorkbook.addWorksheet(sheetName, {
    views: templateSheet.views,
  });

  templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const destRow = dest.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const destCell = destRow.getCell(colNumber);
      destCell.value = cell.value;
      destCell.style = cell.style; // font, fill, border, alignment, numFmt
    });
    if (row.height) destRow.height = row.height;
    destRow.commit();
  });

  templateSheet.columns.forEach((col, i) => {
    if (col && col.width) dest.getColumn(i + 1).width = col.width;
  });

  // Drop the stray B16:B17 merge (empty, and sits inside the data area) —
  // keep every other merge (e.g. the E2:Q2 title merge) as-is.
  (templateSheet.model.merges || [])
    .filter(m => m !== "B16:B17")
    .forEach(m => { try { dest.mergeCells(m); } catch (e) {} });

  return dest;
}

// ── Fill in the title/date and data rows for one month on a cloned sheet ──
//
// Rows are grouped by PO NUMBER first (Excel can only merge *contiguous*
// rows, so line items belonging to the same PO are written next to each
// other rather than in raw entry order).
//
// For a PO group, MERGE_HEADERS columns become one cell spanning the group:
//   - AMOUNT / ORIGINAL PRICE / TOTAL COST SAVINGS hold the SUM of that
//     PO's per-line values (each item has its own number; the merged cell
//     needs the total, not just the first item's figure).
//   - TOTAL AMOUNT holds a formula: Original Price (anchor) − Total Cost
//     Savings (anchor) — i.e. the pre-discount total minus the discount,
//     which recalculates automatically if either is edited later.
//   - PR REQUIRED DATE shows the group's Date Delivered value instead of
//     its own field, per your request (Date Delivered's own column is
//     merged separately, unchanged).
//   - Every other merged column (PR Date, PR No., Requesting Dept., PO
//     Number, End User/s, Supplier, Payment Terms, Remarks, PO Status,
//     Items/Services) repeats the group's first row value.
// Everything NOT in MERGE_HEADERS (Item Code, Item Description,
// Specifications, Qty, UoM, Unit Price, PR Date Received, PO Date) stays
// per-line-item, one value per row, same as a normal export.
//
// IMPORTANT ordering: cells are merged *before* their value is set. Setting
// a value first and merging afterward is what caused Amount/Total Amount
// to come out blank before — merging a range can clear values already
// sitting in it, so the merge always has to happen first.
function fillMonthSheet(sheet, monthAbbr, year, rows) {
  const monthName = MONTH_NAMES[monthAbbr] || monthAbbr;

  sheet.getCell("A3").value = `For the month of ${monthName} ${year}`;
  // Real date value in place of the template's broken external-workbook link
  sheet.getCell("A4").value = new Date(year, Object.keys(MONTH_NAMES).indexOf(monthAbbr), 1);

  const groups = groupRowsByPoNumber(rows);
  let r = DATA_START_ROW;

  groups.forEach(group => {
    const startRow = r;
    const endRow   = r + group.length - 1;

    // 1) Merge first.
    if (endRow > startRow) {
      MERGE_HEADERS.forEach(h => sheet.mergeCells(startRow, COL[h], endRow, COL[h]));
    }

    // 2) Per-line-item columns: every row keeps its own value.
    group.forEach((row, i) => {
      const rr = sheet.getRow(startRow + i);
      HEADERS.forEach(h => {
        if (MERGE_HEADERS.includes(h)) return; // set once at the anchor below
        rr.getCell(COL[h]).value = row[h] ?? "";
      });
      rr.commit();
    });

    // 3) Merged (PO-level) columns: set the anchor cell's value now that
    //    the range is merged.
    const anchor = group[0];
    MERGE_HEADERS.forEach(h => {
      if (["TOTAL AMOUNT", "ORIGINAL PRICE", "TOTAL COST SAVINGS"].includes(h)) return; // handled below
<<<<<<< HEAD
      let val;
      if (h === "PR REQUIRED DATE")   val = anchor["DATE DELIVERED"] ?? "";
      else if (SUM_HEADERS.includes(h)) val = sum(group, h);
      else                             val = anchor[h] ?? "";
      sheet.getCell(startRow, COL[h]).value = val;
=======
      sheet.getCell(startRow, COL[h]).value = anchor[h] ?? "";
>>>>>>> c16602719a7c662c5a52c25710cffccb199f9ed6
    });

    // Original Price / Total Cost Savings: prefer manually entered per-line
    // values if any exist in this group (someone filled them in by hand).
    // Otherwise derive them from the sign of each line's Amount — a
    // discount shows up as its own negative-amount line item (e.g. a
    // "Discount Amount" row), not through these two fields, for POs
    // generated by the scanner. Positive amounts sum into Original Price
    // (the pre-discount total); the absolute value of negative amounts
    // sums into Total Cost Savings (the discount) — same convention the
    // dashboard itself already uses (totalSpend / totalSavings).
    const manualOriginal = sum(group, "ORIGINAL PRICE");
    const manualSavings  = sum(group, "TOTAL COST SAVINGS");
    const originalPrice  = (manualOriginal !== 0 || manualSavings !== 0)
      ? manualOriginal
      : positiveSum(group, "AMOUNT");
    const totalSavings   = (manualOriginal !== 0 || manualSavings !== 0)
      ? manualSavings
      : negativeSumAbs(group, "AMOUNT");

    sheet.getCell(startRow, COL["ORIGINAL PRICE"]).value      = originalPrice;
    sheet.getCell(startRow, COL["TOTAL COST SAVINGS"]).value  = totalSavings;
<<<<<<< HEAD

    const origLetter = colLetter(COL["ORIGINAL PRICE"]);
    const savLetter  = colLetter(COL["TOTAL COST SAVINGS"]);
    sheet.getCell(startRow, COL["TOTAL AMOUNT"]).value = {
      formula: `${origLetter}${startRow}-${savLetter}${startRow}`,
    };
=======
    sheet.getCell(startRow, COL["TOTAL AMOUNT"]).value         = originalPrice - totalSavings;
>>>>>>> c16602719a7c662c5a52c25710cffccb199f9ed6

    r = endRow + 1;
  });
}

// ── Build a single-sheet workbook for one month ────────────────────────────
async function buildMonthWorkbook(monthAbbr, year, rows) {
  const templateSheet = await loadTemplateWorkbook();
  const wb    = new ExcelJS.Workbook();
  const sheet = cloneTemplateSheet(templateSheet, wb, `${monthAbbr}${year}`.substring(0, 31));
  fillMonthSheet(sheet, monthAbbr, year, rows);
  return wb;
}

// ── Build a multi-sheet workbook, one sheet per month that has data ───────
async function buildYearWorkbook(monthsWithRows, year) {
  const templateSheet = await loadTemplateWorkbook();
  const wb = new ExcelJS.Workbook();

  for (const { month, rows } of monthsWithRows) {
    const sheet = cloneTemplateSheet(templateSheet, wb, `${month}${year}`.substring(0, 31));
    fillMonthSheet(sheet, month, year, rows);
  }

  return wb;
}

module.exports = { buildMonthWorkbook, buildYearWorkbook, HEADERS };