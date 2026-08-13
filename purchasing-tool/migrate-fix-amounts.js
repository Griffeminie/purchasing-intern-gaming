/**
 * migrate-fix-amounts.js
 *
 * One-time backfill for rows saved before scanner_api.py was fixed to
 * compute AMOUNT/TOTAL AMOUNT itself. Those older scanned rows have Qty and
 * Unit Price stored correctly, but amount/total_amount were left at 0 in
 * the database — Monitoring's on-screen table masks this by recalculating
 * Qty × Unit Price live for display, but the Excel export reads the stored
 * amount/total_amount columns directly, so it still showed 0 for those rows.
 *
 * This sets amount = total_amount = Qty × Unit Price for every row where
 * that's not already the case, matching the convention used everywhere else
 * in the app (Monitoring's table, the PO Editor, the dashboard queries).
 *
 * Safe to re-run — it's idempotent (rows already correct are simply
 * re-written to the same value, since amount is always derived the same way
 * everywhere else in the app; there's no separate "manually overridden
 * amount" concept to preserve).
 *
 * Usage:
 *   node migrate-fix-amounts.js            (apply the fix)
 *   node migrate-fix-amounts.js --dry-run  (show what WOULD change, no writes)
 */

const { db } = require("./db");

const DRY_RUN = process.argv.includes("--dry-run");

function run() {
  const rows = db.prepare(`
    SELECT id, qty, unit_price, amount, total_amount
    FROM purchase_orders
  `).all();

  const toFix = [];
  for (const r of rows) {
    const qty       = parseFloat(r.qty)        || 0;
    const unitPrice = parseFloat(r.unit_price) || 0;
    const correct   = Math.round(qty * unitPrice * 100) / 100; // round to cents
    const current   = Math.round((r.amount || 0) * 100) / 100;
    if (current !== correct) {
      toFix.push({ id: r.id, from: r.amount, to: correct });
    }
  }

  console.log(`Checked ${rows.length} rows — ${toFix.length} need fixing.`);

  if (!toFix.length) {
    console.log("Nothing to do.");
    return;
  }

  if (DRY_RUN) {
    console.log("\n--dry-run — no changes written. Sample of what would change:");
    toFix.slice(0, 20).forEach(f => console.log(`  id=${f.id}: amount ${f.from} -> ${f.to}`));
    if (toFix.length > 20) console.log(`  ...and ${toFix.length - 20} more`);
    return;
  }

  const update = db.prepare(`
    UPDATE purchase_orders
    SET amount = @amount, total_amount = @amount
    WHERE id = @id
  `);
  const applyAll = db.transaction((list) => {
    for (const f of list) update.run({ id: f.id, amount: f.to });
  });
  applyAll(toFix);

  console.log(`Fixed ${toFix.length} rows. Re-run with --dry-run any time to verify nothing's left.`);
}

run();
