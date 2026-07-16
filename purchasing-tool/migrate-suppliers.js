/**
 * migrate-suppliers.js
 * Imports a supplier JSON file (the output of xlsx_to_json.py, or any file
 * shaped like { suppliers: [...] } / a raw array) straight into SQLite.
 *
 * Usage:
 *   node migrate-suppliers.js                          # uses data/supplier_data.json
 *   node migrate-suppliers.js path/to/other_file.json
 *   node migrate-suppliers.js --dry-run                # parse & report only, no writes
 */

const path = require("path");
const fs   = require("fs");
const { importSuppliers, supplierCount } = require("./db");

const args     = process.argv.slice(2);
const DRY_RUN  = args.includes("--dry-run");
const fileArg  = args.find(a => !a.startsWith("--"));
const JSON_PATH = fileArg
  ? path.resolve(fileArg)
  : path.join(__dirname, "data", "supplier_data.json");

function main() {
  console.log("\n  Purchasing Tool -- Supplier JSON to SQLite Migration");
  console.log("-".repeat(50));

  if (!fs.existsSync(JSON_PATH)) {
    console.error("  File not found: " + JSON_PATH);
    process.exit(1);
  }

  console.log("  Reading: " + JSON_PATH);
  const raw  = fs.readFileSync(JSON_PATH, "utf8");
  const data = JSON.parse(raw);
  const list = Array.isArray(data) ? data : data.suppliers;

  if (!Array.isArray(list)) {
    console.error('  Expected a "suppliers" array (or a top-level array) of records.');
    process.exit(1);
  }

  console.log("  Records found: " + list.length);
  if (data.source_file)  console.log("  Source file  : " + data.source_file);
  if (data.generated_at) console.log("  Generated at : " + data.generated_at);

  if (DRY_RUN) {
    console.log("\n  DRY RUN -- nothing was written. Sample record:");
    console.log(JSON.stringify(list[0], null, 2));
    return;
  }

  console.log("\n  Importing (this replaces the current supplier table)...");
  const count = importSuppliers(list);
  console.log("  Done! " + count + " suppliers imported.");
  console.log("  Total suppliers in database: " + supplierCount());
  console.log("");
}

main();
