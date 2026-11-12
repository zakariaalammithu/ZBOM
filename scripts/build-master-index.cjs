// Build the bundled master data for the free client-side upload modules
// (see ONLINE_PLAN.md). The raw seed workbooks must NOT be published on
// the site, so this script extracts only what the matching/generation
// logic needs into small JSON files under web/masters-data/:
//
//   web/masters-data/zkng-template.json
//     The ZKNG template ("Formet" sheet, rows 1-4 only): the header rows
//     and the template row 4 whose fixed columns (G..FF) the desktop
//     generator copies into every output row. Only the fields SheetJS
//     needs on write are kept per cell: t (type), v (value), f (formula),
//     z (number format). Styles (s) are dropped on purpose — the
//     community SheetJS build does not write them, so the desktop output
//     does not carry them either. Dates are stored as ISO strings and
//     rebuilt as Date objects in the browser.
//
//   web/masters-data/routing-lookup.json
//     The Routing master lookup (materialType+group+fabricType key ->
//     unit / work centre / control key / operation description / base
//     quantity), built with the exact rules of the desktop
//     erp:import-routing-material handler (rows from index 2 on, key
//     whitespace-stripped and lowercased, "coming soon" and duplicate
//     keys skipped).
//
// LSMW BOM and Characteristics need no master at all (the desktop
// handlers compute everything from the uploaded rows), so no JSON is
// produced for them.
//
// Usage: npm run build-masters   (or: node scripts/build-master-index.cjs)
// Re-run after the seed masters in desktop-client/seed-data/ change.
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx");

const REPO_ROOT = path.join(__dirname, "..");
const SEED_DIR = path.join(REPO_ROOT, "desktop-client", "seed-data");
const OUT_DIR = path.join(REPO_ROOT, "web", "masters-data");

const ZKNG_TEMPLATE_FILE = "ZKNG Material Upload File.xlsx";
const ROUTING_MASTER_FILE = "Routing Upload File Master.xlsx";

function cleanCell(cell) {
  if (!cell) return null;
  const out = {};
  if (cell.t !== undefined) out.t = cell.t;
  if (cell.f !== undefined) out.f = cell.f;
  if (cell.v !== undefined) {
    // Dates come back as Date objects with cellDates:true; store ISO.
    out.v = cell.v instanceof Date ? cell.v.toISOString() : cell.v;
  }
  if (cell.z !== undefined) out.z = cell.z;
  // A cell with neither a value nor a formula carries nothing.
  if (out.v === undefined && out.f === undefined) return null;
  return out;
}

function buildZkngTemplate() {
  const source = path.join(SEED_DIR, ZKNG_TEMPLATE_FILE);
  const wb = XLSX.readFile(source, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const cells = {};
  let kept = 0;
  // Only rows 1-4 matter: rows 1-3 are the headers, row 4 is the
  // template row copied into every generated row. Rows 5-300 of the
  // seed repeat row 4 and are never read by the generator.
  for (let r = 1; r <= 4; r += 1) {
    for (let c = 0; c < 162; c += 1) {
      const ref = XLSX.utils.encode_cell({ r: r - 1, c });
      const cleaned = cleanCell(ws[ref]);
      if (cleaned) {
        cells[ref] = cleaned;
        kept += 1;
      }
    }
  }
  const data = {
    source: ZKNG_TEMPLATE_FILE,
    sheetName,
    columns: 162,
    templateRow: 4,
    cells,
  };
  const outPath = path.join(OUT_DIR, "zkng-template.json");
  fs.writeFileSync(outPath, JSON.stringify(data));
  console.log(`zkng-template.json: ${kept} cells from ${sheetName} rows 1-4 -> ${outPath} (${fs.statSync(outPath).size} bytes)`);
  return data;
}

function buildRoutingLookup() {
  const source = path.join(SEED_DIR, ROUTING_MASTER_FILE);
  const wb = XLSX.readFile(source, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  const entries = {};
  let skipped = 0;
  for (const masterRow of matrix.slice(2)) {
    const materialType = String(masterRow[1] ?? "").trim();
    const group = String(masterRow[2] ?? "").trim();
    const fabricType = String(masterRow[3] ?? "").trim();
    const rawLink = `${materialType}${group}${fabricType}`;
    const link = rawLink.replace(/\s+/g, "").toLowerCase();
    if (!link || /^coming soon/i.test(link) || Object.prototype.hasOwnProperty.call(entries, link)) {
      skipped += 1;
      continue;
    }
    entries[link] = {
      unit: String(masterRow[4] ?? "").trim(),
      workCentre: String(masterRow[5] ?? "").trim(),
      controlKey: String(masterRow[6] ?? "").trim(),
      operationDescription: String(masterRow[7] ?? "").trim(),
      baseQuantity: String(masterRow[8] ?? "").trim(),
    };
  }
  const data = {
    source: ROUTING_MASTER_FILE,
    sheetName,
    entries,
    count: Object.keys(entries).length,
  };
  const outPath = path.join(OUT_DIR, "routing-lookup.json");
  fs.writeFileSync(outPath, JSON.stringify(data));
  console.log(`routing-lookup.json: ${data.count} entries (${skipped} skipped) from ${sheetName} -> ${outPath} (${fs.statSync(outPath).size} bytes)`);
  return data;
}

fs.mkdirSync(OUT_DIR, { recursive: true });
buildZkngTemplate();
buildRoutingLookup();
console.log("Done.");
