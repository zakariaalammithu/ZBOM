// Build the compact, pre-parsed search index for the online API from the
// Fabric Material master workbook.
//
// Why this exists: parsing "Fabric Material code-.xlsx" (10.3 MB, 103,710
// data rows) with the xlsx library takes ~12 s and ~800 MB RSS — far beyond
// a Vercel Serverless Function cold start (Hobby: 10 s, 1024 MB). The index
// stores the same rows as plain arrays (columns listed once), which loads
// in ~0.2 s. Re-run after the master workbook changes:
//
//   npm run build-index
//
// Source:  ../Fabric Material code-.xlsx  (sheet "Mastar file")
// Output:  ../data/materials-index.json   { generatedAt, sourceFile,
//            sourceRows, columns, rows: [...] }
// The index is read by api/materials.js (via vercel.json `includeFiles`).
// It is NOT copied into web/ and must not be served as a static file
// (see .vercelignore) — the API only ever returns <= 100 rows per call.
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx");

const SOURCE = path.join(__dirname, "..", "Fabric Material code-.xlsx");
const SHEET = "Mastar file";
const OUT = path.join(__dirname, "..", "data", "materials-index.json");

function main() {
  const started = Date.now();
  const workbook = XLSX.readFile(SOURCE);
  const sheet = workbook.Sheets[SHEET];
  if (!sheet) throw new Error(`Sheet ${JSON.stringify(SHEET)} not found in ${SOURCE}`);
  const objects = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const columns = objects.length ? Object.keys(objects[0]) : [];
  const rows = objects.map((row) => columns.map((column) => row[column] ?? ""));
  const index = {
    generatedAt: new Date().toISOString(),
    sourceFile: path.basename(SOURCE),
    sourceSheet: SHEET,
    sourceRows: objects.length,
    columns,
    rows,
  };
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(index));
  const bytes = fs.statSync(OUT).size;
  console.log(`rows: ${objects.length}`);
  console.log(`columns: ${columns.length} (${columns.join(" | ")})`);
  console.log(`index: ${OUT} (${(bytes / 1048576).toFixed(1)} MB) in ${((Date.now() - started) / 1000).toFixed(1)}s`);
}

main();
