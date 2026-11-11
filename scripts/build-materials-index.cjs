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
// Output:  ../api/materials-data/materials-index-1.json ... -4.json
//          { generatedAt, sourceFile, sourceSheet, sourceRows, columns,
//            rows, chunk, chunks }
// The index is read by api/materials.js (via vercel.json `includeFiles`).
// It lives under api/ on purpose: files under api/ are bundled into the
// function but are NOT served as public static files, so the master
// never reaches a browser directly — the API only ever returns
// <= 100 rows per call. Do NOT move the chunks to data/ (that folder
// is .vercelignore'd, so Vercel would not bundle them at all — live
// /api/materials then 500s with "ENOENT: scandir '/var/task/data'",
// which is exactly what happened on 2026-11-11).
//
// Why chunks, not one file: the single 17.1 MB index
// (data/materials-index.json) would not push to GitHub — the SSH proxy
// tunnel broke with "Broken pipe" on every attempt containing that file,
// while code-only pushes went through fine. So the index is split into
// 4 chunks of ~26k rows (~4.3 MB each), committed and pushed one at a
// time. api/materials.js loads every
// api/materials-data/materials-index-*.json chunk it finds, so a partial
// upload still works (and reports clearly if no chunk is present at all).
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx");

const SOURCE = path.join(__dirname, "..", "Fabric Material code-.xlsx");
const SHEET = "Mastar file";
const CHUNKS = 4;
const OUT_DIR = path.join(__dirname, "..", "api", "materials-data");

function main() {
  const started = Date.now();
  const workbook = XLSX.readFile(SOURCE);
  const sheet = workbook.Sheets[SHEET];
  if (!sheet) throw new Error(`Sheet ${JSON.stringify(SHEET)} not found in ${SOURCE}`);
  const objects = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const columns = objects.length ? Object.keys(objects[0]) : [];
  const rows = objects.map((row) => columns.map((column) => row[column] ?? ""));
  const generatedAt = new Date().toISOString();
  const chunkSize = Math.ceil(rows.length / CHUNKS);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  // Remove any stale chunk files (e.g. from a previous run with a
  // different chunk count) so the API never loads leftovers.
  for (const stale of fs.readdirSync(OUT_DIR)) {
    if (/^materials-index-\d+\.json$/.test(stale)) fs.unlinkSync(path.join(OUT_DIR, stale));
  }

  let totalBytes = 0;
  for (let i = 0; i < CHUNKS; i++) {
    const chunkRows = rows.slice(i * chunkSize, (i + 1) * chunkSize);
    const out = path.join(OUT_DIR, `materials-index-${i + 1}.json`);
    const chunk = {
      generatedAt,
      sourceFile: path.basename(SOURCE),
      sourceSheet: SHEET,
      sourceRows: objects.length,
      columns,
      rows: chunkRows,
      chunk: i + 1,
      chunks: CHUNKS,
    };
    fs.writeFileSync(out, JSON.stringify(chunk));
    const bytes = fs.statSync(out).size;
    totalBytes += bytes;
    console.log(`chunk ${i + 1}/${CHUNKS}: ${chunkRows.length} rows -> ${out} (${(bytes / 1048576).toFixed(1)} MB)`);
  }
  console.log(`rows: ${objects.length}`);
  console.log(`columns: ${columns.length} (${columns.join(" | ")})`);
  console.log(`index: ${CHUNKS} chunks, ${(totalBytes / 1048576).toFixed(1)} MB total, in ${((Date.now() - started) / 1000).toFixed(1)}s`);
}

main();
