// GET /api/materials?q=<text>&limit=<n> — server-side search over the
// Fabric Material master ("Fabric Material code-.xlsx", sheet "Mastar
// file": 103,710 rows x 17 columns).
//
// Case-insensitive substring match across all columns of a row, same
// matching rule as the desktop renderer filterRows()/rowMatchesQuery().
// Returns at most `limit` rows (default 50, max 100) plus the total match
// count — the full 103,710-row master is NEVER sent to the browser.
//
// Performance: parsing the 10.3 MB .xlsx with the xlsx library takes
// ~12 s and ~800 MB RSS, which exceeds a Vercel Hobby cold start
// (10 s / 1024 MB). So this function reads the pre-extracted compact
// index in data/materials-index-*.json instead (4 chunks of ~26k rows,
// ~4.3 MB each, columns listed once, rows as arrays — see
// scripts/build-materials-index.cjs). First load parses them in ~0.2 s;
// the parsed index is cached in module scope, so warm invocations only
// pay the ~0.1-0.3 s search. Re-run `npm run build-index` after the
// master workbook changes.
//
// Why chunks: the single 17.1 MB index file would not push to GitHub —
// the SSH proxy tunnel broke with "Broken pipe" on every attempt
// containing it. The loader below reads EVERY chunk file it finds, so a
// partial upload still serves the rows it has, and it fails with a clear
// 500 ("materials index is not fully uploaded yet") if no chunk is
// present at all, instead of crashing.
//
// Plain Node req/res only — no framework, works on Vercel and locally.
// The index file reaches Vercel via vercel.json `functions.includeFiles`
// (it is excluded from static deployment by .vercelignore).
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const INDEX_DIR = path.join(__dirname, "..", "data");
const CHUNK_FILE = /^materials-index-(\d+)\.json$/;
const DEFAULT_LIMIT = 50;
const MAX_LIMIT = 100;

let cache = null; // { columns, rows, searchText, generatedAt, sourceRows, chunksLoaded }

function loadIndex() {
  if (cache) return cache;
  let files = [];
  try {
    files = fs
      .readdirSync(INDEX_DIR)
      .filter((name) => CHUNK_FILE.test(name))
      .sort((a, b) => Number(a.match(CHUNK_FILE)[1]) - Number(b.match(CHUNK_FILE)[1]));
  } catch (error) {
    throw new Error(`materials index is not fully uploaded yet (could not read index directory: ${error.message})`);
  }
  if (files.length === 0) {
    throw new Error("materials index is not fully uploaded yet (no data/materials-index-*.json chunk found)");
  }

  // Load every chunk present, in chunk-number order, and concatenate rows.
  // A partial upload still serves the rows it has; a chunk that fails to
  // parse is skipped, not fatal — but zero rows overall is an error.
  let columns = null;
  let generatedAt = null;
  let sourceRows = 0;
  const rows = [];
  let loaded = 0;
  for (const file of files) {
    let parsed;
    try {
      parsed = JSON.parse(fs.readFileSync(path.join(INDEX_DIR, file), "utf8"));
    } catch (error) {
      console.error(`materials index: skipping unreadable chunk ${file}: ${error.message}`);
      continue;
    }
    if (!Array.isArray(parsed.rows)) {
      console.error(`materials index: skipping chunk ${file}: no rows array`);
      continue;
    }
    if (!columns) {
      columns = parsed.columns;
      generatedAt = parsed.generatedAt;
    }
    sourceRows = parsed.sourceRows || sourceRows;
    rows.push(...parsed.rows);
    loaded += 1;
  }
  if (rows.length === 0) {
    throw new Error(`materials index is not fully uploaded yet (found ${files.length} chunk file(s), loaded 0 rows)`);
  }
  if (loaded < files.length) {
    console.error(`materials index: loaded ${loaded} of ${files.length} chunk file(s)`);
  }

  cache = {
    columns,
    rows,
    generatedAt,
    sourceRows,
    chunksLoaded: loaded,
    // One lowercase searchable string per row, built once.
    searchText: rows.map((row) =>
      row.map((value) => String(value ?? "")).join(" ").toLowerCase(),
    ),
  };
  return cache;
}

function toObject(columns, row) {
  const object = {};
  columns.forEach((column, i) => {
    object[column] = row[i] ?? "";
  });
  return object;
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(body),
  });
  res.end(body);
}

module.exports = (req, res) => {
  if (req.method !== "GET") {
    res.writeHead(405, { Allow: "GET" });
    res.end("Method Not Allowed");
    return;
  }
  let index;
  try {
    index = loadIndex();
  } catch (error) {
    const message = String(error && error.message ? error.message : error);
    // The "not fully uploaded yet" case gets its own clear error message;
    // anything else keeps the generic one with the detail attached.
    if (message.startsWith("materials index is not fully uploaded yet")) {
      sendJson(res, 500, { error: message });
    } else {
      sendJson(res, 500, {
        error: "materials index could not be loaded",
        detail: message,
      });
    }
    return;
  }

  const url = new URL(req.url, "http://localhost");
  const q = (url.searchParams.get("q") || "").trim().toLowerCase();
  let limit = parseInt(url.searchParams.get("limit") || String(DEFAULT_LIMIT), 10);
  if (!Number.isFinite(limit) || limit < 1) limit = DEFAULT_LIMIT;
  if (limit > MAX_LIMIT) limit = MAX_LIMIT;

  const { columns, rows, searchText } = index;
  const matches = [];
  if (!q) {
    for (let i = 0; i < rows.length && matches.length < limit; i++) matches.push(i);
    sendJson(res, 200, {
      q: "",
      total: rows.length,
      returned: matches.length,
      limit,
      columns,
      source: { file: "Fabric Material code-.xlsx", sheet: "Mastar file", indexGeneratedAt: index.generatedAt },
      rows: matches.map((i) => toObject(columns, rows[i])),
    });
    return;
  }

  let total = 0;
  for (let i = 0; i < rows.length; i++) {
    if (searchText[i].includes(q)) {
      total += 1;
      if (matches.length < limit) matches.push(i);
    }
  }
  sendJson(res, 200, {
    q: url.searchParams.get("q") || "",
    total,
    returned: matches.length,
    limit,
    columns,
    source: { file: "Fabric Material code-.xlsx", sheet: "Mastar file", indexGeneratedAt: index.generatedAt },
    rows: matches.map((i) => toObject(columns, rows[i])),
  });
};
