// web-erp.js — ZBOM Phase 1 online bridge.
//
// Loaded in web/index.html BEFORE renderer.js. It defines window.erp in a
// plain browser, with the same method names as desktop-client/preload.cjs
// (which uses Electron ipcRenderer), but backed by fetch() to /api/*
// where a real online implementation exists, and by safe no-ops elsewhere.
//
// What is REAL in Phase 1:
//   - getState / getAccessState / update — guest state from /api/state,
//     plus the first 50 rows of the real Fabric Material master from
//     /api/materials (full server-side search is wired by web-online.js).
//   - adminLogin — REAL since Phase 2-lite (see ONLINE_PLAN.md): POST
//     /api/login checks the ONE admin account configured by Vercel env
//     vars (ZBOM_ADMIN_USERNAME / ZBOM_ADMIN_PASSWORD_HASH /
//     ZBOM_SESSION_SECRET). On success the 30-minute session token is
//     kept in sessionStorage and sent as `Authorization: Bearer <token>`;
//     getState/getAccessState then return the main-admin access object.
//   - getZkngState / getLsmwBomState / getRoutingState /
//     getCharacteristicsState — live in-memory session states for the
//     free upload modules (see below). masterReady, canManageMaster and
//     canImport are all true for everyone: no login is needed.
//   - Free upload modules — ZKNG / LSMW BOM / Routing / Characteristics
//     generation runs entirely in the browser, ported from the desktop
//     handlers in desktop-client/main.cjs (same matching rules, same
//     output shapes). The masters are bundled as JSON under
//     /web/masters-data/ (built by scripts/build-master-index.cjs); the
//     user's file is read with a file picker and never leaves their
//     computer; the ready file is generated in memory and downloaded
//     with a Blob. A page reload loses the ready file.
//   - exportRows — client-side CSV download (Blob). Excel (.xlsx) is NOT
//     done in the browser: no SheetJS is loaded; the button gets a clear
//     message instead. NOTE: renderer.js only calls exportRows after a
//     sign-in + export permission check, and sign-in is Phase 2, so in
//     Phase 1 the export buttons open the sign-in dialog. exportRows
//     itself is implemented and working when called.
//   - saveSettings — echoes the settings back (no server settings store yet).
//   - recordUsage / accountHeartbeat — { ok: true } (desktop stubs too).
//
// What is HONESTLY NOT built (full Phase 2 needs a database — see
// ONLINE_PLAN.md):
//   - registerAccount / changeOwnPassword / setAdminPassword — the online
//     admin password cannot be changed from inside the app: it lives in
//     the Vercel env var ZBOM_ADMIN_PASSWORD_HASH, so it is changed in
//     Vercel and redeployed. There are NO hardcoded passwords and NO
//     invented accounts anywhere in this file, and only the single env-var
//     admin account can sign in (no other users yet).
//   - All master uploads (uploadZkngMaster, uploadLsmwBomMaster,
//     uploadRoutingMaster, uploadCharacteristicsMaster, uploadMasterSection,
//     clearMasterSection) — the bundled seed masters are fixed: replacing
//     a master is still an admin job for later (Phase 2). The upload
//     buttons answer with a Phase 2 message.
//   - Department files, sub-admin and role management — { ok: false } with a Phase 2 message.
//   - Desktop-only actions (chooseFolder, openPath, chooseProfileImage,
//     reloadApp, software update, saved accounts on this computer) — safe
//     no-ops, they cannot exist in a browser.
//
// This file must never be copied into desktop-client/: the desktop app
// keeps using preload.cjs + Electron IPC.
"use strict";

(function () {
  var PHASE2_LOGIN =
    "অনলাইন লগইন Phase 2-এ আসছে (ডেটাবেস লাগবে)। আপাতত Guest হিসেবে দেখতে পারবেন।";
  var PHASE2_FEATURE =
    "এই কাজটি অনলাইনে এখনো আসেনি — Phase 2-এ আসছে (ডেটাবেস লাগবে)।";
  var DESKTOP_ONLY = "এই কাজটি শুধু ডেস্কটপ অ্যাপে হয়, অনলাইনে উপলব্ধ নয়।";
  var PASSWORD_CHANGE_MESSAGE =
    "অনলাইন অ্যাডমিন পাসওয়ার্ড অ্যাপ থেকে বদলানো যায় না — Vercel-এ ZBOM_ADMIN_PASSWORD_HASH বদলে আবার ডেপ্লয় করতে হবে।";
  var TOKEN_KEY = "zbom-online-token";

  function notAvailable(message) {
    return { ok: false, message: message || PHASE2_FEATURE };
  }

  // Session token storage. sessionStorage is per-tab and disappears when
  // the tab closes — the right scope for a 30-minute admin session.
  // Guarded: some sandboxes have no sessionStorage at all.
  function getToken() {
    try {
      return window.sessionStorage ? window.sessionStorage.getItem(TOKEN_KEY) : null;
    } catch (error) {
      return null;
    }
  }

  function setToken(token) {
    try {
      if (window.sessionStorage) window.sessionStorage.setItem(TOKEN_KEY, token);
    } catch (error) {
      // Storage unavailable — the login still succeeds, the session just
      // will not survive a reload.
    }
  }

  function clearToken() {
    try {
      if (window.sessionStorage) window.sessionStorage.removeItem(TOKEN_KEY);
    } catch (error) {
      // Nothing to clear.
    }
  }

  function authHeaders() {
    var headers = { Accept: "application/json" };
    var token = getToken();
    if (token) headers.Authorization = "Bearer " + token;
    return headers;
  }

  async function apiJson(path) {
    var res = await fetch(path, { headers: authHeaders() });
    if (res.status === 401) {
      // Session expired or revoked server-side: drop the stale token so
      // the next call goes out as guest instead of retrying it forever.
      clearToken();
      throw new Error("Request failed: 401 " + path);
    }
    if (!res.ok) throw new Error("Request failed: " + res.status + " " + path);
    return res.json();
  }

  // POST helper that returns the parsed body even on a 4xx, because the
  // login endpoint answers failures with a JSON { ok: false, message }.
  async function apiPostJson(path, payload) {
    var res = await fetch(path, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify(payload || {}),
    });
    var data = null;
    try {
      data = await res.json();
    } catch (error) {
      data = null;
    }
    if (!data || typeof data !== "object") {
      return { ok: false, message: "Sign in failed: " + res.status };
    }
    return data;
  }

  function guestAccess() {
    return {
      role: "guest",
      fullName: "Guest User",
      username: "",
      employeeId: "",
      departments: [],
      subAdmins: [],
      permissions: ["view"],
    };
  }

  function emptySnapshot() {
    return {
      materials: [],
      summary: [],
      diaWiseSummary: [],
      diaWiseTotals: {},
      finishType: [],
      sources: {
        materials: {
          filename: "Fabric Material code-.xlsx",
          sheet: "Mastar file",
          modifiedAt: null,
        },
      },
      updatedAt: null,
    };
  }

  // First window of the real master, for the initial table render.
  // Full search (any query, whole 103,710-row master) is done by
  // web-online.js through /api/materials.
  async function loadMaterialWindow() {
    try {
      var data = await apiJson("/api/materials?q=&limit=50");
      return Array.isArray(data.rows) ? data.rows : [];
    } catch (error) {
      return [];
    }
  }

  async function getState() {
    var meta = await apiJson("/api/state");
    var materials = await loadMaterialWindow();
    var snapshot = emptySnapshot();
    snapshot.materials = materials;
    snapshot.sources = (meta.snapshot && meta.snapshot.sources) || snapshot.sources;
    // Phase 2-lite: /api/state returns the main-admin access object when
    // the session token is valid; otherwise it returns guest. Trust the
    // server's answer — never upgrade the role in the browser.
    var authenticated = Boolean(meta.authenticated);
    return {
      config: meta.config || {},
      snapshot: snapshot,
      appVersion: meta.appVersion || "1.0.1",
      access: authenticated && meta.access ? meta.access : guestAccess(),
      departments: meta.departments || [],
      departmentStatuses: {},
      departmentFiles: [],
      usageAnalytics: [],
      authenticated: authenticated,
      profileImageDataUrl: "",
      codeRequirementPreview: { rows: [], file: "", sheetName: "" },
    };
  }

  async function getAccessState() {
    var meta = await apiJson("/api/state");
    var authenticated = Boolean(meta.authenticated);
    return {
      access: authenticated && meta.access ? meta.access : guestAccess(),
      departments: meta.departments || [],
      departmentStatuses: {},
      departmentFiles: [],
      usageAnalytics: [],
      authenticated: authenticated,
      profileImageDataUrl: "",
    };
  }

  // Mirrors desktop erp:update -> { ok, changed, message, snapshot }.
  async function update() {
    var state = await getState();
    return {
      ok: true,
      changed: true,
      message: "Online data refreshed from the Fabric Material master (first 50 rows shown — use search for the rest).",
      snapshot: state.snapshot,
    };
  }

  function csvCell(value) {
    var text = value === null || value === undefined ? "" : String(value);
    return /[",\n\r]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  }

  // Client-side CSV download. Called by renderer.js exportCurrent() only
  // after its sign-in/export-permission gate (Phase 2), but implemented
  // here so it works as soon as that gate can pass.
  async function exportRows(payload) {
    var rows = (payload && payload.rows) || [];
    var format = (payload && payload.format) || "csv";
    if (format !== "csv") {
      return {
        ok: false,
        message: "Excel (.xlsx) export অনলাইনে এখনো আসেনি — CSV ডাউনলোড করুন (Phase 2-এ আসছে)।",
      };
    }
    if (rows.length > 50000) {
      return {
        ok: false,
        limitExceeded: true,
        message: "Fabric Material Code exports are limited to 50,000 records at a time. Please make the search or filter more specific.",
      };
    }
    var columns = rows.length ? Object.keys(rows[0]) : [];
    var lines = [columns.map(csvCell).join(",")];
    rows.forEach(function (row) {
      lines.push(columns.map(function (column) { return csvCell(row[column]); }).join(","));
    });
    var name = String((payload && payload.defaultName) || "Zbom-Download").replace(/[<>:"/\\|?*\x00-\x1f]/g, "-");
    var filename = name + ".csv";
    var blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
    return { ok: true, path: filename };
  }

  // ------------------------------------------------------------------
  // Free upload modules — ZKNG / LSMW BOM / Routing / Characteristics
  // ------------------------------------------------------------------
  // Client-side only: no login, no server, no database. The four
  // generators are ported from the desktop handlers in
  // desktop-client/main.cjs (same matching rules, same output shapes).
  // The masters are bundled as JSON under /web/masters-data/ (built by
  // scripts/build-master-index.cjs from the seed workbooks — the raw
  // master .xlsx files are never published). The user's file is read
  // with a file picker and never leaves their computer. The ready file
  // lives only in memory; a page reload loses it (see ONLINE_PLAN.md).

  function requireSheetJS() {
    if (typeof XLSX === "undefined") {
      throw new Error("Spreadsheet engine could not be loaded. Please reload the page and try again.");
    }
    return XLSX;
  }

  function todayDateStr() {
    var now = new Date();
    var day = String(now.getDate()).padStart(2, "0");
    var month = String(now.getMonth() + 1).padStart(2, "0");
    return day + "." + month + "." + now.getFullYear();
  }

  function excelSerialToday() {
    var now = new Date();
    return (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  // Ported verbatim from desktop-client/main.cjs.
  function formatDDMMYYYY(value) {
    if (!value) return "";
    if (value instanceof Date) return String(value.getDate()).padStart(2, "0") + "." + String(value.getMonth() + 1).padStart(2, "0") + "." + value.getFullYear();
    var text = String(value).trim();
    if (!text) return "";
    var isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
    if (isoMatch) return isoMatch[3] + "." + isoMatch[2] + "." + isoMatch[1];
    var slashMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(text);
    if (slashMatch) return String(slashMatch[1]).padStart(2, "0") + "." + String(slashMatch[2]).padStart(2, "0") + "." + slashMatch[3];
    return text;
  }

  // Ported verbatim from desktop-client/main.cjs.
  function normalizedMaterialHeader(value) {
    return String(value || "")
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[\u200b-\u200f\u202a-\u202e\ufeff]/g, "")
      .toUpperCase()
      .replace(/[^A-Z0-9 ]/g, "");
  }

  // Ported verbatim from desktop-client/main.cjs.
  function sourceMaterialValue(row, headings) {
    if (!row) return "";
    var lookup = row._lookup;
    if (!lookup) {
      lookup = {};
      Object.keys(row).forEach(function (k) { lookup[normalizedMaterialHeader(k)] = row[k]; });
      Object.defineProperty(row, "_lookup", { value: lookup, enumerable: false });
    }
    var result = "";
    var found = false;
    (headings || []).some(function (heading) {
      var value = lookup[normalizedMaterialHeader(heading)];
      if (value !== undefined && value !== null && String(value).trim() !== "") {
        result = value;
        found = true;
        return true;
      }
      return false;
    });
    if (!found && headings && headings.length) {
      var first = lookup[normalizedMaterialHeader(headings[0])];
      if (first !== undefined && first !== null) return first;
    }
    return result;
  }

  var mastersCache = {};

  function fetchMasterJson(urlPath) {
    if (mastersCache[urlPath]) return Promise.resolve(mastersCache[urlPath]);
    return fetch(urlPath).then(function (res) {
      if (!res.ok) throw new Error("Master data could not be loaded (" + res.status + "). Please reload the page and try again.");
      return res.json();
    }).then(function (data) {
      mastersCache[urlPath] = data;
      return data;
    });
  }

  // Hidden file picker. Resolves null when the user dismisses it.
  function pickFile(accept) {
    return new Promise(function (resolve) {
      var input = document.createElement("input");
      input.type = "file";
      input.accept = accept || ".xlsx,.xls,.csv";
      input.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0;";
      var done = false;
      function finish(file) {
        if (done) return;
        done = true;
        window.removeEventListener("focus", onFocus);
        try { input.remove(); } catch (error) {}
        resolve(file || null);
      }
      function onFocus() {
        // Some browsers do not fire "cancel": if the picker closed with
        // no file chosen, treat the regained focus as a dismissal.
        window.setTimeout(function () {
          if (!done && !(input.files && input.files.length)) finish(null);
        }, 400);
      }
      input.addEventListener("change", function () {
        finish(input.files && input.files[0] ? input.files[0] : null);
      });
      input.addEventListener("cancel", function () { finish(null); });
      document.body.appendChild(input);
      window.addEventListener("focus", onFocus);
      input.click();
    });
  }

  // Reads the first sheet of an .xlsx/.xls/.csv file into the same row
  // objects the desktop readWorkbook() produces (string values,
  // dateNF "dd.mm.yyyy", empty cells as "").
  function readInputWorkbook(file) {
    var XLSXLib = requireSheetJS();
    function sheetRows(wb) {
      if (!wb || !wb.SheetNames || !wb.SheetNames.length) return [];
      var ws = wb.Sheets[wb.SheetNames[0]];
      if (!ws) return [];
      return XLSXLib.utils.sheet_to_json(ws, { defval: "", raw: false, dateNF: "dd.mm.yyyy" });
    }
    if (/\.csv$/i.test(file.name || "")) {
      return file.text().then(function (text) {
        return sheetRows(XLSXLib.read(text, { type: "string", cellDates: true }));
      });
    }
    return file.arrayBuffer().then(function (buf) {
      return sheetRows(XLSXLib.read(buf, { type: "array", cellDates: true }));
    });
  }

  function triggerDownload(blob, filename) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
    return filename;
  }

  function downloadXlsxWorkbook(wb, filename) {
    var XLSXLib = requireSheetJS();
    var bytes = XLSXLib.write(wb, { bookType: "xlsx", type: "array" });
    return triggerDownload(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
  }

  function downloadTextFile(text, filename) {
    return triggerDownload(new Blob([text], { type: "text/plain;charset=utf-8" }), filename);
  }

  // The material-code headings the desktop generators look for (the
  // fallback rule — a row counts if ANY of its cells is non-empty).
  var MATERIAL_CODE_HEADINGS = ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE", "CODE", "Code", "ITEM CODE", "Item Code", "MATERIAL"];

  function sourceRowsWithMaterial(materialRows) {
    var sourceRows = materialRows.filter(function (row) {
      return String(sourceMaterialValue(row, MATERIAL_CODE_HEADINGS)).trim();
    });
    if (!sourceRows.length) {
      sourceRows = materialRows.filter(function (row) {
        return Object.values(row || {}).some(function (v) { return v !== "" && v !== null && v !== undefined; });
      });
    }
    return sourceRows;
  }

  // ---- ZKNG --------------------------------------------------------
  // Port of generateZkngReadyWorkbook() in desktop-client/main.cjs.
  // The template (Formet sheet, rows 1-4) comes from
  // /web/masters-data/zkng-template.json: rows 1-3 are the headers,
  // row 4 is the template row whose fixed columns (G..FF) are copied
  // into every output row, then A-F / AI / BM / CN / CP / CR / DS are
  // overwritten from the source row. One output row per source row.
  // Rules (exactly as the desktop): B is always "ZKNG"; C is the
  // DESCRIPTION; D is the base unit (or ST for groups 160013-160020,
  // else KG); E is the group; F is the net weight or "0"; AI is KP10
  // (groups 160001-160012) or CC11 (160013-160020); BM is the standard
  // price (10 for ST, 315 for SD descriptions, 375 for YD, else blank).
  function generateZkngWorkbook(template, materialRows) {
    var XLSXLib = requireSheetJS();
    var sourceRows = sourceRowsWithMaterial(materialRows);
    if (!sourceRows.length) throw new Error("No data rows found in the selected Material Code file.");
    var ws = {};
    var templateRow4 = {};
    Object.keys(template.cells || {}).forEach(function (ref) {
      var cell = template.cells[ref];
      var rebuilt = {};
      if (cell.t !== undefined) rebuilt.t = cell.t;
      if (cell.f !== undefined) rebuilt.f = cell.f;
      if (cell.v !== undefined) rebuilt.v = cell.t === "d" ? new Date(cell.v) : cell.v;
      if (cell.z !== undefined) rebuilt.z = cell.z;
      var rowMatch = /^([A-Z]{1,3})([0-9]+)$/.exec(ref);
      var rowNum = rowMatch ? Number(rowMatch[2]) : 0;
      if (rowNum >= 1 && rowNum <= 3) ws[ref] = rebuilt;
      if (rowNum === 4) templateRow4[rowMatch[1]] = rebuilt;
    });
    var today = todayDateStr();
    var excelDate = excelSerialToday();
    sourceRows.forEach(function (sourceRow, rowOffset) {
      var r = 4 + rowOffset;
      Object.keys(templateRow4).forEach(function (colLetter) {
        ws[colLetter + r] = Object.assign({}, templateRow4[colLetter]);
      });
      var materialNumberRaw = sourceMaterialValue(sourceRow, ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"]);
      var description = sourceMaterialValue(sourceRow, ["DESCRIPTION", "Material Description", "MAKTX"]);
      var groupRaw = sourceMaterialValue(sourceRow, ["GROUP", "GRUIP", "Material Group", "MATKL"]);
      // Read but intentionally unused, like the desktop: B is always "ZKNG".
      sourceMaterialValue(sourceRow, ["MATERIAL TYPE", "Material Type", "MTART"]);
      var baseUnitRaw = sourceMaterialValue(sourceRow, ["BASE UNIT OF MEASURE", "Base Unit of Measure", "MEINS", "UOM", "UNIT"]);
      var netWeightRaw = sourceMaterialValue(sourceRow, ["NET WEIGHT", "Net Weight", "NTGEW", "WEIGHT", "Weight"]);
      var materialNumberText = String(materialNumberRaw).trim();
      var materialNumber = /^\d+$/.test(materialNumberText) && Number.isSafeInteger(Number(materialNumberText)) ? Number(materialNumberText) : materialNumberText;
      var groupText = String(groupRaw).trim();
      var group = /^\d+$/.test(groupText) ? Number(groupText) : groupText;
      var groupNumber = Number(groupText);
      var materialType = "ZKNG";
      var baseUnit = String(baseUnitRaw || "").trim() || (groupNumber >= 160013 && groupNumber <= 160020 ? "ST" : "KG");
      var netWeightStr = String(netWeightRaw || "").trim() || "0";
      var productionStorage = groupNumber >= 160001 && groupNumber <= 160012 ? "KP10" : groupNumber >= 160013 && groupNumber <= 160020 ? "CC11" : "KP10";
      var descriptionText = String(description || "");
      var standardPrice = baseUnit === "ST" ? 10 : /^SD/i.test(descriptionText) ? 315 : /^YD/i.test(descriptionText) ? 375 : "";
      ws["A" + r] = { t: typeof materialNumber === "number" ? "n" : "s", v: materialNumber };
      ws["B" + r] = { t: "s", v: materialType };
      ws["C" + r] = { t: "s", v: descriptionText };
      ws["D" + r] = { t: "s", v: baseUnit };
      ws["E" + r] = { t: typeof group === "number" ? "n" : "s", v: group };
      ws["F" + r] = { t: "s", v: netWeightStr };
      ws["AI" + r] = { t: "s", v: productionStorage };
      ws["BM" + r] = { t: "n", v: standardPrice };
      ws["CN" + r] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
      ws["CP" + r] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
      ws["CR" + r] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
      ws["DS" + r] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
    });
    var maxRow = 3 + sourceRows.length;
    ws["!ref"] = XLSXLib.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 161, r: maxRow - 1 } });
    ws["!pageSetup"] = { orientation: "landscape", paperSize: 9, fitToWidth: 1, fitToHeight: 0 };
    var wb = XLSXLib.utils.book_new();
    XLSXLib.utils.book_append_sheet(wb, ws, template.sheetName || "Formet");
    return { wb: wb, rowCount: sourceRows.length, sheetName: template.sheetName || "Formet", range: "A1:FF" + maxRow, date: today };
  }

  // ---- Routing -----------------------------------------------------
  // Port of erp:import-routing-material in desktop-client/main.cjs.
  // The lookup (materialType+group+fabricType -> unit / work centre /
  // control key / operation description / base quantity) comes from
  // /web/masters-data/routing-lookup.json, built with the exact desktop
  // rules ("coming soon" and duplicate keys skipped).
  var routingLookupMap = null;

  function getRoutingLookup(lookupData) {
    if (routingLookupMap) return routingLookupMap;
    routingLookupMap = new Map();
    Object.keys(lookupData.entries || {}).forEach(function (key) {
      routingLookupMap.set(key, lookupData.entries[key]);
    });
    return routingLookupMap;
  }

  function matchRoutingRows(lookup, sourceRows) {
    var routingRows = [];
    var unmatchedKeys = [];
    sourceRows.forEach(function (row) {
      var materialNumberText = String(sourceMaterialValue(row, ["MATERIAL NUMBER", "MATERIAL CODE", "Material Number", "Material Code", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"])).trim();
      if (!materialNumberText) return;
      var materialType = String(sourceMaterialValue(row, ["MATERIAL TYPE", "Material Type"])).trim();
      var group = String(sourceMaterialValue(row, ["GROUP", "MATERIAL GROUP", "Material Group", "MATKL"])).trim();
      var fabricType = String(sourceMaterialValue(row, ["FABRIC TYPE", "Fabric Type"])).trim();
      var link = (materialType + group + fabricType).replace(/\s+/g, "").toLowerCase();
      var matched = lookup.get(link);
      if (!matched) {
        unmatchedKeys.push({ materialNumber: materialNumberText, key: link || "Blank link" });
        return;
      }
      var materialNumber = /^\d+$/.test(materialNumberText) && Number.isSafeInteger(Number(materialNumberText)) ? Number(materialNumberText) : materialNumberText;
      var baseQuantity = /^[-+]?\d+(?:\.\d+)?$/.test(matched.baseQuantity) ? Number(matched.baseQuantity) : matched.baseQuantity;
      routingRows.push([
        materialNumber, 8110, "", matched.unit, "01.01.2017",
        "x" + String(routingRows.length + 1).padStart(7, "0"),
        matched.workCentre, matched.controlKey, matched.operationDescription,
        baseQuantity, "", "      ", 24, "H",
      ]);
    });
    var importedRows = sourceRows.filter(function (row) {
      return String(sourceMaterialValue(row, ["MATERIAL NUMBER", "MATERIAL CODE", "Material Number", "Material Code", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"])).trim();
    }).length;
    return { routingRows: routingRows, unmatchedKeys: unmatchedKeys, importedRows: importedRows };
  }

  function buildRoutingWorkbook(routingRows) {
    var XLSXLib = requireSheetJS();
    var headers = ["Material Number", "Plant", "Material Description", "Material Base Unit of Measure", "Valid from", "Operation Number", "Work Centre", "Control Key", "Oeration Description", "Base Quantity", "Time 1", "UoM of Time1", "Time 2", "UoM of Time2"];
    var sheet = XLSXLib.utils.aoa_to_sheet([headers].concat(routingRows));
    sheet["!cols"] = [10.21, 5.21, 8.93, 6.5, 9.93, 10.36, 11.79, 7.79, 13.93, 6.5, 3.93, 3.64, 6.5, 6.93].map(function (wch) { return { wch: wch }; });
    sheet["!autofilter"] = { ref: "A1:N1" };
    var workbook = XLSXLib.utils.book_new();
    XLSXLib.utils.book_append_sheet(workbook, sheet, "Routing Upload Format");
    return workbook;
  }

  // ---- LSMW BOM ----------------------------------------------------
  // Port of the inline generation in erp:import-lsmw-bom-material in
  // desktop-client/main.cjs. No master is needed: the base quantity is
  // computed from the input row's group / fabric type. Produces four
  // files — Header_BOM.xlsx/.txt and Item_BOM.xlsx/.txt.
  function buildLsmwBomFiles(matRows, dateStr) {
    var XLSXLib = requireSheetJS();
    var codeAliases = ["MATERIAL CODE", "Material Code", "Material Number", "MATERIAL NUMBER", "Material", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE", "CODE", "Code", "ITEM CODE", "Item Code", "MATERIAL"];
    var yarnAliases = ["YARN CODE", "Yarn Code", "BOM component", "BOM Component", "Yarn", "YARN", "YARN COUNT", "Yarn Count", "COMPONENT", "Component"];
    var qtyAliases = ["YARN %", "Yarn %", "Quantity", "QUANTITY", "QTY", "Qty", "PERCENT", "Percent", "%"];
    function getRowMaterialCode(r) {
      var found = String(sourceMaterialValue(r, codeAliases)).trim();
      if (found) return found;
      var values = Object.values(r || {});
      for (var i = 0; i < values.length; i++) {
        var text = String(values[i] || "").trim();
        if (text && text.toLowerCase() !== "null" && text.toLowerCase() !== "undefined") return text;
      }
      return "";
    }
    var uniqueHeaderMap = new Map();
    matRows.forEach(function (r) {
      var code = getRowMaterialCode(r);
      if (!code) return;
      if (!uniqueHeaderMap.has(code)) {
        var fabType = String(sourceMaterialValue(r, ["FABRIC TYPE", "Fabric Type"])).trim().toUpperCase();
        var materialGroup = Number(String(sourceMaterialValue(r, ["MATERIAL GROUP", "GROUP", "Material Group", "MATKL"])).trim());
        var baseQty = 99;
        if (materialGroup >= 160013 && materialGroup <= 160019) baseQty = 1000;
        else if (fabType === "YARN DYED") baseQty = 94;
        else if (materialGroup === 160005 || materialGroup === 160011) baseQty = 99;
        else baseQty = 99.15;
        uniqueHeaderMap.set(code, {
          "Material Number": isNaN(Number(code)) ? code : Number(code),
          Plant: 8110, "BOM Usage": 1, "BOM status": 1,
          "Base quantity": baseQty, "Valid-From Date": "01.01.2017",
        });
      }
    });
    var headerRows = Array.from(uniqueHeaderMap.values());
    var itemRows = matRows.map(function (r) {
      var code = getRowMaterialCode(r);
      var yarnCode = String(sourceMaterialValue(r, yarnAliases)).trim();
      var quantityText = String(sourceMaterialValue(r, qtyAliases)).trim();
      var parsedQuantity = Number(quantityText.replace(/,/g, ""));
      var quantity = quantityText === "" ? "" : Number.isFinite(parsedQuantity) ? parsedQuantity : quantityText;
      if (!code) return null;
      return {
        "Material Number": code, Plant: 8110,
        " BOM component": yarnCode,
        "BOM Item Text (Line 1)": "",
        " Quantity": quantity,
        " Unit": "KG",
        " Item category": "L",
        " Scrap Percent": "",
      };
    }).filter(Boolean);
    var headHeaders = ["Material Number", "Plant", "BOM Usage", "BOM status", "Base quantity", "Valid-From Date", "", "", ""];
    var headMatrix = [headHeaders].concat(headerRows.map(function (r) {
      return [r["Material Number"], r.Plant, r["BOM Usage"], r["BOM status"], r["Base quantity"], r["Valid-From Date"], "", "", ""];
    }));
    var headSheet = XLSXLib.utils.aoa_to_sheet(headMatrix);
    headSheet["!cols"] = [{ wch: 14 }, { wch: 9 }, { wch: 11 }, { wch: 11 }, { wch: 14 }, { wch: 15 }, { wch: 9 }, { wch: 9 }, { wch: 9 }];
    headSheet["!autofilter"] = { ref: "A1:F" + (headerRows.length + 1) };
    for (var row = 2; row <= headerRows.length + 1; row += 1) headSheet["E" + row].z = "0.000";
    var headWb = XLSXLib.utils.book_new();
    XLSXLib.utils.book_append_sheet(headWb, headSheet, "Sheet1");
    var headTxtLines = [headHeaders.slice(0, 6).join("\t") + "\t\t\t"];
    headerRows.forEach(function (r) {
      var bq = Number(r["Base quantity"]).toFixed(3);
      headTxtLines.push([r["Material Number"], r.Plant, r["BOM Usage"], r["BOM status"], bq, r["Valid-From Date"]].join("\t") + "\t\t\t");
    });
    var headTxt = headTxtLines.join("\r\n") + "\r\n";
    var itemHeaders = ["Material Number", "Plant", " BOM component", "BOM Item Text (Line 1)", " Quantity", " Unit", " Item category", " Scrap Percent"];
    var itemMatrix = [itemHeaders].concat(itemRows.map(function (r) {
      return [r["Material Number"], r.Plant, r[" BOM component"], "", r[" Quantity"], r[" Unit"], r[" Item category"], ""];
    }));
    var itemSheet = XLSXLib.utils.aoa_to_sheet(itemMatrix);
    itemSheet["!cols"] = [{ wch: 14 }, { wch: 13 }, { wch: 16 }, { wch: 19 }, { wch: 15 }, { wch: 11 }, { wch: 16 }, { wch: 14 }];
    itemSheet["!autofilter"] = { ref: "A1:H" + (itemRows.length + 1) };
    for (var row2 = 2; row2 <= itemRows.length + 1; row2 += 1) {
      if (itemSheet["E" + row2] && itemSheet["E" + row2].t === "n") itemSheet["E" + row2].z = "0.000";
    }
    var itemWb = XLSXLib.utils.book_new();
    XLSXLib.utils.book_append_sheet(itemWb, itemSheet, "Sheet1");
    var itemTxtLines = [itemHeaders.join("\t")];
    itemRows.forEach(function (r) {
      var quantity = typeof r[" Quantity"] === "number" ? r[" Quantity"].toFixed(3) : r[" Quantity"];
      itemTxtLines.push([r["Material Number"], r.Plant, r[" BOM component"], "", quantity, r[" Unit"], r[" Item category"], ""].join("\t"));
    });
    var itemTxt = itemTxtLines.join("\r\n") + "\r\n";
    return {
      headerRows: headerRows,
      itemRows: itemRows,
      files: [
        { name: "Header_BOM-" + dateStr + ".xlsx", kind: "xlsx", wb: headWb },
        { name: "Header_BOM-" + dateStr + ".txt", kind: "txt", text: headTxt },
        { name: "Item_BOM-" + dateStr + ".xlsx", kind: "xlsx", wb: itemWb },
        { name: "Item_BOM-" + dateStr + ".txt", kind: "txt", text: itemTxt },
      ],
    };
  }

  function downloadLsmwFile(fileEntry) {
    if (fileEntry.kind === "txt") return downloadTextFile(fileEntry.text, fileEntry.name);
    return downloadXlsxWorkbook(fileEntry.wb, fileEntry.name);
  }

  // ---- Characteristics ---------------------------------------------
  // Port of generateCharacteristicsReadyWorkbook() in
  // desktop-client/main.cjs. No master is needed: a pure mapping of the
  // Material Code columns onto the fixed 15 headers, Plant 8110.
  function buildCharacteristicsWorkbook(materialRows) {
    var XLSXLib = requireSheetJS();
    var sourceRows = sourceRowsWithMaterial(materialRows);
    if (!sourceRows.length) throw new Error("No data rows found in the selected Material Code file.");
    var headers = [
      "Material Number", "Plant", "Material Description", "Fabric Type",
      "Fabrication", "Design / Jacquard Design", "Fabric Composition",
      "GSM", "Finish Type", "Yarn Count & Process-1", "Fibre Type",
      "Special Treatment", "Color Code / Color", "Color Combo", "Repeat Size",
    ];
    var charRows = sourceRows.map(function (r) {
      var matNumText = String(sourceMaterialValue(r, ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"])).trim();
      var matNum = /^\d+$/.test(matNumText) && Number.isSafeInteger(Number(matNumText)) ? Number(matNumText) : matNumText;
      var gsmRaw = sourceMaterialValue(r, ["GSM", "Gsm"]);
      var gsmVal = /^\d+(?:\.\d+)?$/.test(gsmRaw) ? Number(gsmRaw) : gsmRaw;
      return [
        matNum,
        8110,
        String(sourceMaterialValue(r, ["DESCRIPTION", "Material Description", "MAKTX"])).trim(),
        String(sourceMaterialValue(r, ["FABRIC TYPE", "Fabric Type"])).trim(),
        String(sourceMaterialValue(r, ["FABRICATION", "Fabrication"])).trim(),
        String(sourceMaterialValue(r, ["DESIGN/JACQUARD DESIGN", "DESIGN", "JACQUARD DESIGN"])).trim(),
        String(sourceMaterialValue(r, ["FABRIC COMPOSITION", "Fabric Composition", "COMPOSITION"])).trim(),
        gsmVal,
        String(sourceMaterialValue(r, ["FINISH TYPE", "Finish Type"])).trim(),
        String(sourceMaterialValue(r, ["YARN COUNT & PROCESS-1", "YARN COUNT&PROCESS", "Yarn Count & Process-1"])).trim(),
        String(sourceMaterialValue(r, ["FIBRE TYPE", "Fibre Type"])).trim(),
        String(sourceMaterialValue(r, ["SPECIAL TREATMENT", "Special Treatment"])).trim(),
        String(sourceMaterialValue(r, ["COLOR/COLOR CODE", "COLOR CODE / COLOR", "COLOR CODE", "COLOR"])).trim(),
        String(sourceMaterialValue(r, ["COLOR COMBO", "Color Combo"])).trim(),
        String(sourceMaterialValue(r, ["REPEAT SIZE", "REPET SIZE", "Repeat Size"])).trim(),
      ];
    });
    var outSheet = XLSXLib.utils.aoa_to_sheet([headers].concat(charRows));
    outSheet["!autofilter"] = { ref: "A1:O" + (charRows.length + 1) };
    outSheet["!cols"] = [14, 8, 30, 16, 16, 22, 20, 8, 14, 22, 14, 18, 18, 16, 14].map(function (w) { return { wch: w }; });
    var outWb = XLSXLib.utils.book_new();
    XLSXLib.utils.book_append_sheet(outWb, outSheet, "Characteristics Upload Format");
    return { wb: outWb, rowCount: charRows.length };
  }

  // ---- In-memory sessions ------------------------------------------
  // One session per module: the state the UI reads, plus the generated
  // workbook(s) held in memory until downloaded or cleared. A page
  // reload drops them (nothing is stored on a server or in localStorage).
  function initialZkngState() {
    return {
      masterReady: true, zkngMasterReady: true,
      zkngMasterFile: "ZKNG Master.xlsx", zkngMasterUpdatedAt: null,
      canManageMaster: true, canImport: true,
      materialFile: null, generatedAt: null,
      importedRows: 0, matchedRows: 0, unmatchedRows: 0,
      materialReadyFile: null, materialReady: false, ready: false,
    };
  }
  function initialLsmwState() {
    return {
      masterReady: true, lsmwBomMasterReady: true,
      lsmwBomMasterFile: "LSMW BOM Upload Master File.xlsx", lsmwBomMasterUpdatedAt: null,
      canManageMaster: true, canImport: true,
      materialFile: null, generatedAt: null,
      importedRows: 0, matchedRows: 0, unmatchedRows: 0,
      headerRowsCount: 0, itemRowsCount: 0,
      generatedFiles: null, ready: false,
    };
  }
  function initialRoutingState() {
    return {
      masterReady: true, masterFile: "Built-in Routing Master", masterUpdatedAt: null,
      canManageMaster: true, canImport: true,
      materialFile: null, generatedAt: null,
      importedRows: 0, matchedRows: 0, unmatchedRows: 0, unmatchedKeys: [],
      readyFile: null, ready: false,
    };
  }
  function initialCharacteristicsState() {
    return {
      masterReady: true,
      canManageMaster: true, canImport: true,
      materialFile: null, generatedAt: null,
      importedRows: 0, matchedRows: 0, unmatchedRows: 0,
      materialReadyFile: null, materialReady: false, ready: false,
    };
  }

  var zkngSession = { state: initialZkngState(), workbook: null, filename: null };
  var lsmwSession = { state: initialLsmwState(), files: null };
  var routingSession = { state: initialRoutingState(), workbook: null, filename: null };
  var characteristicsSession = { state: initialCharacteristicsState(), workbook: null, filename: null };

  async function getZkngState() { return zkngSession.state; }
  async function getLsmwBomState() { return lsmwSession.state; }
  async function getRoutingState() { return routingSession.state; }
  async function getCharacteristicsState() { return characteristicsSession.state; }

  async function importZkngMaterial() {
    try {
      requireSheetJS();
      var file = await pickFile(".xlsx,.xls,.csv");
      if (!file) return { ok: false, canceled: true };
      var matRows = await readInputWorkbook(file);
      if (!matRows.length) return { ok: false, message: "No valid rows were found in the selected Material Code file." };
      var template = await fetchMasterJson("/web/masters-data/zkng-template.json");
      var dateStr = todayDateStr();
      var generated = generateZkngWorkbook(template, matRows);
      var readyName = "ZKNG Material Upload File - Ready-" + dateStr + ".xlsx";
      var nextState = Object.assign(initialZkngState(), {
        materialFile: file.name, generatedAt: new Date().toISOString(),
        importedRows: generated.rowCount, matchedRows: generated.rowCount, unmatchedRows: 0,
        materialReadyFile: readyName, materialReady: true, ready: true,
      });
      zkngSession.state = nextState;
      zkngSession.workbook = generated.wb;
      zkngSession.filename = readyName;
      return { ok: true, message: "Processing completed: " + generated.rowCount + " rows added to the ZKNG Material Upload Ready file.", state: nextState };
    } catch (error) {
      return { ok: false, message: error && error.message ? error.message : String(error) };
    }
  }

  async function downloadZkngReady() {
    if (!zkngSession.workbook || !zkngSession.state.materialReady) {
      return { ok: false, message: "No ready ZKNG Material Upload file available." };
    }
    var filename = downloadXlsxWorkbook(zkngSession.workbook, zkngSession.filename);
    return { ok: true, path: filename, message: "Downloaded " + filename + " to Downloads." };
  }

  async function importLsmwBomMaterial() {
    try {
      requireSheetJS();
      var file = await pickFile(".xlsx,.xls,.csv");
      if (!file) return { ok: false, canceled: true };
      var matRows = await readInputWorkbook(file);
      if (!matRows.length) return { ok: false, message: "No valid rows were found in the selected Material Code file." };
      var dateStr = todayDateStr();
      var generated = buildLsmwBomFiles(matRows, dateStr);
      var generatedFiles = generated.files.map(function (f) { return f.name; });
      var nextState = Object.assign(initialLsmwState(), {
        materialFile: file.name, generatedAt: new Date().toISOString(),
        importedRows: generated.itemRows.length, matchedRows: generated.itemRows.length, unmatchedRows: 0,
        headerRowsCount: generated.headerRows.length, itemRowsCount: generated.itemRows.length,
        generatedFiles: generatedFiles, ready: true,
      });
      lsmwSession.state = nextState;
      lsmwSession.files = generated.files;
      return { ok: true, message: "Processing completed: " + generated.itemRows.length + " rows added to the LSMW BOM Upload Ready files.", state: nextState };
    } catch (error) {
      return { ok: false, message: error && error.message ? error.message : String(error) };
    }
  }

  async function downloadLsmwBomReady(fileType) {
    var files = lsmwSession.files;
    if (!lsmwSession.state.ready || !files || !files.length) {
      return { ok: false, message: "No ready LSMW BOM files available." };
    }
    fileType = fileType || "all";
    var selected = files;
    if (fileType !== "all") {
      var expected = String(fileType).toLowerCase();
      var ext = expected.slice(expected.lastIndexOf("."));
      var stem = expected.slice(0, expected.lastIndexOf(".")).split("/").pop();
      selected = files.filter(function (f) {
        return f.name.toLowerCase().indexOf(stem + "-") === 0 && f.name.toLowerCase().slice(f.name.lastIndexOf(".")) === ext;
      });
    }
    if (!selected.length) return { ok: false, message: "Requested LSMW BOM file not found." };
    var downloaded = [];
    selected.forEach(function (f) {
      downloadLsmwFile(f);
      downloaded.push(f.name);
    });
    if (fileType === "all") {
      return { ok: true, path: downloaded[0], paths: downloaded, message: "Downloaded all 4 LSMW BOM files to Downloads." };
    }
    return { ok: true, path: downloaded[0], paths: downloaded, message: "Downloaded " + downloaded[0] + " to Downloads." };
  }

  async function clearLsmwBomSession() {
    lsmwSession.state = initialLsmwState();
    lsmwSession.files = null;
    return { ok: true, message: "LSMW BOM session cleared.", state: lsmwSession.state };
  }

  async function importRoutingMaterial() {
    try {
      requireSheetJS();
      var file = await pickFile(".xlsx,.xls,.csv");
      if (!file) return { ok: false, canceled: true };
      var sourceRows = await readInputWorkbook(file);
      if (!sourceRows.length) return { ok: false, message: "No valid Material Code rows were found." };
      var lookupData = await fetchMasterJson("/web/masters-data/routing-lookup.json");
      var lookup = getRoutingLookup(lookupData);
      var matched = matchRoutingRows(lookup, sourceRows);
      if (!matched.importedRows) return { ok: false, message: "A Material Number or Material Code column is required." };
      var dateStr = todayDateStr();
      var readyFile = "Routing Upload File - Ready-" + dateStr + ".xlsx";
      var hasUnmatched = matched.unmatchedKeys.length > 0;
      var next = Object.assign(initialRoutingState(), {
        materialFile: file.name, generatedAt: new Date().toISOString(),
        importedRows: matched.importedRows, matchedRows: matched.routingRows.length,
        unmatchedRows: matched.unmatchedKeys.length, unmatchedKeys: matched.unmatchedKeys,
        readyFile: hasUnmatched ? null : readyFile,
        ready: !hasUnmatched,
      });
      if (hasUnmatched) {
        routingSession.workbook = null;
        routingSession.filename = null;
      } else {
        routingSession.workbook = buildRoutingWorkbook(matched.routingRows);
        routingSession.filename = readyFile;
      }
      routingSession.state = next;
      return {
        ok: !hasUnmatched,
        message: hasUnmatched
          ? matched.unmatchedKeys.length + " Routing link(s) were not found in the Master file. Admin must update the Routing Master."
          : "Routing processing completed: " + matched.routingRows.length + " SAP-ready rows generated.",
        state: Object.assign({}, next, { masterReady: true }),
      };
    } catch (error) {
      return { ok: false, message: error && error.message ? error.message : String(error) };
    }
  }

  async function downloadRoutingReady() {
    if (!routingSession.workbook || !routingSession.state.ready) {
      return { ok: false, message: "No ready Routing Upload file available." };
    }
    var filename = downloadXlsxWorkbook(routingSession.workbook, routingSession.filename);
    return { ok: true, path: filename, message: "Downloaded " + filename + " to Downloads." };
  }

  // The desktop uses one clear handler for both Routing and ZKNG
  // (erp:clear-routing-session, kind "routing" | "zkng"); the renderer
  // calls it the same way.
  async function clearRoutingSession(kind) {
    if (kind === "zkng") {
      zkngSession.state = initialZkngState();
      zkngSession.workbook = null;
      zkngSession.filename = null;
      return { ok: true, message: "ZKNG session cleared.", state: zkngSession.state };
    }
    routingSession.state = initialRoutingState();
    routingSession.workbook = null;
    routingSession.filename = null;
    return { ok: true, message: "Routing session cleared.", state: routingSession.state };
  }

  async function importCharacteristicsMaterial() {
    try {
      requireSheetJS();
      var file = await pickFile(".xlsx,.xls,.csv");
      if (!file) return { ok: false, canceled: true };
      var matRows = await readInputWorkbook(file);
      if (!matRows.length) return { ok: false, message: "No valid rows were found in the selected Material Code file." };
      var dateStr = todayDateStr();
      var generated = buildCharacteristicsWorkbook(matRows);
      var readyName = "Characteristics Upload File - Ready-" + dateStr + ".xlsx";
      var nextState = Object.assign(initialCharacteristicsState(), {
        materialFile: file.name, generatedAt: new Date().toISOString(),
        importedRows: generated.rowCount, matchedRows: generated.rowCount, unmatchedRows: 0,
        materialReadyFile: readyName, materialReady: true, ready: true,
      });
      characteristicsSession.state = nextState;
      characteristicsSession.workbook = generated.wb;
      characteristicsSession.filename = readyName;
      return { ok: true, message: "Processing completed: " + generated.rowCount + " rows added to the Characteristics Upload Ready file.", state: nextState };
    } catch (error) {
      return { ok: false, message: error && error.message ? error.message : String(error) };
    }
  }

  async function downloadCharacteristicsReady() {
    if (!characteristicsSession.workbook || !characteristicsSession.state.materialReady) {
      return { ok: false, message: "No ready Characteristics Upload file available." };
    }
    var filename = downloadXlsxWorkbook(characteristicsSession.workbook, characteristicsSession.filename);
    return { ok: true, path: filename, message: "Downloaded " + filename + " to Downloads." };
  }

  async function clearCharacteristicsSession() {
    characteristicsSession.state = initialCharacteristicsState();
    characteristicsSession.workbook = null;
    characteristicsSession.filename = null;
    return { ok: true, message: "Characteristics session cleared.", state: characteristicsSession.state };
  }

  window.erp = {
    // ---- real in Phase 1 (API-backed) ----
    getState: getState,
    getAccessState: getAccessState,
    update: update,
    saveSettings: async function (settings) { return settings || {}; },
    exportRows: exportRows,

    // ---- service states: free upload modules (in-memory sessions) ----
    getZkngState,
    getLsmwBomState,
    getRoutingState,
    getCharacteristicsState,

    // ---- harmless counters (desktop handlers are stubs too) ----
    recordUsage: async function () { return { ok: true }; },
    accountHeartbeat: async function () { return { ok: true }; },

    // ---- auth: real single-admin login since Phase 2-lite ----
    // renderer.js calls adminLogin({ username, password, remember }) and
    // expects { ok, access, message } — same shape as the desktop handler.
    // On success the session token is kept in sessionStorage; getState /
    // getAccessState then return the main-admin access object.
    adminLogin: async function (payload) {
      var result;
      try {
        result = await apiPostJson("/api/login", {
          username: (payload && payload.username) || "",
          password: (payload && payload.password) || "",
        });
      } catch (error) {
        return { ok: false, message: "অনলাইনে সাইন ইন করা যায়নি — পরে আবার চেষ্টা করুন।" };
      }
      if (result.ok && result.token) setToken(result.token);
      return result;
    },
    registerAccount: async function () { return notAvailable(PHASE2_LOGIN); },
    logout: async function () {
      clearToken();
      return { ok: true, message: "Signed out (Guest view)." };
    },
    changeOwnPassword: async function () { return notAvailable(PASSWORD_CHANGE_MESSAGE); },
    setAdminPassword: async function () { return notAvailable(PASSWORD_CHANGE_MESSAGE); },
    getRememberedLogin: async function () { return { savedAccounts: [], lastSelectedUsername: "" }; },
    getSavedAccounts: async function () { return { savedAccounts: [], lastSelectedUsername: "" }; },
    saveAccountLocally: async function () { return notAvailable(PHASE2_LOGIN); },
    removeSavedAccount: async function () { return notAvailable(PHASE2_LOGIN); },

    // ---- uploads / imports / generation / downloads ----
    // The four upload modules are free for everyone (client-side, no
    // login); only master replacement stays admin-only (Phase 2).
    uploadZkngMaster: async function () { return notAvailable(); },
    importZkngMaterial,
    downloadZkngReady,
    uploadLsmwBomMaster: async function () { return notAvailable(); },
    importLsmwBomMaterial,
    downloadLsmwBomReady,
    clearLsmwBomSession,
    uploadRoutingMaster: async function () { return notAvailable(); },
    importRoutingMaterial,
    downloadRoutingReady,
    clearRoutingSession,
    uploadCharacteristicsMaster: async function () { return notAvailable(); },
    importCharacteristicsMaterial,
    downloadCharacteristicsReady,
    clearCharacteristicsSession,
    uploadMasterSection: async function () { return notAvailable(); },
    clearMasterSection: async function () { return notAvailable(); },
    uploadDepartmentFile: async function () { return notAvailable(); },
    saveDepartmentFile: async function () { return notAvailable(); },
    downloadResource: async function () { return notAvailable(); },
    saveSubAdmin: async function () { return notAvailable(); },
    manageSubAdmin: async function () { return notAvailable(); },
    changeAccountRole: async function () { return notAvailable(); },
    reviewSubAdminRequest: async function () { return notAvailable(); },
    manageDepartment: async function () { return notAvailable(); },

    // ---- desktop-only: safe no-ops in a browser ----
    chooseFolder: async function () { return null; },
    openPath: async function () { return undefined; },
    chooseProfileImage: async function () { return null; },
    reloadApp: async function () { return notAvailable(DESKTOP_ONLY); },
    checkSoftwareUpdate: async function () {
      return {
        ok: true, available: false,
        message: "Online-এ সফ্টওয়্যার আপডেট নেই — নতুন ভার্সন এলে ডেস্কটপ ইনস্টলার দিয়ে আপডেট করুন।",
      };
    },
    installSoftwareUpdate: async function () { return notAvailable(DESKTOP_ONLY); },
    publishSoftwareUpdate: async function () { return notAvailable(DESKTOP_ONLY); },
  };
})();
