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
//     getCharacteristicsState — empty-but-valid states matching the shapes
//     renderer.js expects (see desktop main.cjs get-*-state handlers).
//     Import buttons stay disabled, downloads stay disabled: nothing is
//     generated yet.
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
//   - All master uploads, material imports, ZKNG / LSMW BOM / Routing /
//     Characteristics generation and downloads, department files, sub-admin
//     and role management — { ok: false } with a Phase 2 message.
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

  window.erp = {
    // ---- real in Phase 1 (API-backed) ----
    getState: getState,
    getAccessState: getAccessState,
    update: update,
    saveSettings: async function (settings) { return settings || {}; },
    exportRows: exportRows,

    // ---- service states: empty-but-valid shapes (nothing generated yet) ----
    getZkngState: async function () {
      return {
        masterReady: true, zkngMasterReady: true,
        zkngMasterFile: "ZKNG Master.xlsx", zkngMasterUpdatedAt: null,
        canManageMaster: false, canImport: false,
        materialFile: null, generatedAt: null,
        importedRows: 0, matchedRows: 0, unmatchedRows: 0,
        materialReadyFile: null, materialReady: false, ready: false,
      };
    },
    getLsmwBomState: async function () {
      return {
        masterReady: true, masterFile: "LSMW BOM Upload Master File.xlsx",
        masterUpdatedAt: null, canManageMaster: false, canImport: false,
        materialFile: null, generatedAt: null, importedRows: 0,
        headerRowsCount: 0, itemRowsCount: 0, generatedFiles: [], ready: false,
      };
    },
    getRoutingState: async function () {
      return {
        masterReady: true, masterFile: "Built-in Routing Master",
        masterUpdatedAt: null, canManageMaster: false, canImport: false,
        materialFile: null, generatedAt: null,
        importedRows: 0, matchedRows: 0, unmatchedRows: 0, unmatchedKeys: [],
        readyFile: null, ready: false,
      };
    },
    getCharacteristicsState: async function () {
      return {
        // The Characteristics master is NOT bundled (desktop: admin uploads
        // it), so online it is honestly not ready until Phase 2 uploads.
        masterReady: false, canManageMaster: false, canImport: false,
        materialFile: null, generatedAt: null,
        importedRows: 0, matchedRows: 0, unmatchedRows: 0,
        readyFile: null, ready: false,
      };
    },

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

    // ---- uploads / imports / generation / downloads: Phase 2 ----
    uploadZkngMaster: async function () { return notAvailable(); },
    importZkngMaterial: async function () { return notAvailable(); },
    downloadZkngReady: async function () { return notAvailable(); },
    uploadLsmwBomMaster: async function () { return notAvailable(); },
    importLsmwBomMaterial: async function () { return notAvailable(); },
    downloadLsmwBomReady: async function () { return notAvailable(); },
    clearLsmwBomSession: async function () { return notAvailable(); },
    uploadRoutingMaster: async function () { return notAvailable(); },
    importRoutingMaterial: async function () { return notAvailable(); },
    downloadRoutingReady: async function () { return notAvailable(); },
    clearRoutingSession: async function () { return notAvailable(); },
    uploadCharacteristicsMaster: async function () { return notAvailable(); },
    importCharacteristicsMaterial: async function () { return notAvailable(); },
    downloadCharacteristicsReady: async function () { return notAvailable(); },
    clearCharacteristicsSession: async function () { return notAvailable(); },
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
