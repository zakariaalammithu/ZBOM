// GET /api/state — Phase 1 online state for ZBOM.
//
// Mirrors the shape of desktop main.cjs `erp:get-state` (config, snapshot,
// appVersion, access, departments, departmentStatuses, departmentFiles,
// authenticated, profileImageDataUrl, codeRequirementPreview) but with
// honest Phase 1 values: no account is signed in (guest, authenticated:
// false), no summary/finish-type data is loaded yet, and the material
// snapshot itself comes from /api/materials (see web/web-erp.js).
//
// Phase 2-lite: if the request carries a valid admin session token
// (`Authorization: Bearer <token>`, issued by POST /api/login), the
// response carries the main-admin access object and authenticated: true
// instead of the guest response. Any other token (or none) gets exactly
// the guest response. Plain Node req/res only — no framework, works on
// Vercel and locally.
"use strict";

const auth = require("./_auth.cjs");

const DEPARTMENTS = [
  "Circular Knitting",
  "Flat Knitting",
  "Yarn Processing",
  "Fabric Processing",
  "Fabric Finishing",
  "Fabric Quality",
];

const GUEST_ACCESS = {
  role: "guest",
  fullName: "Guest User",
  username: "",
  employeeId: "",
  departments: [],
  subAdmins: [],
  permissions: ["view"],
};

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
  // Phase 2-lite admin session: a valid Bearer token upgrades this
  // response from guest to the main-admin access object. Anything else
  // (no token, bad signature, expired) stays guest — silently.
  let authenticated = false;
  let access = GUEST_ACCESS;
  const header = String((req.headers && req.headers.authorization) || "");
  const match = /^Bearer\s+(.+)$/.exec(header);
  if (match) {
    const config = auth.loadAdminConfig();
    const username = auth.verifyToken(match[1].trim(), config.secret);
    if (username && username.toLowerCase() === config.username) {
      authenticated = true;
      access = auth.adminAccess(config.username);
    }
  }
  sendJson(res, 200, {
    appVersion: "1.0.1",
    phase: 1,
    online: true,
    config: { sharePath: "", autoUpdate: false },
    snapshot: {
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
    },
    access,
    departments: DEPARTMENTS,
    departmentStatuses: {},
    departmentFiles: [],
    usageAnalytics: [],
    authenticated,
    profileImageDataUrl: "",
    codeRequirementPreview: { rows: [], file: "", sheetName: "" },
    materialsSource: {
      file: "Fabric Material code-.xlsx",
      sheet: "Mastar file",
      rows: 103710,
      note: "Search via GET /api/materials?q=... — the full master is never sent to the browser.",
    },
  });
};
