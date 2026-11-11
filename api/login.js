// POST /api/login — Phase 2-lite single-admin login for ZBOM online.
//
// The ONE admin account is configured by Vercel environment variables
// (see ONLINE_PLAN.md): ZBOM_ADMIN_USERNAME (default "31001625"),
// ZBOM_ADMIN_PASSWORD_HASH (JSON {"salt","hash"}, PBKDF2 — generate it
// with scripts/make-admin-hash.cjs), ZBOM_SESSION_SECRET (token signing).
// There is no user database; no other account can sign in.
//
// Success: 200 { ok: true, token, access } — token is an HMAC-signed
// session token, valid 30 minutes (see api/_auth.cjs). The browser sends
// it back as `Authorization: Bearer <token>`.
// Failure — wrong username, wrong password, or env vars unset — is
// always 401 { ok: false, message: "Incorrect username or password." }.
// The message never says which field failed, and a missing server
// configuration looks exactly like a wrong password.
// Plain Node req/res only — no framework, works on Vercel and locally.
"use strict";

const auth = require("./_auth.cjs");

const FAIL = { ok: false, message: "Incorrect username or password." };

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(body),
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve) => {
    let data = "";
    req.on("data", (chunk) => {
      data += chunk;
      // Never accept an unbounded body for a login request.
      if (data.length > 8192) req.destroy();
    });
    req.on("end", () => resolve(data));
    req.on("error", () => resolve(""));
  });
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method Not Allowed");
    return;
  }
  let body = {};
  try {
    body = JSON.parse((await readBody(req)) || "{}");
  } catch (error) {
    sendJson(res, 401, FAIL);
    return;
  }
  const config = auth.loadAdminConfig();
  const username = String(body.username || "").trim().toLowerCase();
  const password = String(body.password || "");
  // Diagnostic only: log WHICH check failed to Vercel Runtime Logs so a
  // persistent 401 can be diagnosed. The 401 response stays identical,
  // and nothing secret is ever logged (no password, no hash, no token).
  let failReason = null;
  if (!config.credential) {
    failReason = "no credential (ZBOM_ADMIN_PASSWORD_HASH missing or not JSON)";
  } else if (!config.secret) {
    failReason = "no secret (ZBOM_SESSION_SECRET missing)";
  } else if (username !== config.username) {
    failReason = "username mismatch";
  } else if (!auth.verifyPassword(password, config.credential)) {
    failReason = "password mismatch";
  }
  if (failReason) {
    console.log("ZBOM login fail: " + failReason);
    sendJson(res, 401, FAIL);
    return;
  }
  const token = auth.issueToken(config.username, config.secret);
  if (!token) {
    sendJson(res, 401, FAIL);
    return;
  }
  sendJson(res, 200, { ok: true, token, access: auth.adminAccess() });
};
