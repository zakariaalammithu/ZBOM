// api/_auth.cjs — shared auth helpers for the online API.
//
// Underscore-prefixed so Vercel does NOT treat this file as a serverless
// function; it is required by api/login.js and api/state.js.
//
// Phase 2-lite: ONE admin account only, configured by environment
// variables (see ONLINE_PLAN.md). There is no user database here.
//   ZBOM_ADMIN_USERNAME       — admin login email (e.g. mithusquare@gmail.com)
//   ZBOM_ADMIN_PASSWORD_HASH  — JSON {"salt","hash"} (PBKDF2, see below)
//   ZBOM_SESSION_SECRET       — long random string used to sign tokens
//
// Password hashing matches the desktop app exactly
// (desktop-client/main.cjs hashPassword): PBKDF2-SHA256, 150,000
// iterations, 32-byte key, hex salt (16 random bytes) and hex hash.
// Sessions last 30 minutes, same as the desktop SESSION_TIMEOUT_MS.
// Tokens are base64url(JSON payload) + "." + base64url(HMAC-SHA256).
//
// If any env var is missing or malformed, verification always fails —
// never crash, never accept. NO password or hash value may ever be
// written into this file or any other repo file.
"use strict";

const crypto = require("node:crypto");

const SESSION_TTL_MS = 30 * 60 * 1000;

const DEPARTMENTS = [
  "Circular Knitting",
  "Flat Knitting",
  "Yarn Processing",
  "Fabric Processing",
  "Fabric Finishing",
  "Fabric Quality",
];

// main_admin permission set, copied from desktop-client/main.cjs
// ROLE_PERMISSIONS.main_admin.
const MAIN_ADMIN_PERMISSIONS = [
  "view",
  "export",
  "upload",
  "import",
  "manage_all_data",
  "approve",
  "review",
  "department_control",
  "user_management",
  "system_control",
];

function adminAccess(username) {
  const loginId = String(username || "31001625").trim().toLowerCase();
  return {
    username: loginId,
    email: loginId,
    role: "main_admin",
    fullName: "Zakaria Alam Mithu",
    employeeId: "31001625",
    designation: "CTM",
    departments: DEPARTMENTS.slice(),
    subAdmins: [],
    permissions: MAIN_ADMIN_PERMISSIONS.slice(),
    passwordSet: true,
  };
}

function loadAdminConfig() {
  const username = String(process.env.ZBOM_ADMIN_USERNAME || "31001625").trim().toLowerCase();
  let credential = null;
  try {
    const parsed = JSON.parse(process.env.ZBOM_ADMIN_PASSWORD_HASH || "");
    if (parsed && typeof parsed.salt === "string" && typeof parsed.hash === "string") {
      credential = { salt: parsed.salt, hash: parsed.hash };
    }
  } catch (error) {
    credential = null;
  }
  const secret = process.env.ZBOM_SESSION_SECRET || "";
  return { username, credential, secret };
}

function verifyPassword(password, credential) {
  if (!credential || !credential.salt || !credential.hash) return false;
  let calculated;
  try {
    calculated = crypto
      .pbkdf2Sync(String(password), credential.salt, 150000, 32, "sha256")
      .toString("hex");
  } catch (error) {
    return false;
  }
  let calculatedBuffer;
  let storedBuffer;
  try {
    calculatedBuffer = Buffer.from(calculated, "hex");
    storedBuffer = Buffer.from(credential.hash, "hex");
  } catch (error) {
    return false;
  }
  return (
    calculatedBuffer.length === storedBuffer.length &&
    calculatedBuffer.length > 0 &&
    crypto.timingSafeEqual(calculatedBuffer, storedBuffer)
  );
}

function base64urlEncode(buffer) {
  return Buffer.from(buffer).toString("base64url");
}

function issueToken(username, secret) {
  if (!secret) return null;
  const payload = { u: username, role: "main_admin", exp: Date.now() + SESSION_TTL_MS };
  const payloadPart = base64urlEncode(JSON.stringify(payload));
  const signature = crypto.createHmac("sha256", secret).update(payloadPart).digest();
  return payloadPart + "." + base64urlEncode(signature);
}

function verifyToken(token, secret) {
  if (!token || typeof token !== "string" || !secret) return null;
  const parts = token.split(".");
  if (parts.length !== 2 || !parts[0] || !parts[1]) return null;
  let expected;
  try {
    expected = crypto.createHmac("sha256", secret).update(parts[0]).digest();
  } catch (error) {
    return null;
  }
  let given;
  try {
    given = Buffer.from(parts[1], "base64url");
  } catch (error) {
    return null;
  }
  if (given.length !== expected.length || !crypto.timingSafeEqual(given, expected)) {
    return null;
  }
  let payload;
  try {
    payload = JSON.parse(Buffer.from(parts[0], "base64url").toString("utf8"));
  } catch (error) {
    return null;
  }
  if (!payload || payload.role !== "main_admin" || typeof payload.u !== "string") return null;
  if (typeof payload.exp !== "number" || payload.exp <= Date.now()) return null;
  return payload.u;
}

module.exports = {
  SESSION_TTL_MS,
  DEPARTMENTS,
  MAIN_ADMIN_PERMISSIONS,
  adminAccess,
  loadAdminConfig,
  verifyPassword,
  issueToken,
  verifyToken,
};
