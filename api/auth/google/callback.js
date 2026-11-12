// GET /api/auth/google/callback — finish "Login with Google".
//
// Exchanges the code for tokens, reads the Google account email, and
// signs the user in ONLY if that email is on the allowlist
// (GOOGLE_ALLOWED_EMAILS, comma-separated; default: ZBOM_ADMIN_USERNAME).
// On success it returns a tiny page that stores the same HMAC session
// token the password login issues (sessionStorage "zbom-online-token")
// and sends the browser back to "/". On any failure it shows a plain
// error page. No password, code, or token is ever logged.
// Plain Node req/res only — no framework, works on Vercel and locally.
"use strict";

const crypto = require("node:crypto");
const auth = require("../../_auth.cjs");

const TOKEN_URL = "https://oauth2.googleapis.com/token";
const USERINFO_URL = "https://openidconnect.googleapis.com/v1/userinfo";

function sendHtml(res, status, title, body, token) {
  let extra = "";
  if (token) {
    extra =
      "<script>try{sessionStorage.setItem('zbom-online-token'," +
      JSON.stringify(token) +
      ");}catch(e){}location.replace('/');</script>";
  }
  const html =
    "<!doctype html><meta charset=utf-8><title>" +
    title +
    "</title><body style=\"font-family:sans-serif;max-width:560px;margin:60px auto;padding:0 16px\">" +
    "<h2>" +
    title +
    "</h2><p>" +
    body +
    "</p><p><a href=\"/\">Back to ZBOM</a></p>" +
    extra +
    "</body>";
  res.writeHead(status, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(html),
  });
  res.end(html);
}

function callbackUrl(req) {
  const host = String(
    (req.headers && (req.headers["x-forwarded-host"] || req.headers.host)) || "",
  )
    .split(",")[0]
    .trim()
    .toLowerCase();
  const ok =
    host === "bom.b2bgdc.com" || /^[a-z0-9-]+\.vercel\.app$/.test(host);
  return "https://" + (ok ? host : "bom.b2bgdc.com") + "/api/auth/google/callback";
}

function verifyState(state, secret) {
  if (!state || typeof state !== "string" || !secret) return false;
  const parts = state.split(".");
  if (parts.length !== 4 || parts[0] !== "v1") return false;
  const exp = Number(parts[1]);
  if (!Number.isFinite(exp) || exp <= Date.now()) return false;
  const data = parts.slice(0, 3).join(".");
  let expected;
  let given;
  try {
    expected = crypto.createHmac("sha256", secret).update(data).digest();
    given = Buffer.from(parts[3], "base64url");
  } catch (error) {
    return false;
  }
  return (
    given.length === expected.length && crypto.timingSafeEqual(given, expected)
  );
}

function allowedEmails() {
  const list = String(process.env.GOOGLE_ALLOWED_EMAILS || "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
  if (list.length) return list;
  const admin = String(process.env.ZBOM_ADMIN_USERNAME || "").trim().toLowerCase();
  return admin ? [admin] : [];
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.writeHead(405, { Allow: "GET" });
    res.end("Method Not Allowed");
    return;
  }
  const secret = String(process.env.ZBOM_SESSION_SECRET || "");
  const clientId = String(process.env.GOOGLE_CLIENT_ID || "").trim();
  const clientSecret = String(process.env.GOOGLE_CLIENT_SECRET || "");
  if (!clientId || !clientSecret || !secret) {
    sendHtml(
      res,
      503,
      "Google sign-in is not configured yet",
      "The admin needs to add GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in Vercel and redeploy.",
    );
    return;
  }
  const url = new URL(req.url || "/", "http://localhost");
  const code = url.searchParams.get("code") || "";
  const state = url.searchParams.get("state") || "";
  const error = url.searchParams.get("error") || "";
  if (error || !code || !verifyState(state, secret)) {
    sendHtml(
      res,
      400,
      "Google sign-in did not complete",
      "The sign-in was cancelled or the request expired. Please try again.",
    );
    return;
  }
  // Exchange the code for tokens.
  let tokens;
  try {
    const tokenRes = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        code: code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: callbackUrl(req),
        grant_type: "authorization_code",
      }),
    });
    tokens = await tokenRes.json();
  } catch (err) {
    tokens = null;
  }
  if (!tokens || !tokens.access_token) {
    sendHtml(
      res,
      502,
      "Google sign-in did not complete",
      "Google did not accept the sign-in. Please try again.",
    );
    return;
  }
  // Read the Google account email.
  let profile;
  try {
    const meRes = await fetch(USERINFO_URL, {
      headers: { Authorization: "Bearer " + tokens.access_token },
    });
    profile = await meRes.json();
  } catch (err) {
    profile = null;
  }
  const email = String((profile && profile.email) || "").trim().toLowerCase();
  if (!email || (profile && profile.email_verified === false)) {
    sendHtml(
      res,
      403,
      "Google sign-in did not complete",
      "Google did not return a verified email address.",
    );
    return;
  }
  if (!allowedEmails().includes(email)) {
    sendHtml(
      res,
      403,
      "This Google account cannot sign in to ZBOM",
      "Only the admin email can sign in here.",
    );
    return;
  }
  const token = auth.issueToken(email, secret);
  if (!token) {
    sendHtml(
      res,
      500,
      "Google sign-in did not complete",
      "The session could not be created. Please try again.",
    );
    return;
  }
  sendHtml(res, 200, "Signed in with Google", "Taking you back to ZBOM…", token);
};
