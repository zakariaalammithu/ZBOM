// GET /api/auth/google — start "Login with Google" (Phase 2-lite).
//
// Redirects the browser to Google's authorization page. The admin must
// first create a Google Cloud OAuth 2.0 client (web application) and set
// GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET in Vercel; the authorized
// redirect URIs to register in Google are:
//   https://bom.b2bgdc.com/api/auth/google/callback
//   https://zbom.vercel.app/api/auth/google/callback
// Only the admin email(s) can finish sign-in — see callback.js.
// Plain Node req/res only — no framework, works on Vercel and locally.
"use strict";

const crypto = require("node:crypto");

const GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";

function sendHtml(res, status, title, body) {
  const html =
    "<!doctype html><meta charset=utf-8><title>" +
    title +
    "</title><body style=\"font-family:sans-serif;max-width:560px;margin:60px auto;padding:0 16px\">" +
    "<h2>" +
    title +
    "</h2><p>" +
    body +
    "</p><p><a href=\"/\">Back to ZBOM</a></p></body>";
  res.writeHead(status, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(html),
  });
  res.end(html);
}

// The OAuth redirect URI must match exactly what is registered in
// Google. Derive it from the request host, allowing only this
// project's own domains; anything else falls back to the primary one.
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

function makeState(secret) {
  const exp = Date.now() + 10 * 60 * 1000;
  const nonce = crypto.randomBytes(16).toString("hex");
  const data = "v1." + exp + "." + nonce;
  const sig = crypto.createHmac("sha256", secret).update(data).digest("base64url");
  return data + "." + sig;
}

module.exports = (req, res) => {
  if (req.method !== "GET") {
    res.writeHead(405, { Allow: "GET" });
    res.end("Method Not Allowed");
    return;
  }
  const clientId = String(process.env.GOOGLE_CLIENT_ID || "").trim();
  const secret = String(process.env.ZBOM_SESSION_SECRET || "");
  if (!clientId || !secret) {
    sendHtml(
      res,
      503,
      "Google sign-in is not configured yet",
      "The admin needs to add GOOGLE_CLIENT_ID (and ZBOM_SESSION_SECRET) in Vercel and redeploy.",
    );
    return;
  }
  const params = new URLSearchParams({
    client_id: clientId,
    redirect_uri: callbackUrl(req),
    response_type: "code",
    scope: "openid email profile",
    state: makeState(secret),
    prompt: "select_account",
  });
  res.writeHead(302, { Location: GOOGLE_AUTH_URL + "?" + params.toString() });
  res.end();
};
