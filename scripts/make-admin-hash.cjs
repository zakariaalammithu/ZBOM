// Generate the admin password hash for the Vercel env var
// ZBOM_ADMIN_PASSWORD_HASH (Phase 2-lite online login).
//
// Usage (type the password once, then Enter):
//
//   node scripts/make-admin-hash.cjs
//
// The password is read from stdin one line at a time with terminal echo
// turned off, never from argv, and never written to any file. The script
// prints ONLY the JSON value to paste into Vercel, plus one instruction
// line. The password itself is never printed, logged, or saved anywhere.
// The hash uses the desktop format exactly
// (desktop-client/main.cjs hashPassword): PBKDF2-SHA256, 150,000
// iterations, 32-byte key, random 16-byte hex salt.
//
// In Vercel: Project Settings -> Environment Variables, add
//   ZBOM_ADMIN_PASSWORD_HASH = <the JSON printed below>
// and ZBOM_SESSION_SECRET = <any long random string> (if not set yet),
// then redeploy. Changing the password later means generating a new
// hash here and updating that env var.
"use strict";

const crypto = require("node:crypto");

const MIN_LENGTH = 6;

function readPasswordLine(prompt) {
  return new Promise((resolve) => {
    process.stdout.write(prompt);
    const stdin = process.stdin;
    let input = "";
    const onData = (chunk) => {
      for (const byte of chunk) {
        if (byte === 0x0d || byte === 0x0a) {
          stdin.removeListener("data", onData);
          if (stdin.isTTY) stdin.setRawMode(false);
          stdin.pause();
          process.stdout.write("\n");
          resolve(input);
          return;
        }
        if (byte === 0x03) {
          // Ctrl-C: abort without printing anything.
          process.stdout.write("\n");
          process.exit(1);
        }
        if (byte === 0x7f) {
          input = input.slice(0, -1);
        } else if (byte >= 0x20) {
          input += String.fromCharCode(byte);
        }
      }
    };
    if (stdin.isTTY) stdin.setRawMode(true);
    stdin.resume();
    stdin.on("data", onData);
  });
}

async function main() {
  const password = await readPasswordLine("New online admin password (min 6 characters): ");
  if (typeof password !== "string" || password.length < MIN_LENGTH) {
    process.stderr.write("Password must contain at least 6 characters.\n");
    process.exit(1);
  }
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.pbkdf2Sync(password, salt, 150000, 32, "sha256").toString("hex");
  process.stdout.write(JSON.stringify({ salt, hash }) + "\n");
  process.stdout.write(
    "Paste the JSON above as the Vercel environment variable ZBOM_ADMIN_PASSWORD_HASH, then redeploy.\n",
  );
}

main();
