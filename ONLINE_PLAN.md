# ZBOM Online — Plan (Phase 1 done, Phase 2–3 ahead)

## Phase 1 — DONE (this scaffold, uncommitted, for review)

Goal: prove the ZBOM UI runs on the web (Vercel) with **real material search**.
It is NOT the full online system.

### What was built

- `web/` — copy of `desktop-client/renderer/` (`index.html`, `renderer.js`,
  `styles.css`, `assets/`) plus:
  - `web/web-erp.js` — defines `window.erp` in the browser BEFORE
    `renderer.js` loads (script order fixed in the copied `index.html`).
    Real: `getState`/`getAccessState`/`update` via `/api/state` +
    `/api/materials` (guest, first 50 real rows); service states
    (`getZkngState`, `getLsmwBomState`, `getRoutingState`,
    `getCharacteristicsState`) as empty-but-valid shapes; `exportRows` as
    client-side CSV download; `saveSettings` echoes back;
    `recordUsage`/`accountHeartbeat` return `{ok:true}`.
  - `web/web-online.js` — loaded AFTER `renderer.js`; intercepts the
    `#materialSearch` input and searches the whole master server-side via
    `/api/materials?q=...`, then re-renders. Brief "0 matching records"
    flash is possible before server results arrive.
  - The desktop `index.html` fallback stub is REMOVED in the web copy: it
    fabricated a signed-in admin ("Zakaria Alam Mithu", role `admin`,
    `authenticated: true`) with no login. Online there is no such account.
- `api/state.js` — `GET /api/state`: guest state, appVersion `1.0.1`,
  the 6 departments, `authenticated: false`.
- `api/materials.js` — `GET /api/materials?q=&limit=50`: case-insensitive
  substring search across all 17 columns of the Fabric Material master
  (103,710 rows), returns at most `limit` rows (default 50, max 100) +
  the true total. The full master is never sent to the browser.
- `api/materials-data/materials-index-1.json` … `-4.json` (4 chunks, ~4.0–4.5 MB
  each, 17.1 MB total) — pre-extracted compact index (columns once, rows
  as arrays, 103,710 rows). Regenerate with `npm run build-index`
  (`scripts/build-materials-index.cjs`) after the master workbook changes.
  **Why chunks, not one file:** the single 17.1 MB `materials-index.json`
  would not push to GitHub — the SSH proxy tunnel broke with "Broken
  pipe" on every attempt containing that file (6 failures), while the
  code-only push went through fine. So the index is split into 4 chunks
  of ~26k rows each, committed and pushed one chunk per commit.
  `api/materials.js` globs `api/materials-data/materials-index-*.json` and concatenates
  the rows, so a partial upload still serves the rows it has; with no
  chunk at all it returns a clear 500 ("materials index is not fully
  uploaded yet").
- `vercel.json` — `/` → `/web/index.html`; `api/materials.js` bundles
  `api/materials-data/materials-index-*.json` (glob) via `includeFiles`. No framework,
  plain Node.
- `.vercelignore` — keeps the real masters and `desktop-client/`
  out of the public static deployment (otherwise anyone could download the
  master workbooks, and the desktop renderer's fake-admin stub would be
  reachable on the web).
- Root `package.json` — only for `xlsx` (used by the index builder) and
  the `build-index` script. The API itself needs no dependencies.

### Why the index file exists (measured, not guessed)

- Parsing `Fabric Material code-.xlsx` (10.3 MB) with `xlsx` 0.18.5 takes
  **~11.6 s and ~818 MB RSS** — beyond a Vercel Hobby cold start
  (10 s / 1024 MB). Full-row JSON would be 46.6 MB.
- The array index (17.1 MB, split into 4 chunks of ~4.0–4.5 MB) parses
  in **~0.2 s**; warm searches take **~0.1–0.3 s**. First load is cached
  in the function's module scope.
- Local verification (see "Verified locally" below) ran the same code the
  Vercel functions use.

### Verified locally (2026-11-11, Node v24.22.1)

- `node --check` passes on `web/web-erp.js`, `web/web-online.js`,
  `api/state.js`, `api/materials.js`, `scripts/build-materials-index.cjs`.
- Local static+API test server (in `/tmp`, not in the repo):
  - `GET /api/state` → 200, `appVersion "1.0.1"`, `authenticated: false`,
    6 departments.
  - `GET /api/materials?q=` → 200, first 50 rows of 103,710.
  - `GET /api/materials?q=TERRY&limit=50` → 200, 3,985 total matches,
    50 rows returned, first row Grey Material `1400000000`.
  - `GET /api/materials?q=1400000001` → 200, 1 total match (exact code);
    `q=terry` (lowercase) → same 3,985 (case-insensitive); `limit=500`
    is capped to 100; a no-match query returns `total: 0`.
  - `GET /` → 200, serves `web/index.html`, and the script order in it is
    `web-erp.js` → `renderer.js` → `web-online.js`.
  - Cold index load (fresh Node process, first `/api/materials` call):
    **479 ms, RSS 187 MB** — inside Vercel Hobby limits (10 s / 1024 MB).
    Warm search: **~34 ms**. The raw `.xlsx` parse it replaces took
    ~11.6 s / ~818 MB RSS.
  - Chunked-index re-verification (2026-11-11, after splitting into 4
    chunks): mock req/res test against the chunk files — `q=TERRY` →
    3,985 total (first row Grey Material `1400000000`),
    `q=1400000001` → 1, `q=zzzz-no-match` → 0, empty `q` → 103,710,
    lowercase `q=terry` → same 3,985 — all pass, cold load 549 ms.
    Edge cases: no chunk file at all → 500 "materials index is not fully
    uploaded yet"; only chunk 1 present → 200 with 25,928 rows (partial
    upload still serves what it has).
- `git status`: **no modifications under `desktop-client/`** — the desktop
  app is untouched. Everything new is uncommitted (no commit, no push).

### ⚠️ Verify on the FIRST Vercel deploy

**Fixed 2026-11-11 (after the first live check failed): two bugs found on
zbom.vercel.app — see "Live fixes" below. Re-verify after this fix:**
1. `GET /api/materials?q=TERRY` must return rows. If it 500s with
   "materials index is not fully uploaded yet", a chunk file is missing
   from the deployment (check the Deployments file list for
   `api/materials-data/materials-index-*.json`).
2. Open the deployed site — it must look styled (blue sidebar, cards),
   not plain unstyled text. If it is unstyled, an asset path is wrong
   again (check the browser console for 404s).
3. Type a material code in the search box — results must come from the
   server (total > 50 proves it).

### Live fixes (2026-11-11, after the user sent a screenshot of zbom.vercel.app)

1. **Unstyled page — relative asset paths.** `web/index.html` used
   relative paths (`styles.css`, `renderer.js`, `web-erp.js`,
   `web-online.js`, `assets/zcode.ico`), but Vercel serves it at `/`
   via a rewrite, so the browser asked for `/styles.css`, `/renderer.js`
   etc. — all 404. Without CSS every normally-hidden element showed at
   once, which is exactly the mess in the screenshot. Fixed: all refs
   are now absolute under `/web/` (`/web/styles.css`, `/web/renderer.js`,
   `/web/web-erp.js`, `/web/web-online.js`, `/web/assets/zcode.ico`).
2. **`/api/materials` 500 — index never bundled.** Live error was
   `ENOENT: scandir '/var/task/data'`: `.vercelignore` excluded `data/`,
   so `includeFiles` never bundled the chunks. Fixed: chunks moved to
   `api/materials-data/` (files under `api/` are bundled into the
   function but never served as public static files), `api/materials.js`
   now reads `path.join(__dirname, "materials-data")`, the builder
   writes there, `vercel.json` `includeFiles` is
   `api/materials-data/materials-index-*.json`, and the `data` line is
   gone from `.vercelignore`. Do not move the chunks back to `data/`.

## What does NOT work online yet (Phase 1)

Everything below shows the UI but does nothing real, or is disabled. This
is deliberate — each needs a database (Phase 2).

- **Sign-in / accounts**: Guest only in Phase 1; since Phase 2-lite the
  single env-var admin can sign in (`adminLogin` → `POST /api/login`,
  30-minute token). `registerAccount`, `changeOwnPassword`,
  `setAdminPassword` still return `{ok:false}` — the online admin
  password is changed in Vercel (`ZBOM_ADMIN_PASSWORD_HASH`), not in
  the app; other users still need the full Phase 2 database below.
  No passwords, no invented accounts anywhere in the code.
- **Circular Knitting summary + Dia Wise Summary**: no data loaded (0 rows).
- **Finish Type Short Name**: no data loaded (0 rows).
- **Code Requirement preview**: empty.
- **ZKNG / LSMW BOM / Routing / Characteristics generators**: no master
  upload, no material import, no Ready-file download. Import buttons are
  disabled (`canImport: false`); the service headers say "Ready" only
  because the secured templates exist in the repo, not because a file can
  be generated. Characteristics honestly reports `masterReady: false`
  (its master is not bundled — desktop admin uploads it).
- **Department workspace**: no department files/statuses/uploads.
- **Administration**: no sub-admins, role changes, user management, or
  admin dashboard data.
- **Settings**: the LAN share path is meaningless online; `saveSettings`
  only echoes back, nothing is stored.
- **Export**: `exportRows` (CSV via Blob) is implemented, but
  `renderer.js` only calls it after a sign-in + export-permission check,
  so the export buttons open the sign-in dialog in Phase 1. Excel
  (`.xlsx`) export is not implemented in the browser (no SheetJS loaded);
  it returns a clear message asking for CSV instead.
- **Software update, profile photo, "Open folder"**: desktop-only,
  safe no-ops online.
- **Overview counts**: material count shows the true total (103,710)
  after a search; summary / finish-type counts are 0 (no data yet).

## Phase 2-lite — Vercel env-var admin login (DONE, uncommitted)

One admin can now sign in on the website — no database needed. The admin
account is configured by Vercel environment variables; the password hash
uses the desktop format exactly (PBKDF2-SHA256, 150,000 iterations,
32-byte key — see `desktop-client/main.cjs` `hashPassword`).

### Env vars (Vercel Project Settings → Environment Variables)

| Name | Value format |
|---|---|
| `ZBOM_ADMIN_USERNAME` | Admin username, e.g. `31001625` (default if unset: `31001625`; compared case-insensitively) |
| `ZBOM_ADMIN_PASSWORD_HASH` | JSON `{"salt":"<32 hex chars>","hash":"<64 hex chars>"}` — generate with `node scripts/make-admin-hash.cjs` (reads the password from stdin with echo off, never from argv, never writes it anywhere) |
| `ZBOM_SESSION_SECRET` | Any long random string (token signing — keep it secret, never commit it) |

After adding/changing them: **redeploy** in Vercel (Deployments → ⋯ →
Redeploy), because serverless functions read env vars at deploy time.

### What was built

- `api/_auth.cjs` — shared helpers (underscore-prefixed, so Vercel does
  not treat it as a function): PBKDF2 verification with
  `timingSafeEqual`, HMAC-SHA256 session tokens
  (`base64url(payload).base64url(signature)`, payload `{u, role:
  "main_admin", exp}`, 30-minute expiry like desktop
  `SESSION_TIMEOUT_MS`), and the main-admin access object (username
  `31001625`, full name "Zakaria Alam Mithu", designation "CTM", all 6
  departments, the `main_admin` permission set copied from desktop
  `ROLE_PERMISSIONS`). Missing/malformed env vars make verification
  always fail — never crash, never accept.
- `api/login.js` — `POST /api/login` `{username, password}` → on success
  `200 {ok:true, token, access}`; wrong username, wrong password, or
  unset env vars all return the same `401 {ok:false, message:
  "Incorrect username or password."}` (never says which field failed).
- `api/state.js` — reads `Authorization: Bearer <token>`; a valid token
  returns `authenticated: true` with the main-admin access object,
  anything else returns exactly the guest response.
- `web/web-erp.js` — real `adminLogin` (POST `/api/login`, token kept in
  `sessionStorage` as `zbom-online-token`, sent on later `/api/state`
  calls, cleared on 401); `getState`/`getAccessState` now trust the
  server's `authenticated`/`access` instead of always returning guest;
  `logout` clears the token; `changeOwnPassword`/`setAdminPassword`
  return `{ok:false}` with the message that the online admin password
  cannot be changed from inside the app — change
  `ZBOM_ADMIN_PASSWORD_HASH` in Vercel and redeploy. `registerAccount`
  still returns the Phase 2 (Supabase) message.
- `scripts/make-admin-hash.cjs` — generates the hash JSON (see above).
- No password or hash value is stored in any repo file.

### Limits (honest)

- **Single admin only** — no other users, no sign-up, no sub-admins.
- **No in-app password change** — the password lives in a Vercel env
  var, so changing it means generating a new hash and redeploying.
- The desktop admin password (LAN share `ZBOM-Admin/access.json`) is
  **separate** — changing one does not change the other.
- Everything else from "What does NOT work online yet" still applies:
  master uploads, ZKNG/LSMW/Routing/Characteristics generation, Circular
  Knitting summary, Finish Type, Code Requirement, departments,
  administration — those still need the full Phase 2 database below.

## Phase 2 — full online system (needs a database)

Proposed: **Supabase** (Postgres + Auth + Storage) + the existing Vercel
functions as the API. Same `window.erp` method names, so the UI barely
changes; the work is behind the API.

- Accounts: Supabase Auth (proper password hashing, sessions, 30-min
  timeout like desktop) + a `profiles` table for role/department
  (`guest` → `main_admin`, same roles as desktop). Admin-created accounts
  vs open sign-up — **decision needed from the user**.
- Masters: Supabase Storage buckets for the master workbooks
  (Fabric Material, Circular Knitting summary, Finish Type, ZKNG/LSMW/
  Routing/Characteristics templates); upload = admin only; a `masters`
  table records version + updated-at. The LAN share stays as an import
  source, not the live store.
- Generators: move the `main.cjs` ZKNG / LSMW BOM / Routing /
  Characteristics logic into API functions that read the masters from
  Storage and return generated `.xlsx`/`.txt` downloads.
- Search/summary data: master rows in Postgres (not the JSON index) so
  summary, dia-wise totals and finish-type can be served too.
- Security: never put a service-role key in the browser; Row Level
  Security on; the 10 MB+ masters are not public files.

### What Phase 2 needs from the user

1. A **Supabase project** (user creates it at supabase.com, or gives
   access): project URL + anon key for the browser, service-role key for
   the server only (Vercel env vars — never committed, never in chat).
2. The account decision above (admin-created only, or open sign-up?).
3. Which master files are the current live ones (the repo-root copies, or
   newer ones from the LAN share?).

## Phase 3 — desktop joins the same system

- `desktop-client/main.cjs` gains an "online mode": its `erp:*` handlers
  call the same `/api/*` endpoints instead of the LAN share when online.
  `preload.cjs` and `renderer.js` do not change — they already talk to
  `window.erp` only.
- Offline: the desktop keeps its local cache (`loadCache`) and works
  offline; when back online it syncs (uploads/downloads since last sync).
  The LAN share becomes a legacy import path, not the live store.
- One account, one master, one dataset — desktop and web see the same
  data. That is the whole point of doing online first.
