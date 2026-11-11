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
- `data/materials-index.json` (17.1 MB) — pre-extracted compact index
  (columns once, rows as arrays). Regenerate with `npm run build-index`
  (`scripts/build-materials-index.cjs`) after the master workbook changes.
- `vercel.json` — `/` → `/web/index.html`; `api/materials.js` bundles
  `data/materials-index.json` via `includeFiles`. No framework, plain Node.
- `.vercelignore` — keeps the real masters, `desktop-client/` and `data/`
  out of the public static deployment (otherwise anyone could download the
  master workbooks, and the desktop renderer's fake-admin stub would be
  reachable on the web).
- Root `package.json` — only for `xlsx` (used by the index builder) and
  the `build-index` script. The API itself needs no dependencies.

### Why the index file exists (measured, not guessed)

- Parsing `Fabric Material code-.xlsx` (10.3 MB) with `xlsx` 0.18.5 takes
  **~11.6 s and ~818 MB RSS** — beyond a Vercel Hobby cold start
  (10 s / 1024 MB). Full-row JSON would be 46.6 MB.
- The array index (17.1 MB) parses in **~0.2 s**; warm searches take
  **~0.1–0.3 s**. First load is cached in the function's module scope.
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
- `git status`: **no modifications under `desktop-client/`** — the desktop
  app is untouched. Everything new is uncommitted (no commit, no push).

### ⚠️ Verify on the FIRST Vercel deploy

1. `GET /api/materials?q=TERRY` must return rows. If it 500s with
   "materials index could not be loaded", the `includeFiles` +
   `.vercelignore` interaction needs revisiting (fallback: remove `data`
   from `.vercelignore` — but then the index becomes publicly downloadable
   at `/data/materials-index.json`, a data-exposure tradeoff to decide then).
2. Open the deployed site and type a material code in the search box —
   results must come from the server (total > 50 proves it).

## What does NOT work online yet (Phase 1)

Everything below shows the UI but does nothing real, or is disabled. This
is deliberate — each needs a database (Phase 2).

- **Sign-in / accounts**: Guest only. `adminLogin`, `registerAccount`,
  `changeOwnPassword`, `setAdminPassword`, saved accounts — all return
  `{ok:false, message:"অনলাইন লগইন Phase 2-এ আসছে (ডেটাবেস লাগবে)"}`.
  No passwords, no invented accounts anywhere in Phase 1 code.
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
