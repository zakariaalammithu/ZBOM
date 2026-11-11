# ZBOM

ZBOM v1.0.1 — offline Windows ERP client for knitting material and ZSummary data, built with Electron 37.2.6.

## What it does

- Overview dashboard with material records, finish types, department files and last sync
- Fabric Material Code and Finish Type Short Name search, filters and CSV/Excel export
- Circular Knitting summary and Dia Wise Summary
- SAP-ready file generators: ZKNG Material Upload, LSMW BOM Upload (Header/Item, .xlsx + .txt), Routing Upload, Characteristics Upload and Code Requirement
- Departments: Circular Knitting, Flat Knitting, Yarn Processing, Fabric Processing, Fabric Finishing, Fabric Quality
- Admin and account roles, sub-admins, password management and software updates from a LAN shared folder

## Run

```sh
cd desktop-client
pnpm install
pnpm start
```

Preview mode (auto-refresh, no install): double-click `desktop-client/Start-ZBOM-Preview.cmd`.
Production launcher: `Launch-ZBOM-Desktop.cmd`.

## Build

```sh
cd desktop-client
pnpm run dist
```

This creates the Windows installer in `desktop-client/release/` (`ZBOM-Setup-1.0.1.exe`).

## Data

The app reads its source files from a LAN shared folder (default `\\172.26.43.3\ZKNG- Material code`) and keeps an offline cache. The `seed-data/` folder inside `desktop-client/` holds the bundled master templates, and the top-level `Code UPload file/`, `LSMW BOM Upload file/` and `Routing/` folders hold example input and ready output files.
