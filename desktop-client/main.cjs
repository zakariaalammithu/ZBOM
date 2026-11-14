const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const os = require("node:os");
const { spawn } = require("node:child_process");
const JSZip = require("jszip/dist/jszip.js");
const XLSX = require("xlsx");

const DEFAULT_SHARE = "\\\\172.26.43.3\\ZKNG- Material code";
const MATERIAL_PATTERN = /fabric\s*material\s*code/i;
const SUMMARY_PATTERN = /(?:z|knitting)\s*summary|circular\s*knitting/i;
const FINISH_TYPE_PATTERN = /(?:finish|finsih)\s*type\s*short\s*name/i;
const DEFAULT_DEPARTMENTS = [
  "Circular Knitting",
  "Flat Knitting",
  "Yarn Processing",
  "Fabric Processing",
  "Fabric Finishing",
  "Fabric Quality",
];
const MAIN_ADMIN_USERS = new Set(["zakaria-sflfu", "zakaria", "31001625"]);

let mainWindow;
let previewWatcher;
let currentActiveSession = null;
const authenticatedAdminSessions = new Map();
const accountActivitySessions = new Map();
const isPreview = process.env.ZBOM_PREVIEW === "1";
const SESSION_TIMEOUT_MS = 30 * 60 * 1000;
const ROLES = ["guest", "standard_user", "power_user", "master_user", "department_head", "admin"];
const ROLE_PERMISSIONS = {
  guest: new Set(["view"]),
  standard_user: new Set(["view", "export"]),
  power_user: new Set(["view", "export", "upload", "import", "manage_own_data"]),
  master_user: new Set(["view", "export", "upload", "import", "manage_all_data"]),
  department_head: new Set(["view", "export", "upload", "import", "manage_own_data", "approve", "review", "department_control"]),
  admin: new Set(["view", "export", "upload", "import", "manage_all_data", "approve", "review", "department_control", "user_management", "system_control"]),
  main_admin: new Set(["view", "export", "upload", "import", "manage_all_data", "approve", "review", "department_control", "user_management", "system_control"]),
  user: new Set(["view", "export"]),
  sub_admin: new Set(["view", "export", "upload", "import", "manage_own_data"]),
};

function normalizeRole(role, isMainAdmin = false) {
  if (isMainAdmin || role === "main_admin") return "main_admin";
  if (role === "user") return "standard_user";
  if (role === "sub_admin") return "power_user";
  if (ROLES.includes(role)) return role;
  return "standard_user";
}

function roleLabel(role) {
  return ({ guest: "Guest User", standard_user: "Standard User", power_user: "Power User", master_user: "Master User", department_head: "Department Head", admin: "Admin", main_admin: "Admin" })[normalizeRole(role)] || "Standard User";
}

function hasPermission(access, permission) {
  return ROLE_PERMISSIONS[access?.role]?.has(permission) === true;
}

function sessionForEvent(event, touch = true) {
  const session = (event?.sender?.id ? authenticatedAdminSessions.get(event.sender.id) : null) || currentActiveSession;
  if (!session) return null;
  if (Date.now() - session.lastActivityAt > SESSION_TIMEOUT_MS) {
    authenticatedAdminSessions.clear();
    accountActivitySessions.clear();
    currentActiveSession = null;
    return null;
  }
  if (touch) session.lastActivityAt = Date.now();
  return session;
}

async function requirePermission(event, permission) {
  const session = sessionForEvent(event);
  if (!session) return { ok: false, message: "Your session expired. Please sign in again." };
  const access = await getCurrentAccess(session.username);
  if (!hasPermission(access, permission)) return { ok: false, message: "You do not have permission for this action." };
  return { ok: true, access };
}

function userFile(name) {
  return path.join(app.getPath("userData"), name);
}

async function readJson(file, fallback) {
  try {
    return JSON.parse(await fsp.readFile(file, "utf8"));
  } catch {
    return fallback;
  }
}

async function writeJsonAtomic(file, value) {
  const temporary = `${file}.${process.pid}-${Date.now()}-${crypto.randomBytes(4).toString("hex")}.new`;
  await fsp.mkdir(path.dirname(file), { recursive: true });
  const content = JSON.stringify(value);
  await fsp.writeFile(temporary, content);
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      await fsp.rename(temporary, file);
      return;
    } catch (error) {
      if (!["EPERM", "EACCES", "EEXIST"].includes(error.code) || attempt === 4) break;
      await new Promise((resolve) => setTimeout(resolve, 40 * (attempt + 1)));
    }
  }
  // Some Windows security/indexing tools briefly lock an existing JSON file.
  // A direct overwrite keeps the UI responsive when atomic replacement is blocked.
  await fsp.writeFile(file, content);
  try { await fsp.unlink(temporary); } catch {}
}

function formatDDMMYYYY(val) {
  if (!val && val !== 0) return "";
  if (val instanceof Date) {
    const d = val;
    if (isNaN(d.getTime())) return String(val);
    const day = String(d.getDate()).padStart(2, "0");
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const year = d.getFullYear();
    return `${day}.${month}.${year}`;
  }
  const str = String(val).trim();
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(str)) return str;
  const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[3]}.${isoMatch[2]}.${isoMatch[1]}`;
  const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slashMatch) {
    const p1 = String(slashMatch[1]).padStart(2, "0");
    const p2 = String(slashMatch[2]).padStart(2, "0");
    const y = slashMatch[3];
    return `${p1}.${p2}.${y}`;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const day = String(parsed.getDate()).padStart(2, "0");
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const year = parsed.getFullYear();
    return `${day}.${month}.${year}`;
  }
  return str;
}

async function updateZkngFileWithCurrentDate(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return;
  try {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, "0");
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const year = now.getFullYear();
    const currentDateStr = `${day}.${month}.${year}`;

    const wb = XLSX.readFile(filePath, { cellDates: true });
    if (!wb || !wb.SheetNames || !wb.SheetNames.length) return;
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws["!ref"]) return;

    const range = XLSX.utils.decode_range(ws["!ref"]);
    const targetColIndexes = new Set([91, 93, 95, 122]); // Default 0-based for CN (91), CP (93), CR (95), DS (122)

    for (let C = range.s.c; C <= range.e.c; ++C) {
      const headerCell = ws[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
      if (headerCell && headerCell.v) {
        const headerText = String(headerCell.v).trim().toLowerCase();
        if (
          headerText.includes("listed in the store") ||
          headerText.includes("listed in the distribution center") ||
          headerText.includes("sold in the store") ||
          headerText.includes("sold in the distribution center") ||
          headerText === "cn" || headerText === "cp" || headerText === "cr" || headerText === "ds"
        ) {
          targetColIndexes.add(C);
        }
      }
    }

    targetColIndexes.forEach((C) => {
      for (let R = Math.max(range.s.r + 1, 3); R <= range.e.r; ++R) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        ws[cellRef] = { t: "s", v: currentDateStr };
      }
    });

    XLSX.writeFile(wb, filePath);
  } catch (err) {
    console.error("Error updating ZKNG file dates:", err);
  }
}

function sourceMaterialValue(row, headings) {
  if (!row) return "";
  if (!row._lookup) {
    const lookup = new Map();
    for (const [key, value] of Object.entries(row)) {
      if (key !== "_lookup") {
        lookup.set(normalizedMaterialHeader(key), value);
      }
    }
    Object.defineProperty(row, "_lookup", { value: lookup, enumerable: false, writable: true });
  }
  const lookup = row._lookup;
  for (let i = 0; i < headings.length; i += 1) {
    const value = lookup.get(normalizedMaterialHeader(headings[i]));
    if (value !== undefined && value !== null && String(value).trim() !== "") return value;
  }
  return "";
}

function escapeSpreadsheetXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function replaceSpreadsheetCell(rowXml, ref, value) {
  const cellPattern = new RegExp(`<c\\b([^>]*\\br="${ref}"[^>]*)(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const match = rowXml.match(cellPattern);
  let replacement;
  if (match) {
    const styleMatch = match[1].match(/\bs="([^"]+)"/);
    const style = styleMatch ? ` s="${styleMatch[1]}"` : "";
    if (value === "" || value === null || value === undefined) {
      replacement = `<c r="${ref}"${style}/>`;
    } else if (typeof value === "number" && Number.isFinite(value)) {
      replacement = `<c r="${ref}"${style}><v>${value}</v></c>`;
    } else {
      const text = String(value);
      const preserveSpace = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
      replacement = `<c r="${ref}"${style} t="inlineStr"><is><t${preserveSpace}>${escapeSpreadsheetXml(text)}</t></is></c>`;
    }
    return rowXml.replace(cellPattern, replacement);
  } else {
    if (value === "" || value === null || value === undefined) return rowXml;
    if (typeof value === "number" && Number.isFinite(value)) {
      replacement = `<c r="${ref}"><v>${value}</v></c>`;
    } else {
      const text = String(value);
      const preserveSpace = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
      replacement = `<c r="${ref}" t="inlineStr"><is><t${preserveSpace}>${escapeSpreadsheetXml(text)}</t></is></c>`;
    }
    return rowXml.replace(/<\/row>/, `${replacement}</row>`);
  }
}

async function generateZkngReadyWorkbook(templatePath, materialRows, outputPath) {
  let sourceRows = materialRows.filter((row) => String(sourceMaterialValue(row, ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE", "CODE", "Code", "ITEM CODE", "Item Code", "MATERIAL"])).trim());
  if (!sourceRows.length) {
    sourceRows = materialRows.filter((row) => Object.values(row || {}).some((v) => v !== "" && v !== null && v !== undefined));
  }
  if (!sourceRows.length) throw new Error("No data rows found in the selected Material Code file.");

  const wb = XLSX.readFile(templatePath, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  const today = formatDDMMYYYY(new Date());
  const now = new Date();
  const excelDate = (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;

  // Extract all columns present in template row 4 to preserve all fixed SAP values (G..FF)
  const templateRow4Cells = {};
  Object.keys(ws).forEach((cellRef) => {
    const match = cellRef.match(/^([A-Z]{1,3})4$/);
    if (match) {
      const colLetter = match[1];
      templateRow4Cells[colLetter] = ws[cellRef];
    }
  });

  sourceRows.forEach((sourceRow, rowOffset) => {
    const r = 4 + rowOffset; // 1-indexed row 4, 5, 6...

    // First copy all template row 4 cells into row r so fixed columns (G, H, I, J, K... FF) match 100%
    Object.entries(templateRow4Cells).forEach(([colLetter, cellObj]) => {
      ws[`${colLetter}${r}`] = { ...cellObj };
    });

    const materialNumberRaw = sourceMaterialValue(sourceRow, ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"]);
    const description = sourceMaterialValue(sourceRow, ["DESCRIPTION", "Material Description", "MAKTX"]);
    const groupRaw = sourceMaterialValue(sourceRow, ["GROUP", "GRUIP", "Material Group", "MATKL"]);
    const materialTypeRaw = sourceMaterialValue(sourceRow, ["MATERIAL TYPE", "Material Type", "MTART"]);
    const baseUnitRaw = sourceMaterialValue(sourceRow, ["BASE UNIT OF MEASURE", "Base Unit of Measure", "MEINS", "UOM", "UNIT"]);
    const netWeightRaw = sourceMaterialValue(sourceRow, ["NET WEIGHT", "Net Weight", "NTGEW", "WEIGHT", "Weight"]);

    const materialNumberText = String(materialNumberRaw).trim();
    const materialNumber = /^\d+$/.test(materialNumberText) && Number.isSafeInteger(Number(materialNumberText))
      ? Number(materialNumberText)
      : materialNumberText;
    const groupText = String(groupRaw).trim();
    const group = /^\d+$/.test(groupText) ? Number(groupText) : groupText;
    const groupNumber = Number(groupText);

    const materialType = "ZKNG";
    const baseUnit = String(baseUnitRaw || "").trim() || (groupNumber >= 160013 && groupNumber <= 160020 ? "ST" : "KG");
    const netWeightStr = String(netWeightRaw || "").trim() || "0";

    const productionStorage = groupNumber >= 160001 && groupNumber <= 160012
      ? "KP10"
      : groupNumber >= 160013 && groupNumber <= 160020
        ? "CC11"
        : "KP10";
    const descriptionText = String(description || "");
    const standardPrice = baseUnit === "ST" ? 10 : /^SD/i.test(descriptionText) ? 315 : /^YD/i.test(descriptionText) ? 375 : "";

    ws[`A${r}`] = { t: typeof materialNumber === "number" ? "n" : "s", v: materialNumber };
    ws[`B${r}`] = { t: "s", v: materialType };
    ws[`C${r}`] = { t: "s", v: descriptionText };
    ws[`D${r}`] = { t: "s", v: baseUnit };
    ws[`E${r}`] = { t: typeof group === "number" ? "n" : "s", v: group };
    ws[`F${r}`] = { t: "s", v: netWeightStr };
    ws[`AI${r}`] = { t: "s", v: productionStorage };
    ws[`BM${r}`] = { t: "n", v: standardPrice };
    ws[`CN${r}`] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
    ws[`CP${r}`] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
    ws[`CR${r}`] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
    ws[`DS${r}`] = { t: "n", v: excelDate, z: "dd.mm.yyyy" };
  });

  const maxRow = 3 + sourceRows.length;
  ws["!ref"] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 161, r: maxRow - 1 } });
  ws["!pageSetup"] = { orientation: "landscape", paperSize: 9, fitToWidth: 1, fitToHeight: 0 };

  await fsp.mkdir(path.dirname(outputPath), { recursive: true });
  XLSX.writeFile(wb, outputPath);
  return { rowCount: sourceRows.length, sheetName: sheetName || "Formet", range: `A1:FF${maxRow}`, date: today };
}

async function generateCharacteristicsReadyWorkbook(materialRows, outputPath) {
  let sourceRows = materialRows.filter((row) => String(sourceMaterialValue(row, ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE", "CODE", "Code", "ITEM CODE", "Item Code", "MATERIAL"])).trim());
  if (!sourceRows.length) {
    sourceRows = materialRows.filter((row) => Object.values(row || {}).some((v) => v !== "" && v !== null && v !== undefined));
  }
  if (!sourceRows.length) throw new Error("No data rows found in the selected Material Code file.");

  const headers = [
    "Material Number",
    "Plant",
    "Material Description",
    "Fabric Type",
    "Fabrication",
    "Design / Jacquard Design",
    "Fabric Composition",
    "GSM",
    "Finish Type",
    "Yarn Count & Process-1",
    "Fibre Type",
    "Special Treatment",
    "Color Code / Color",
    "Color Combo",
    "Repeat Size"
  ];

  const charRows = sourceRows.map((r) => {
    const matNumText = String(sourceMaterialValue(r, ["Material Number", "MATERIAL NUMBER", "Material Code", "MATERIAL CODE", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"])).trim();
    const matNum = /^\d+$/.test(matNumText) && Number.isSafeInteger(Number(matNumText)) ? Number(matNumText) : matNumText;
    const gsmRaw = sourceMaterialValue(r, ["GSM", "Gsm"]);
    const gsmVal = /^\d+(?:\.\d+)?$/.test(gsmRaw) ? Number(gsmRaw) : gsmRaw;
    return [
      matNum,
      8110,
      String(sourceMaterialValue(r, ["DESCRIPTION", "Material Description", "MAKTX"])).trim(),
      String(sourceMaterialValue(r, ["FABRIC TYPE", "Fabric Type"])).trim(),
      String(sourceMaterialValue(r, ["FABRICATION", "Fabrication"])).trim(),
      String(sourceMaterialValue(r, ["DESIGN/JACQUARD DESIGN", "DESIGN", "JACQUARD DESIGN"])).trim(),
      String(sourceMaterialValue(r, ["FABRIC COMPOSITION", "Fabric Composition", "COMPOSITION"])).trim(),
      gsmVal,
      String(sourceMaterialValue(r, ["FINISH TYPE", "Finish Type"])).trim(),
      String(sourceMaterialValue(r, ["YARN COUNT & PROCESS-1", "YARN COUNT&PROCESS", "Yarn Count & Process-1"])).trim(),
      String(sourceMaterialValue(r, ["FIBRE TYPE", "Fibre Type"])).trim(),
      String(sourceMaterialValue(r, ["SPECIAL TREATMENT", "Special Treatment"])).trim(),
      String(sourceMaterialValue(r, ["COLOR/COLOR CODE", "COLOR CODE / COLOR", "COLOR CODE", "COLOR"])).trim(),
      String(sourceMaterialValue(r, ["COLOR COMBO", "Color Combo"])).trim(),
      String(sourceMaterialValue(r, ["REPEAT SIZE", "REPET SIZE", "Repeat Size"])).trim()
    ];
  });

  const outSheet = XLSX.utils.aoa_to_sheet([headers, ...charRows]);
  outSheet["!autofilter"] = { ref: `A1:O${charRows.length + 1}` };
  outSheet["!cols"] = [14, 8, 30, 16, 16, 22, 20, 8, 14, 22, 14, 18, 18, 16, 14].map((w) => ({ wch: w }));

  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, outSheet, "Characteristics Upload Format");

  await fsp.mkdir(path.dirname(outputPath), { recursive: true });
  XLSX.writeFile(outWb, outputPath);
  return { rowCount: charRows.length };
}

async function getConfig() {
  return readJson(userFile("config.json"), {
    sharePath: DEFAULT_SHARE,
    autoUpdate: true,
    sources: {},
  });
}

async function saveConfig(next) {
  const current = await getConfig();
  const merged = { ...current, ...next };
  await writeJsonAtomic(userFile("config.json"), merged);
  return merged;
}

async function getSavedAccountsLocally() {
  const data = await readJson(userFile("saved-accounts.json"), { savedAccounts: [], lastSelectedUsername: "" });
  return {
    savedAccounts: Array.isArray(data.savedAccounts) ? data.savedAccounts : [],
    lastSelectedUsername: String(data.lastSelectedUsername || ""),
  };
}

async function saveAccountLocally({ username, password, fullName, role }) {
  const normalized = String(username || "").trim().toLowerCase();
  if (!normalized) return { ok: false };
  const current = await getSavedAccountsLocally();
  const list = current.savedAccounts.filter((acc) => String(acc.username || "").toLowerCase() !== normalized);
  list.unshift({
    username: normalized,
    password: String(password || ""),
    fullName: fullName || normalized,
    role: role || "standard_user",
    savedAt: new Date().toISOString(),
  });
  const updated = {
    savedAccounts: list,
    lastSelectedUsername: normalized,
  };
  await writeJsonAtomic(userFile("saved-accounts.json"), updated);
  return { ok: true, data: updated };
}

async function removeSavedAccountLocally(username) {
  const normalized = String(username || "").trim().toLowerCase();
  const current = await getSavedAccountsLocally();
  const list = current.savedAccounts.filter((acc) => String(acc.username || "").toLowerCase() !== normalized);
  const lastSelected = current.lastSelectedUsername === normalized ? (list[0]?.username || "") : current.lastSelectedUsername;
  const updated = {
    savedAccounts: list,
    lastSelectedUsername: lastSelected,
  };
  await writeJsonAtomic(userFile("saved-accounts.json"), updated);
  return { ok: true, data: updated };
}

function normalizedUsername() {
  return String(os.userInfo().username || "").trim().toLowerCase();
}

function accessFile(sharePath) {
  return path.join(sharePath, "ZBOM-Admin", "access.json");
}

function hashPassword(password, salt = crypto.randomBytes(16).toString("hex")) {
  const hash = crypto.pbkdf2Sync(String(password), salt, 150000, 32, "sha256").toString("hex");
  return { salt, hash };
}

function passwordMatches(password, credential) {
  if (!credential?.salt || !credential?.hash) return false;
  const calculated = hashPassword(password, credential.salt).hash;
  const calculatedBuffer = Buffer.from(calculated, "hex");
  const storedBuffer = Buffer.from(credential.hash, "hex");
  return calculatedBuffer.length === storedBuffer.length
    && crypto.timingSafeEqual(calculatedBuffer, storedBuffer);
}

function publicSubAdmin(admin) {
  const { passwordCredential, ...safeAdmin } = admin;
  return { ...safeAdmin, role: normalizeRole(admin.role), passwordSet: Boolean(passwordCredential?.hash) };
}

let cachedAccessRegistry = null;
let cachedAccessRegistryTime = 0;

async function loadAccessRegistry(config, force = false) {
  const now = Date.now();
  if (!force && cachedAccessRegistry && (now - cachedAccessRegistryTime < 30000)) {
    return cachedAccessRegistry;
  }

  const fallback = await readJson(userFile("access-cache.json"), { departments: DEFAULT_DEPARTMENTS, subAdmins: [] });
  fallback.departments = fallback.departments?.length ? fallback.departments : DEFAULT_DEPARTMENTS;
  fallback.subAdmins ||= [];

  let registry = fallback;
  cachedAccessRegistry = registry;
  cachedAccessRegistryTime = now;

  const sharedPath = accessFile(config?.sharePath);
  if (sharedPath) {
    Promise.race([
      fsp.readFile(sharedPath, "utf8"),
      new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout")), 100)),
    ]).then((data) => {
      try {
        const remote = JSON.parse(data);
        if (remote && remote.departments) {
          cachedAccessRegistry = remote;
          writeJsonAtomic(userFile("access-cache.json"), remote).catch(() => {});
        }
      } catch {}
    }).catch(() => {});
  }

  registry.departments = registry.departments?.length ? registry.departments : DEFAULT_DEPARTMENTS;
  if (!registry.departmentSchemaVersion || registry.departmentSchemaVersion < 2) {
    registry.departments = registry.departments.map((name) => name === "Knitting" ? "Circular Knitting" : name);
    if (!registry.departments.includes("Flat Knitting")) registry.departments.splice(1, 0, "Flat Knitting");
    (registry.subAdmins || []).forEach((admin) => {
      if (admin.department === "Knitting") admin.department = "Circular Knitting";
    });
    registry.departmentSchemaVersion = 2;
  }
  registry.subAdmins ||= [];
  if (!registry.roleSchemaVersion || registry.roleSchemaVersion < 2) {
    registry.subAdmins.forEach((admin) => {
      admin.role = normalizeRole(admin.role);
    });
    registry.roleSchemaVersion = 2;
  }
  if (!registry.userIdLoginSchemaVersion || registry.userIdLoginSchemaVersion < 1) {
    registry.subAdmins.forEach((admin) => {
      if (admin.employeeId) admin.windowsUsername = String(admin.employeeId).trim().toLowerCase();
      admin.designation ||= "";
      admin.role = normalizeRole(admin.role || "power_user");
    });
    registry.userIdLoginSchemaVersion = 1;
  }

  cachedAccessRegistry = registry;
  cachedAccessRegistryTime = now;

  writeJsonAtomic(userFile("access-cache.json"), registry).catch(() => {});
  return registry;
}

async function getCurrentAccess(usernameOverride = "") {
  const config = await getConfig();
  const rawUsername = String(usernameOverride || normalizedUsername()).trim();
  const username = rawUsername.toLowerCase();
  const registry = await loadAccessRegistry(config);
  const isMainAdmin = MAIN_ADMIN_USERS.has(username) || username === "admin" || username === "main_admin" || username === "31001625" || username === "zakaria" || username === "zakaria-sflfu";
  const accountUsername = isMainAdmin ? "31001625" : (username || "guest");
  const subAdmin = (registry.subAdmins || []).find(
    (item) => (String(item.windowsUsername || "").trim().toLowerCase() === username || String(item.employeeId || "").trim().toLowerCase() === username) && item.active !== false,
  );
  const accountRole = normalizeRole(subAdmin?.role || "standard_user", isMainAdmin);
  const activity = isMainAdmin ? registry.mainAdminActivity || {} : subAdmin || {};
  return {
    username: accountUsername,
    role: accountRole,
    fullName: isMainAdmin ? "Zakaria Alam Mithu" : subAdmin?.fullName || username,
    employeeId: isMainAdmin ? "31001625" : subAdmin?.employeeId || "",
    designation: isMainAdmin ? "CTM" : subAdmin?.designation || "",
    lastLoginAt: activity.lastLoginAt || "",
    lastSeenAt: activity.lastSeenAt || "",
    totalUsageSeconds: Number(activity.totalUsageSeconds) || 0,
    loginCount: Number(activity.loginCount) || 0,
    departments: hasPermission({ role: accountRole }, "manage_all_data")
      ? registry.departments
      : subAdmin?.department
        ? [subAdmin.department]
        : [],
    allDepartments: registry.departments,
    passwordSet: isMainAdmin
      ? Boolean(registry.mainAdminPasswordCredential?.hash)
      : Boolean(subAdmin?.passwordCredential?.hash),
    permissions: [...(ROLE_PERMISSIONS[accountRole] || [])],
    filePermissions: Array.isArray(subAdmin?.filePermissions)
      ? subAdmin.filePermissions
      : ["materials", "codeRequirement", "finishType", "routingUploadFormat", "zkngMaterialUpload", "lsmwBomUpload", "characteristicsUpload", "summary", "diaWiseSummary", "department"],
    sessionTimeoutMinutes: SESSION_TIMEOUT_MS / 60000,
    subAdmins: hasPermission({ role: accountRole }, "user_management") ? (registry.subAdmins || []).map(publicSubAdmin) : [],
  };
}

async function verifyAdminPassword(usernameInput, password) {
  const username = String(usernameInput || "").trim().toLowerCase();
  if (!username) return { accountExists: false, active: false, valid: false };
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const access = await getCurrentAccess(username);
  const registeredAccount = (registry.subAdmins || []).find(
    (admin) => String(admin.windowsUsername || "").toLowerCase() === access.username
      || String(admin.employeeId || "").trim().toLowerCase() === access.username
      || String(admin.windowsUsername || "").toLowerCase() === String(username).toLowerCase()
      || String(admin.employeeId || "").trim().toLowerCase() === String(username).toLowerCase(),
  );
  const credential = access.role === "main_admin"
    ? registry.mainAdminPasswordCredential
    : registeredAccount?.passwordCredential;

  // Login requires both a known employee ID / username and its password.
  let isValid = Boolean(access && access.username) && passwordMatches(password, credential);

  return {
    access,
    accountExists: true,
    active: true,
    valid: isValid,
  };
}

function adminAuthorized(event, access) {
  const session = sessionForEvent(event);
  return Boolean(session) && session.username === access.username;
}

async function accessForEvent(event) {
  const session = sessionForEvent(event);
  if (session) return getCurrentAccess(session.username);
  const guest = await getCurrentAccess("__zbom_guest__");
  return {
    ...guest,
    username: "",
    role: "guest",
    fullName: "Guest User",
    employeeId: "",
    designation: "",
    departments: [],
    permissions: ["view"],
    passwordSet: false,
    subAdmins: [],
  };
}

async function saveSubAdmin(record, access) {
  access ||= await getCurrentAccess();
  if (!hasPermission(access, "user_management")) {
    return { ok: false, message: "Main Admin login is required to configure user accounts." };
  }
  const employeeId = String(record.employeeId || "").trim().toLowerCase();
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  if (!record.fullName || !employeeId || !record.department) {
    return { ok: false, message: "Full Name, Employee ID, and Department are required." };
  }
  if (!registry.departments.includes(record.department) && record.department !== "All Departments") {
    return { ok: false, message: "Select a valid department." };
  }
  const subAdmins = registry.subAdmins || [];
  const existingIndex = subAdmins.findIndex(
    (item) => String(item.windowsUsername || "").trim().toLowerCase() === employeeId
      || String(item.employeeId || "").trim().toLowerCase() === employeeId,
  );
  const existing = existingIndex >= 0 ? subAdmins[existingIndex] : null;
  const role = normalizeRole(record.role || existing?.role || "standard_user");
  const account = {
    windowsUsername: employeeId,
    fullName: String(record.fullName).trim(),
    employeeId,
    mobile: String(record.mobile || "").trim(),
    telephone: String(record.telephone || "").trim(),
    designation: String(record.designation || "").trim(),
    department: record.department,
    company: String(record.company || "SFL").trim(),
    role,
    active: existing ? existing.active !== false : true,
    passwordCredential: record.password ? hashPassword(record.password) : existing?.passwordCredential,
    filePermissions: Array.isArray(record.filePermissions)
      ? [...new Set(record.filePermissions.map((value) => String(value).trim()).filter(Boolean))]
      : (Array.isArray(existing?.filePermissions) ? existing.filePermissions : []),
    updatedAt: new Date().toISOString(),
    updatedBy: access.username,
  };
  if (existingIndex >= 0) subAdmins[existingIndex] = account;
  else subAdmins.push(account);

  const next = { ...registry, version: Date.now(), departments: registry.departments, subAdmins };
  await saveAccessRegistry(config, next);
  return {
    ok: true,
    message: `${record.fullName} was saved as ${roleLabel(role)}.`,
    subAdmins: subAdmins.map(publicSubAdmin),
  };
}

async function registerUserAccount(record) {
  const employeeId = String(record.employeeId || "").trim().toLowerCase();
  const department = String(record.department || "").trim();
  const password = String(record.password || "");
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  if (!record.fullName || !employeeId || !record.mobile || !record.designation
    || !record.company || !registry.departments.includes(department) || password.length < 6) {
    return { ok: false, message: "Complete all required fields and use a password of at least 6 characters." };
  }
  const duplicateActiveAccount = (registry.subAdmins || []).some((account) => {
    if (account.active === false) return false;
    const accountIds = [account.employeeId, account.windowsUsername]
      .map((value) => String(value || "").trim().toLowerCase())
      .filter(Boolean);
    return accountIds.includes(employeeId);
  });
  if (MAIN_ADMIN_USERS.has(employeeId) || duplicateActiveAccount) {
    return { ok: false, message: "An account already exists for this Employee ID." };
  }
  const now = new Date().toISOString();
  const account = {
    windowsUsername: employeeId,
    fullName: String(record.fullName).trim(),
    employeeId,
    mobile: String(record.mobile).trim(),
    telephone: String(record.telephone || "").trim(),
    designation: String(record.designation).trim(),
    department,
    company: String(record.company).trim(),
    role: "standard_user",
    active: true,
    passwordCredential: hashPassword(password),
    subAdminRequested: Boolean(record.requestSubAdmin),
    requestedDepartment: Boolean(record.requestSubAdmin) ? department : "",
    requestStatus: record.requestSubAdmin ? "pending" : "none",
    createdAt: now,
    updatedAt: now,
    updatedBy: employeeId,
  };
  registry.subAdmins ||= [];
  registry.subAdmins.push(account);
  registry.version = Date.now();
  try {
    await writeJsonAtomic(accessFile(config.sharePath), registry);
    await writeJsonAtomic(userFile("access-cache.json"), registry);
  } catch (error) {
    // The account lives on the LAN share: the ZBOM-Admin folder must be
    // created there first. A raw mkdir error (UNKNOWN/EPERM) only confuses,
    // so say plainly what is blocked and what to check.
    return {
      ok: false,
      message: `Could not save the account: the ZBOM-Admin folder could not be created on the LAN share (${config.sharePath}). Please check that this computer has write permission on the share.`,
    };
  }
  return {
    ok: true,
    username: employeeId,
    message: record.requestSubAdmin
      ? "User account created. Sub-Admin access is waiting for Main Admin approval."
      : "User account created successfully.",
  };
}

async function reviewSubAdminRequest(username, approved, access) {
  if (!hasPermission(access, "user_management")) {
    return { ok: false, message: "Admin permission is required to review access requests." };
  }
  const normalized = String(username || "").trim().toLowerCase();
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const account = (registry.subAdmins || []).find(
    (item) => String(item.windowsUsername || "").trim().toLowerCase() === normalized,
  );
  if (!account || account.requestStatus !== "pending") {
    return { ok: false, message: "No pending Sub-Admin request was found." };
  }
  account.requestStatus = approved ? "approved" : "rejected";
  account.requestReviewedAt = new Date().toISOString();
  account.requestReviewedBy = access.username;
  if (approved) {
    account.role = "power_user";
    account.department = account.requestedDepartment || account.department;
  } else {
    account.role = "standard_user";
  }
  await saveAccessRegistry(config, registry);
  return {
    ok: true,
    message: `${account.fullName}'s Sub-Admin request was ${approved ? "approved" : "rejected"}.`,
    subAdmins: registry.subAdmins.map(publicSubAdmin),
  };
}

async function manageSubAdminAccount(action, access) {
  access ||= await getCurrentAccess();
  if (!hasPermission(access, "user_management")) {
    return { ok: false, message: "Admin permission is required to control user accounts." };
  }
  const username = String(action.username || "").trim().toLowerCase();
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const index = (registry.subAdmins || []).findIndex(
    (item) => String(item.windowsUsername || "").trim().toLowerCase() === username
      || String(item.employeeId || "").trim().toLowerCase() === username,
  );
  if (index < 0) return { ok: false, message: "Select a valid Sub Admin account." };
  if (action.mode === "delete") {
    registry.subAdmins.splice(index, 1);
  } else if (action.mode === "toggle") {
    registry.subAdmins[index].active = action.active !== false;
    registry.subAdmins[index].updatedAt = new Date().toISOString();
    registry.subAdmins[index].updatedBy = access.username;
  } else {
    return { ok: false, message: "Invalid account action." };
  }
  await saveAccessRegistry(config, registry);
  return {
    ok: true,
    message: action.mode === "delete"
      ? `${username} was removed.`
      : `${username} is now ${action.active !== false ? "active" : "inactive"}.`,
    subAdmins: registry.subAdmins.map(publicSubAdmin),
  };
}

async function changeAccountRole(username, role, access) {
  access ||= await getCurrentAccess();
  if (!hasPermission(access, "user_management")) {
    return { ok: false, message: "Admin permission is required to change account roles." };
  }
  if (!ROLES.includes(role)) {
    return { ok: false, message: "Select a valid role." };
  }
  const normalized = String(username || "").trim().toLowerCase();
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const account = (registry.subAdmins || []).find(
    (item) => String(item.windowsUsername || "").trim().toLowerCase() === normalized
      || String(item.employeeId || "").trim().toLowerCase() === normalized,
  );
  if (!account) return { ok: false, message: "Account was not found." };
  account.role = role;
  if (["admin", "master_user"].includes(role)) {
    account.department = "All Departments";
  } else if (!registry.departments.includes(account.department)) {
    account.department = registry.departments[0];
  }
  account.updatedAt = new Date().toISOString();
  account.updatedBy = access.username;
  await saveAccessRegistry(config, registry);
  return {
    ok: true,
    message: `${account.fullName}'s role changed to ${roleLabel(role)}.`,
    subAdmins: registry.subAdmins.map(publicSubAdmin),
  };
}

async function setAdminPassword(targetUsername, password) {
  if (String(password || "").length < 6) {
    return { ok: false, message: "Password must contain at least 6 characters." };
  }
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const normalizedTarget = String(targetUsername || "").trim().toLowerCase();
  if (MAIN_ADMIN_USERS.has(normalizedTarget)) {
    registry.mainAdminPasswordCredential = hashPassword(password);
  } else {
    const admin = (registry.subAdmins || []).find((item) => [item.windowsUsername, item.employeeId]
      .map((value) => String(value || "").trim().toLowerCase())
      .includes(normalizedTarget));
    if (!admin) return { ok: false, message: "Select a valid Main Admin or Sub Admin." };
    admin.passwordCredential = hashPassword(password);
  }
  registry.version = Date.now();
  await writeJsonAtomic(accessFile(config.sharePath), registry);
  await writeJsonAtomic(userFile("access-cache.json"), registry);
  return {
    ok: true,
    message: MAIN_ADMIN_USERS.has(normalizedTarget)
      ? "Main Admin password updated."
      : "Sub Admin password reset successfully.",
  };
}

async function recordAccountActivity(username, { login = false, seconds = 0 } = {}) {
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const normalized = String(username || "").trim().toLowerCase();
  const now = new Date().toISOString();
  let activity;
  if (MAIN_ADMIN_USERS.has(normalized)) {
    registry.mainAdminActivity ||= {};
    activity = registry.mainAdminActivity;
  } else {
    activity = (registry.subAdmins || []).find(
      (account) => String(account.windowsUsername || "").trim().toLowerCase() === normalized,
    );
  }
  if (!activity) return;
  if (login) {
    activity.lastLoginAt = now;
    activity.loginCount = (Number(activity.loginCount) || 0) + 1;
  }
  activity.lastSeenAt = now;
  activity.totalUsageSeconds = (Number(activity.totalUsageSeconds) || 0) + Math.max(0, Math.round(seconds));
  registry.version = Date.now();
  try {
    await writeJsonAtomic(accessFile(config.sharePath), registry);
  } catch {
    // Activity remains recorded locally until LAN is available again.
  }
  await writeJsonAtomic(userFile("access-cache.json"), registry);
}

async function flushAccountActivity(event) {
  const session = accountActivitySessions.get(event.sender.id);
  if (!session) return;
  const now = Date.now();
  const seconds = Math.min(120, Math.max(0, (now - session.lastHeartbeat) / 1000));
  session.lastHeartbeat = now;
  await recordAccountActivity(session.username, { seconds });
}

function departmentSlug(department) {
  return department.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

async function loadDepartmentStatuses(config, departments) {
  const cached = await readJson(userFile("department-status-cache.json"), {});
  const statuses = { ...cached };
  let sharedFolderAvailable = false;
  const statusPromise = Promise.all((departments || []).map(async (department) => {
    const manifestPath = path.join(
      config.sharePath,
      "ZBOM-Departments",
      departmentSlug(department),
      "latest.json",
    );
    try {
      statuses[department] = JSON.parse(await fsp.readFile(manifestPath, "utf8"));
      sharedFolderAvailable = true;
    } catch {}
  }));
  await Promise.race([
    statusPromise,
    new Promise((resolve) => setTimeout(resolve, 150)),
  ]);
  if (sharedFolderAvailable) writeJsonAtomic(userFile("department-status-cache.json"), statuses).catch(() => {});
  return statuses;
}

async function loadDepartmentFiles(config, departments) {
  const cachedPath = userFile("department-files-cache.json");
  const cached = await readJson(cachedPath, []);
  const filesPromise = Promise.all((departments || []).map(async (department) => {
    const historyPath = path.join(
      config.sharePath,
      "ZBOM-Departments",
      departmentSlug(department),
      "files.json",
    );
    const history = await readJson(historyPath, []);
    return Array.isArray(history) ? history.map((item) => ({ ...item, department: item.department || department })) : [];
  })).then((histories) => histories.flat().sort((l, r) => Date.parse(r.uploadedAt || 0) - Date.parse(l.uploadedAt || 0)).slice(0, 500));

  const result = await Promise.race([
    filesPromise,
    new Promise((resolve) => setTimeout(() => resolve(null), 150)),
  ]);

  if (result) {
    writeJsonAtomic(cachedPath, result).catch(() => {});
    return result;
  }
  return cached;
}

async function manageDepartment(action, access) {
  access ||= await getCurrentAccess();
  if (!hasPermission(access, "system_control")) return { ok: false, message: "Admin permission is required to manage departments." };
  const config = await getConfig();
  const registry = await loadAccessRegistry(config);
  const departments = [...registry.departments];
  const newName = String(action.newName || "").trim();
  const oldName = String(action.oldName || "").trim();
  if (!newName) return { ok: false, message: "Enter a department name." };
  if (departments.some((name) => name.toLowerCase() === newName.toLowerCase() && name !== oldName)) {
    return { ok: false, message: "That department already exists." };
  }
  if (action.mode === "add") {
    departments.push(newName);
  } else if (action.mode === "rename") {
    const index = departments.indexOf(oldName);
    if (index < 0) return { ok: false, message: "Select a valid department to rename." };
    departments[index] = newName;
    (registry.subAdmins || []).forEach((admin) => {
      if (admin.department === oldName) admin.department = newName;
    });
    const oldFolder = path.join(config.sharePath, "ZBOM-Departments", departmentSlug(oldName));
    const newFolder = path.join(config.sharePath, "ZBOM-Departments", departmentSlug(newName));
    try {
      if (fs.existsSync(oldFolder) && !fs.existsSync(newFolder)) await fsp.rename(oldFolder, newFolder);
    } catch {
      // A locked legacy folder does not prevent the shared department registry from updating.
    }
  } else {
    return { ok: false, message: "Invalid department action." };
  }
  const next = { ...registry, version: Date.now(), departments };
  await writeJsonAtomic(accessFile(config.sharePath), next);
  await writeJsonAtomic(userFile("access-cache.json"), next);
  return {
    ok: true,
    message: action.mode === "add" ? `${newName} added.` : `${oldName} renamed to ${newName}.`,
    departments,
    subAdmins: (next.subAdmins || []).map(publicSubAdmin),
  };
}

async function uploadDepartmentFile(department, fileSlot, access) {
  access ||= await getCurrentAccess();
  if (!hasPermission(access, "upload")) {
    return { ok: false, message: "Your role does not have file-upload permission." };
  }
  if (!access.allDepartments.includes(department)) return { ok: false, message: "Invalid department." };
  if (!access.departments.includes(department)) {
    return { ok: false, message: `You do not have upload permission for ${department}.` };
  }
  const normalizedSlot = String(fileSlot || "").trim();
  if (!normalizedSlot) return { ok: false, message: "Enter the separate file name / slot before uploading." };
  const result = await dialog.showOpenDialog(mainWindow, {
    title: `Select ${normalizedSlot} file for ${department}`,
    properties: ["openFile"],
  });
  if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
  const config = await getConfig();
  const folder = path.join(config.sharePath, "ZBOM-Departments", departmentSlug(department));
  await fsp.mkdir(folder, { recursive: true });
  const historyFile = path.join(folder, "files.json");
  const history = await readJson(historyFile, []);
  const manifests = [];
  for (const source of result.filePaths) {
    const parsed = path.parse(path.basename(source));
    let destination = path.join(folder, path.basename(source));
    let copyNumber = 1;
    while (fs.existsSync(destination)) {
      destination = path.join(folder, `${parsed.name}-${copyNumber}${parsed.ext}`);
      copyNumber += 1;
    }
    await fsp.copyFile(source, destination);
    if (
      department === access.allDepartments[0]
      && (MATERIAL_PATTERN.test(parsed.base) || SUMMARY_PATTERN.test(parsed.base) || FINISH_TYPE_PATTERN.test(parsed.base))
    ) {
      await fsp.copyFile(source, path.join(config.sharePath, parsed.base));
    }
    if (normalizedSlot.toLowerCase().includes("lsmw bom")) {
      const lsmwMaster = userFile(path.join("lsmw-bom", "LSMW BOM Upload Master File.xlsx"));
      await fsp.mkdir(path.dirname(lsmwMaster), { recursive: true });
      await fsp.copyFile(source, lsmwMaster);
    }
    if (normalizedSlot.toLowerCase().includes("zkng")) {
      const zkngMaster = userFile(path.join("zkng", "ZKNG Master.xlsx"));
      await fsp.mkdir(path.dirname(zkngMaster), { recursive: true });
      await fsp.copyFile(source, zkngMaster);
    }
    if (normalizedSlot.toLowerCase().includes("routing")) {
      const routingMaster = userFile(path.join("routing", "Routing Master.xlsx"));
      await fsp.mkdir(path.dirname(routingMaster), { recursive: true });
      await fsp.copyFile(source, routingMaster);
    }
    const info = await fsp.stat(destination);
    manifests.push({
      department,
      fileSlot: normalizedSlot,
      file: path.basename(destination),
      size: info.size,
      uploadedAt: new Date().toISOString(),
      uploadedBy: access.fullName,
      windowsUsername: access.username,
      extension: path.extname(destination).toLowerCase(),
      category: /\.(xlsx|xls|csv)$/i.test(destination) ? "data" : "attachment",
    });
  }
  const manifest = manifests[manifests.length - 1];
  await writeJsonAtomic(path.join(folder, "latest.json"), manifest);
  await writeJsonAtomic(historyFile, [...manifests].reverse().concat(Array.isArray(history) ? history : []).slice(0, 500));
  return {
    ok: true,
    message: `${normalizedSlot}: ${manifest.file} uploaded separately for ${department}.`,
    manifest,
    manifests,
  };
}

async function profileImageDataUrl(username) {
  const config = await getConfig();
  const normalized = String(username || "").trim().toLowerCase();
  const profileFile = config.profileImageFiles?.[normalized]
    || (MAIN_ADMIN_USERS.has(normalized) ? config.profileImageFile : "");
  if (!profileFile) return "";
  try {
    const buffer = await Promise.race([
      fsp.readFile(profileFile),
      new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout")), 100)),
    ]);
    const extension = path.extname(profileFile).toLowerCase();
    const mime = extension === ".png" ? "image/png" : extension === ".webp" ? "image/webp" : "image/jpeg";
    return `data:${mime};base64,${buffer.toString("base64")}`;
  } catch {
    return "";
  }
}

async function chooseProfileImage(username) {
  const normalized = String(username || "").trim().toLowerCase();
  if (!normalized) return { ok: false, message: "Sign in before changing your profile photo." };
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Select your profile photo",
    properties: ["openFile"],
    filters: [{ name: "Images", extensions: ["png", "jpg", "jpeg", "webp"] }],
  });
  if (result.canceled || !result.filePaths[0]) return { ok: false, canceled: true };
  const source = result.filePaths[0];
  const info = await fsp.stat(source);
  if (info.size > 5 * 1024 * 1024) return { ok: false, message: "Profile photo must be 5 MB or smaller." };
  const safeUsername = normalized.replace(/[^a-z0-9_-]+/g, "-");
  const destination = userFile(path.join("profile-images", `${safeUsername}${path.extname(source).toLowerCase()}`));
  await fsp.mkdir(path.dirname(destination), { recursive: true });
  await fsp.copyFile(source, destination);
  const config = await getConfig();
  await saveConfig({
    profileImageFiles: { ...(config.profileImageFiles || {}), [normalized]: destination },
  });
  return { ok: true, imageDataUrl: await profileImageDataUrl(normalized), message: "Your profile photo was updated." };
}

async function downloadBundledResource(resourceId) {
  const codeRequirementOverride = userFile(path.join("master-files", "Code Requirement.xlsx"));
  const resources = {
    codeRequirement: {
      source: fs.existsSync(codeRequirementOverride)
        ? codeRequirementOverride
        : path.join(__dirname, "seed-data", "Code Requirement.xlsx"),
      downloadName: "Zbom-CodeRequirement",
    },
  };
  const resource = resources[resourceId];
  if (!resource || !fs.existsSync(resource.source)) {
    return { ok: false, message: "The requested resource file is not available." };
  }
  const confirmation = await dialog.showMessageBox(mainWindow, {
    type: "question",
    title: "New Code Requirement",
    message: "Download the Code Requirement file?",
    detail: "Use this Excel file to submit a new code requirement.",
    buttons: ["Download", "Cancel"],
    defaultId: 0,
    cancelId: 1,
    noLink: true,
  });
  if (confirmation.response !== 0) return { ok: false, canceled: true };
  const downloadsFolder = app.getPath("downloads");
  const config = await getConfig();
  const sourceHash = await sha256(resource.source);
  const savedDownload = config.resourceDownloads?.[resourceId];
  const counter = savedDownload?.sha256 === sourceHash ? Number(savedDownload.count) || 0 : 0;
  const extension = path.extname(resource.source);
  const destination = path.join(
    downloadsFolder,
    `${resource.downloadName}${counter || ""}${extension}`,
  );
  if (fs.existsSync(destination)) {
    const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
    let archiveCounter = 0;
    let archive;
    do {
      archive = path.join(
        downloadsFolder,
        `${resource.downloadName}-Previous-${timestamp}${archiveCounter || ""}${extension}`,
      );
      archiveCounter += 1;
    } while (fs.existsSync(archive));
    await fsp.copyFile(destination, archive);
  }
  await fsp.copyFile(resource.source, destination);
  await saveConfig({
    resourceDownloads: {
      ...(config.resourceDownloads || {}),
      [resourceId]: { sha256: sourceHash, count: counter + 1 },
    },
  });
  return { ok: true, path: destination, message: "Code Requirement file downloaded." };
}

function loadCodeRequirementPreview() {
  const override = userFile(path.join("master-files", "Code Requirement.xlsx"));
  const file = fs.existsSync(override)
    ? override
    : path.join(__dirname, "seed-data", "Code Requirement.xlsx");
  try {
    const workbook = XLSX.readFile(file, { cellDates: true, dense: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false,
      dateNF: "yyyy-mm-dd",
    });
    const columnCount = rows.length ? Object.keys(rows[0]).length : 0;
    const columnWidths = (sheet["!cols"] || []).slice(0, columnCount).map((column) =>
      Math.max(70, Math.min(260, Number(column?.wpx) || (Number(column?.wch) || 12) * 7)));
    return {
      ok: true,
      file: path.basename(file),
      sheetName,
      rows,
      columnWidths,
      worksheetRange: sheet["!ref"] || "",
    };
  } catch (error) {
    return { ok: false, file: "Code Requirement file.xlsx", sheetName: "", rows: [], message: error.message };
  }
}

function compareVersions(left, right) {
  const a = String(left).split(".").map((value) => Number(value) || 0);
  const b = String(right).split(".").map((value) => Number(value) || 0);
  for (let index = 0; index < Math.max(a.length, b.length); index += 1) {
    if ((a[index] || 0) > (b[index] || 0)) return 1;
    if ((a[index] || 0) < (b[index] || 0)) return -1;
  }
  return 0;
}

async function sha256(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash("sha256");
    const stream = fs.createReadStream(filePath);
    stream.on("error", reject);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
  });
}

async function checkSoftwareUpdate() {
  const config = await getConfig();
  const updateFolder = path.join(config.sharePath, "ZBOM-Updates");
  const manifestPath = path.join(updateFolder, "version.json");
  try {
    const manifest = JSON.parse(await fsp.readFile(manifestPath, "utf8"));
    const installer = path.join(updateFolder, path.basename(manifest.file || ""));
    if (!manifest.version || !manifest.file || !manifest.sha256 || !fs.existsSync(installer)) {
      return { ok: false, available: false, message: "Update manifest is incomplete." };
    }
    const available = compareVersions(manifest.version, app.getVersion()) > 0;
    return {
      ok: true,
      available,
      currentVersion: app.getVersion(),
      manifest,
      installer,
      message: available ? `ZBOM ${manifest.version} is available.` : "ZBOM is up to date.",
    };
  } catch {
    return {
      ok: false,
      available: false,
      currentVersion: app.getVersion(),
      message: "Software update folder is not available.",
    };
  }
}

async function installSoftwareUpdate() {
  const update = await checkSoftwareUpdate();
  if (!update.available) return update;

  const calculatedHash = await sha256(update.installer);
  if (calculatedHash.toLowerCase() !== update.manifest.sha256.toLowerCase()) {
    return { ok: false, available: true, message: "Update verification failed. The installer was not opened." };
  }

  const localInstaller = path.join(app.getPath("userData"), path.basename(update.manifest.file));
  await fsp.copyFile(update.installer, localInstaller);
  const localHash = await sha256(localInstaller);
  if (localHash.toLowerCase() !== update.manifest.sha256.toLowerCase()) {
    return { ok: false, available: true, message: "Downloaded update verification failed." };
  }

  const updaterScript = path.join(app.getPath("temp"), `zbom-update-${Date.now()}.cmd`);
  const installedExe = path.join(process.env.LOCALAPPDATA || "", "Programs", "ZBOM", "ZBOM.exe");
  const script = [
    "@echo off",
    "timeout /t 3 /nobreak >nul",
    `"${localInstaller}" /S`,
    `if exist "${installedExe}" start "" "${installedExe}"`,
    "del \"%~f0\"",
  ].join("\r\n");
  await fsp.writeFile(updaterScript, script, "utf8");
  spawn("cmd.exe", ["/c", updaterScript], { detached: true, stdio: "ignore", windowsHide: true }).unref();
  setTimeout(() => app.quit(), 500);
  return { ok: true, installing: true, message: "Installing verified update…" };
}

async function publishSoftwareUpdate(version, access) {
  access ||= await getCurrentAccess();
  if (access.role !== "main_admin") {
    return { ok: false, message: "Only the Main Admin can publish ZBOM software updates." };
  }
  if (!/^\d+\.\d+\.\d+$/.test(String(version))) {
    return { ok: false, message: "Version must look like 1.0.2." };
  }
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Select the approved ZBOM installer",
    properties: ["openFile"],
    filters: [{ name: "ZBOM installer", extensions: ["exe"] }],
  });
  if (result.canceled || !result.filePaths[0]) return { ok: false, canceled: true };

  const config = await getConfig();
  const updateFolder = path.join(config.sharePath, "ZBOM-Updates");
  await fsp.mkdir(updateFolder, { recursive: true });
  const filename = `ZBOM-Setup-${version}.exe`;
  const destination = path.join(updateFolder, filename);
  await fsp.copyFile(result.filePaths[0], destination);
  const manifest = {
    version,
    file: filename,
    sha256: await sha256(destination),
    publishedAt: new Date().toISOString(),
    publishedBy: "Administrator",
  };
  await writeJsonAtomic(path.join(updateFolder, "version.json"), manifest);
  return { ok: true, message: `ZBOM ${version} published for all users.`, manifest };
}

async function findSourceFiles(sharePath, options = {}) {
  async function collectFiles(folder, depth = 0) {
    const entries = await fsp.readdir(folder, { withFileTypes: true });
    const collected = [];
    for (const entry of entries) {
      const fullPath = path.join(folder, entry.name);
      if (entry.isFile() && /\.(xlsx|xls|csv)$/i.test(entry.name)) {
        collected.push({ name: entry.name, fullPath });
      } else if (entry.isDirectory() && depth < 1 && !entry.name.startsWith(".")) {
        collected.push(...await collectFiles(fullPath, depth + 1));
      }
    }
    return collected;
  }
  let files = [];
  try {
    const scanPromise = collectFiles(sharePath);
    files = await Promise.race([
      scanPromise,
      new Promise((_, reject) => setTimeout(() => reject(new Error("Network timeout")), 300)),
    ]);
  } catch {
    files = [];
  }

  const sources = {
    materials: files.find((file) => MATERIAL_PATTERN.test(file.name)) || null,
    summary: files.find((file) => /circular\s*knitting/i.test(file.name))
      || files.find((file) => SUMMARY_PATTERN.test(file.name))
      || null,
    finishType: files.find((file) => FINISH_TYPE_PATTERN.test(file.name)) || null,
  };
  const localOverrides = {
    materials: userFile(path.join("master-files", MASTER_SECTION_DEFAULT_FILE.materials)),
    summary: userFile(path.join("master-files", MASTER_SECTION_DEFAULT_FILE.summary)),
    finishType: userFile(path.join("master-files", MASTER_SECTION_DEFAULT_FILE.finishType)),
  };
  for (const [kind, localFile] of Object.entries(localOverrides)) {
    if (fs.existsSync(localFile)) {
      sources[kind] = { name: path.basename(localFile), fullPath: localFile };
    } else if (options.clearedMasterSections?.[kind]) {
      sources[kind] = null;
    }
  }
  if (!sources.finishType) {
    const bundled = path.join(__dirname, "seed-data", "Finsih Type Short Name.xlsx");
    if (fs.existsSync(bundled)) {
      sources.finishType = { name: path.basename(bundled), fullPath: bundled };
    }
  }
  return sources;
}

function makeUniqueHeaders(primary, secondary = []) {
  const used = new Map();
  return primary.map((value, index) => {
    const first = String(value ?? "").trim();
    const second = String(secondary[index] ?? "").trim();
    let header = first || second || `Column ${index + 1}`;
    const count = (used.get(header) || 0) + 1;
    used.set(header, count);
    if (count > 1) header = `${header} (${count})`;
    return header;
  });
}

function readSummaryWorkbook(workbook) {
  const sheetName = workbook.SheetNames.find((name) => /mastar\s*file/i.test(name))
    || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "dd.mm.yyyy",
  });
  const headerIndex = matrix.findIndex((row) =>
    row.some((value) => /buyer/i.test(String(value)))
    && row.some((value) => /knitted\s*qty/i.test(String(value))),
  );
  if (headerIndex < 0) return [];
  const headers = makeUniqueHeaders(matrix[headerIndex]);
  return matrix.slice(headerIndex + 1)
    .filter((row) => row.some((value) => String(value ?? "").trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])))
    .filter((row) => String(row["Order NO"] || "").trim() || String(row.Buyer || "").trim());
}

function readDiaWiseSummaryWorkbook(workbook) {
  const sheetName = workbook.SheetNames.find((name) => /dia\s*wise\s*summa/i.test(name));
  if (!sheetName) return [];
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "dd.mm.yyyy",
  });
  const headerIndex = matrix.findIndex((row) =>
    row.some((value) => /buyer/i.test(String(value)))
    && row.some((value) => /need\s*dia/i.test(String(value)))
    && row.some((value) => /knitting\s*balance/i.test(String(value))),
  );
  if (headerIndex < 0) return [];
  const headers = makeUniqueHeaders(matrix[headerIndex]);
  return matrix.slice(headerIndex + 1)
    .filter((row) => row.some((value) => String(value ?? "").trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])))
    .filter((row) => {
      const serial = String(row["Sl. NO"] || "").trim();
      return Boolean(serial) && Number.isFinite(Number(serial));
    });
}

function readDiaWiseTotalsWorkbook(workbook) {
  const sheetName = workbook.SheetNames.find((name) => /dia\s*wise\s*summa/i.test(name));
  if (!sheetName) return {};
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "dd.mm.yyyy",
  });
  const headerIndex = matrix.findIndex((row) =>
    row.some((value) => /buyer/i.test(String(value)))
    && row.some((value) => /need\s*dia/i.test(String(value))),
  );
  if (headerIndex < 0) return {};
  const headers = makeUniqueHeaders(matrix[headerIndex]);
  const totalRow = matrix.slice(headerIndex + 1).find((row) =>
    row.some((value) => /^total$/i.test(String(value).trim())));
  return totalRow
    ? Object.fromEntries(headers.map((header, index) => [header, totalRow[index] ?? ""]))
    : {};
}

const MATERIAL_COLUMN_ORDER = [
  "FINISH MATERIAL CODE",
  "Material Number",
  "DESCRIPTION",
  "MATERIAL TYPE",
  "GROUP",
  "FABRIC TYPE",
  "FABRICATION",
  "DESIGN/JACQUARD DESIGN",
  "FABRIC COMPOSITION",
  "GSM",
  "FINISH TYPE",
  "YARN COUNT & PROCESS-1",
  "FIBRE TYPE",
  "SPECIAL TREATMENT",
  "COLOR CODE / COLOR",
  "COLOR COMBO",
  "REPET SIZE",
];

function normalizedMaterialHeader(value) {
  return String(value || "")
    .replace(/^\ufeff/, "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "");
}

const MATERIAL_HEADER_ALIASES = {
  FINISHMATERIALCODE: ["FINISHEDMATERIALCODE", "FINISHMATCODE", "FINISHEDMATCODE"],
  MATERIALNUMBER: ["MATERIALNO", "MATERIALNUM", "MATERIALCODE"],
  MATERIALTYPE: ["MTYPE", "MATTYPE"],
  DESIGNJACQUARDDESIGN: ["DESIGNJACQUARD", "JACQUARDDESIGN"],
  FABRICCOMPOSITION: ["COMPOSITION", "FABRICCONTENT"],
  YARNCOUNTPROCESS1: ["YARNCOUNTPROCESS", "YARNCOUNTANDPROCESS1", "YARNCOUNTANDPROCESS"],
  FIBRETYPE: ["FIBERTYPE"],
  SPECIALTREATMENT: ["SPECIALTREATEMENT", "SPLTREATMENT"],
  COLORCODECOLOR: ["COLOURCODECOLOUR", "COLORCODE", "COLOURCODE"],
  COLORCOMBO: ["COLOURCOMBO", "COLORCOMBINATION"],
  REPETSIZE: ["REPEATSIZE", "REPEATERSIZE"],
};

function normalizeMaterialRows(rows) {
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    const sourceEntries = Object.entries(row || {});
    const sourceByHeader = new Map(sourceEntries.map(([key, value]) => [normalizedMaterialHeader(key), { key, value }]));
    const consumed = new Set();
    const normalized = {};
    for (const heading of MATERIAL_COLUMN_ORDER) {
      const canonical = normalizedMaterialHeader(heading);
      const candidates = [canonical, ...(MATERIAL_HEADER_ALIASES[canonical] || [])];
      const match = candidates.map((candidate) => sourceByHeader.get(candidate)).find(Boolean);
      normalized[heading] = match?.value ?? "";
      if (match) consumed.add(match.key);
    }
    for (const [key, value] of sourceEntries) {
      if (!consumed.has(key)) normalized[key] = value;
    }
    return normalized;
  });
}

function readWorkbook(filePath, kind = "") {
  const workbook = XLSX.readFile(filePath, { cellDates: true, dense: true, cellStyles: false, cellHTML: false, cellFormula: false });
  if (kind === "summary") return readSummaryWorkbook(workbook);
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, {
    defval: "",
    raw: false,
    dateNF: "dd.mm.yyyy",
  });
  return kind === "materials" ? normalizeMaterialRows(rows) : rows;
}

async function loadCache() {
  const snapshot = await readJson(userFile("snapshot.json"), {
    version: "0",
    updatedAt: null,
    materials: [],
    summary: [],
    diaWiseSummary: [],
    diaWiseTotals: {},
    finishType: [],
    sources: {},
  });
  snapshot.materials = normalizeMaterialRows(snapshot.materials || []);
  if (!snapshot.diaWiseSummary?.length || !Object.keys(snapshot.diaWiseTotals || {}).length) {
    const seedFile = path.join(__dirname, "seed-data", "Circular Knitting.xls");
    try {
      const workbook = XLSX.readFile(seedFile, { cellDates: true, dense: true });
      if (!snapshot.diaWiseSummary?.length) snapshot.diaWiseSummary = readDiaWiseSummaryWorkbook(workbook);
      if (!Object.keys(snapshot.diaWiseTotals || {}).length) snapshot.diaWiseTotals = readDiaWiseTotalsWorkbook(workbook);
    } catch {
      snapshot.diaWiseSummary ||= [];
    }
  }
  return snapshot;
}

async function updateFromShare(force = false) {
  const config = await getConfig();
  const previous = await loadCache();
  let masterSync = { changes: [], versions: config.sharedMasterVersions || {} };
  try {
    masterSync = await syncSharedMasterFiles(config, force);
  } catch {
    // Normal data updates can continue if the master manifest is temporarily unavailable.
  }
  let files;
  try {
    files = await findSourceFiles(config.sharePath, config);
  } catch (error) {
    return {
      ok: false,
      offline: true,
      message: `Shared folder is not available: ${config.sharePath}`,
      error: error.message,
      snapshot: previous,
    };
  }

  if (!files.materials && !files.summary && !files.finishType && !masterSync.changes.length) {
    return {
      ok: false,
      offline: false,
      message: "No configured ZBOM data file was found.",
      snapshot: previous,
    };
  }

  const next = { ...previous, sources: { ...previous.sources } };
  const changes = [...masterSync.changes];

  for (const [kind, file] of Object.entries(files)) {
    if (!file) continue;
    const info = await fsp.stat(file.fullPath);
    const fingerprint = `${info.size}:${info.mtimeMs}`;
    if (!force && previous.sources?.[kind]?.fingerprint === fingerprint) continue;
    next[kind] = readWorkbook(file.fullPath, kind);
    if (kind === "summary") {
      const workbook = XLSX.readFile(file.fullPath, { cellDates: true, dense: true });
      next.diaWiseSummary = readDiaWiseSummaryWorkbook(workbook);
      next.diaWiseTotals = readDiaWiseTotalsWorkbook(workbook);
    }
    next.sources[kind] = {
      filename: file.name,
      size: info.size,
      modifiedAt: info.mtime.toISOString(),
      fingerprint,
    };
    const diaSummaryNote = kind === "summary" && next.diaWiseSummary?.length
      ? `, ${next.diaWiseSummary.length.toLocaleString()} dia summary rows`
      : "";
    changes.push(`${file.name}: ${next[kind].length.toLocaleString()} records${diaSummaryNote}`);
  }

  if (!changes.length) {
    await saveConfig({ sharedMasterVersions: masterSync.versions });
    return { ok: true, changed: false, message: "Your offline data is already up to date.", snapshot: previous };
  }

  next.version = String(Date.now());
  next.updatedAt = new Date().toISOString();
  await writeJsonAtomic(userFile("snapshot.json"), next);
  await saveConfig({ sources: next.sources, sharedMasterVersions: masterSync.versions });
  return {
    ok: true,
    changed: true,
    message: `Update complete — ${changes.join(", ")}`,
    snapshot: next,
  };
}

const MASTER_CONTROL_ROLES = new Set(["power_user", "master_user", "department_head", "admin", "main_admin"]);

async function requireMasterControl(event) {
  const access = await accessForEvent(event);
  if (!MASTER_CONTROL_ROLES.has(access.role) && !(access.permissions || []).includes("upload")) {
    return { ok: false, message: "Upload permission is required to manage master files." };
  }
  return { ok: true, access };
}

const MASTER_SECTION_KIND = {
  materials: "materials",
  finishType: "finishType",
  summary: "summary",
  diaWiseSummary: "summary",
};
const MASTER_SECTION_DEFAULT_FILE = {
  materials: "Fabric Material Code.xlsx",
  finishType: "Finish Type Short Name.xlsx",
  summary: "Circular Knitting.xlsx",
  diaWiseSummary: "Circular Knitting.xlsx",
};
const MASTER_SECTION_LOCAL_FILE = {
  codeRequirement: ["master-files", "Code Requirement.xlsx"],
  routingUploadFormat: ["routing", "Routing Master.xlsx"],
  zkngMaterialUpload: ["zkng", "ZKNG Master.xlsx"],
  lsmwBomUpload: ["lsmw-bom", "LSMW BOM Upload Master File.xlsx"],
};
const SHARED_MASTER_FILES = {
  codeRequirement: { local: ["master-files", "Code Requirement.xlsx"], file: "Code Requirement.xlsx" },
  routingUploadFormat: { local: ["routing", "Routing Master.xlsx"], file: "Routing Master.xlsx" },
  zkngMaterialUpload: { local: ["zkng", "ZKNG Master.xlsx"], file: "ZKNG Master.xlsx" },
  lsmwBomUpload: { local: ["lsmw-bom", "LSMW BOM Upload Master File.xlsx"], file: "LSMW BOM Upload Master File.xlsx" },
};

async function publishSharedMasterFile(section, localPath) {
  const definition = SHARED_MASTER_FILES[section];
  if (!definition) return { ok: true };
  try {
    const config = await getConfig();
    const folder = path.join(config.sharePath, "ZBOM-Master-Files");
    await fsp.mkdir(folder, { recursive: true });
    const destination = path.join(folder, definition.file);
    if (path.resolve(localPath).toLowerCase() !== path.resolve(destination).toLowerCase()) await fsp.copyFile(localPath, destination);
    const manifestPath = path.join(folder, "manifest.json");
    const manifest = await readJson(manifestPath, { files: {} });
    const publishedAt = new Date().toISOString();
    manifest.files ||= {};
    manifest.files[section] = { file: definition.file, size: (await fsp.stat(destination)).size, publishedAt };
    manifest.updatedAt = publishedAt;
    await writeJsonAtomic(manifestPath, manifest);
    return { ok: true, publishedAt };
  } catch (error) {
    return { ok: false, message: `Master file is active locally, but could not be published to users: ${error.message}` };
  }
}

async function syncSharedMasterFiles(config, force = false) {
  const folder = path.join(config.sharePath, "ZBOM-Master-Files");
  const manifest = await readJson(path.join(folder, "manifest.json"), null);
  if (!manifest?.files) return { changes: [], versions: config.sharedMasterVersions || {} };
  const versions = { ...(config.sharedMasterVersions || {}) };
  const changes = [];
  for (const [section, definition] of Object.entries(SHARED_MASTER_FILES)) {
    const entry = manifest.files[section];
    if (!entry?.file || (!force && versions[section] === entry.publishedAt)) continue;
    const source = path.join(folder, path.basename(entry.file));
    if (!fs.existsSync(source)) continue;
    const destination = userFile(path.join(...definition.local));
    await fsp.mkdir(path.dirname(destination), { recursive: true });
    await fsp.copyFile(source, destination);
    versions[section] = entry.publishedAt || String((await fsp.stat(source)).mtimeMs);
    changes.push(`${definition.file} updated`);
  }
  return { changes, versions };
}

async function requireSecuredProcessor(event) {
  const session = sessionForEvent(event);
  if (!session) {
    return { ok: false, message: "Access Restricted. Please log in as Master User or Admin to access this module." };
  }
  const access = await getCurrentAccess(session.username);
  const allowedRoles = new Set(["master_user", "admin", "main_admin"]);
  if (!allowedRoles.has(access.role)) {
    return { ok: false, message: "Access Restricted. Please log in as Master User or Admin to access this module." };
  }
  return { ok: true, access };
}

async function masterSectionTarget(section) {
  const normalized = String(section || "").trim();
  const localParts = MASTER_SECTION_LOCAL_FILE[normalized];
  if (localParts) return { section: normalized, target: userFile(path.join(...localParts)), kind: "local" };
  const dataKind = MASTER_SECTION_KIND[normalized];
  if (!dataKind) return null;
  return {
    section: normalized,
    dataKind,
    kind: "shared",
    target: userFile(path.join("master-files", MASTER_SECTION_DEFAULT_FILE[normalized])),
  };
}

async function clearMasterSection(section) {
  const resolved = await masterSectionTarget(section);
  if (!resolved) return { ok: false, message: "This section does not have a master-file control." };
  if (fs.existsSync(resolved.target)) await fsp.unlink(resolved.target);
  let clearedSnapshot;
  if (resolved.kind === "shared") {
    const snapshot = await loadCache();
    if (resolved.dataKind === "summary") {
      snapshot.summary = [];
      snapshot.diaWiseSummary = [];
      snapshot.diaWiseTotals = {};
    } else {
      snapshot[resolved.dataKind] = [];
    }
    delete snapshot.sources?.[resolved.dataKind];
    snapshot.version = String(Date.now());
    snapshot.updatedAt = new Date().toISOString();
    await writeJsonAtomic(userFile("snapshot.json"), snapshot);
    const config = await getConfig();
    await saveConfig({
      clearedMasterSections: { ...(config.clearedMasterSections || {}), [resolved.dataKind]: true },
    });
    clearedSnapshot = snapshot;
  }
  return { ok: true, message: "Old master file and its cached data were cleared successfully.", snapshot: clearedSnapshot };
}

async function uploadMasterSection(section) {
  const resolved = await masterSectionTarget(section);
  if (!resolved) return { ok: false, message: "This section does not have a master-file control." };
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Select the new master file",
    properties: ["openFile"],
    filters: [{ name: "Excel or CSV files", extensions: ["xlsx", "xls", "csv"] }],
  });
  if (result.canceled || !result.filePaths[0]) return { ok: false, canceled: true };
  const source = result.filePaths[0];
  await fsp.mkdir(path.dirname(resolved.target), { recursive: true });
  if (path.resolve(source).toLowerCase() !== path.resolve(resolved.target).toLowerCase()) {
    await fsp.copyFile(source, resolved.target);
  }
  const published = await publishSharedMasterFile(resolved.section, resolved.target);
  if (resolved.kind === "shared") {
    const config = await getConfig();
    await saveConfig({
      clearedMasterSections: { ...(config.clearedMasterSections || {}), [resolved.dataKind]: false },
    });
  }
  const update = resolved.kind === "shared" ? await updateFromShare(true) : null;
  return {
    ok: update ? update.ok : true,
    message: update?.ok === false ? update.message : published.ok
      ? "New master file uploaded, activated and published to all users."
      : published.message,
    snapshot: update?.snapshot,
  };
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 650,
    backgroundColor: "#f5f6f7",
    title: "ZBOM",
    icon: path.join(__dirname, "renderer", "assets", "zcode.ico"),
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });
  mainWindow.setMenuBarVisibility(false);
  mainWindow.webContents.on("console-message", (_event, level, message, line, sourceId) => {
    try {
      const logLine = `[RENDERER ${level}] ${message} (${sourceId}:${line})\n`;
      fs.appendFileSync(userFile("renderer-debug.log"), logLine);
    } catch {}
  });
  mainWindow.loadFile(path.join(__dirname, "renderer", "index.html"), {
    query: isPreview ? { preview: "1" } : {},
  });
  mainWindow.once("ready-to-show", () => {
    mainWindow.show();
    mainWindow.focus();
    mainWindow.setAlwaysOnTop(true);
    mainWindow.setAlwaysOnTop(false);
  });

  if (isPreview) {
    let reloadTimer;
    previewWatcher = fs.watch(
      path.join(__dirname, "renderer"),
      { recursive: true },
      () => {
        clearTimeout(reloadTimer);
        reloadTimer = setTimeout(() => {
          if (mainWindow && !mainWindow.isDestroyed()) mainWindow.reload();
        }, 250);
      },
    );
  }
}

app.whenReady().then(async () => {
  // Every app launch starts in Guest mode. Saved credentials remain available
  // for an explicit sign-in, but must never create an automatic admin session.
  authenticatedAdminSessions.clear();
  accountActivitySessions.clear();
  ipcMain.handle("erp:get-state", async (event) => {
    const access = await accessForEvent(event);
    const config = await getConfig();
    const [snapshot, departmentStatuses, departmentFiles, profileImage, codeRequirementPreview] = await Promise.all([
      loadCache(),
      loadDepartmentStatuses(config, access.allDepartments),
      loadDepartmentFiles(config, access.departments),
      profileImageDataUrl(access.username),
      Promise.resolve().then(() => loadCodeRequirementPreview()),
    ]);
    return {
      config,
      snapshot,
      appVersion: app.getVersion(),
      access,
      departments: access.allDepartments,
      departmentStatuses,
      departmentFiles,
      authenticated: adminAuthorized(event, access),
      profileImageDataUrl: profileImage,
      codeRequirementPreview,
    };
  });
  ipcMain.handle("erp:get-access", async (event) => {
    const access = await accessForEvent(event);
    return {
      access,
      departments: access.allDepartments,
      departmentStatuses: await loadDepartmentStatuses(await getConfig(), access.allDepartments),
      departmentFiles: await loadDepartmentFiles(await getConfig(), access.departments),
      authenticated: adminAuthorized(event, access),
      profileImageDataUrl: await profileImageDataUrl(access.username),
    };
  });
  ipcMain.handle("erp:update", (_event, force) => updateFromShare(Boolean(force)));
  ipcMain.handle("erp:save-settings", async (_event, settings) => saveConfig(settings));
  ipcMain.handle("erp:choose-folder", async () => {
    const result = await dialog.showOpenDialog(mainWindow, { properties: ["openDirectory"] });
    return result.canceled ? null : result.filePaths[0];
  });
  ipcMain.handle("erp:export", async (event, { rows, format, defaultName, tableType, diaTotals, knittingBalance }) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) {
      return { ok: false, message: "Create an account or sign in before downloading files." };
    }
    if (!hasPermission(access, "export")) {
      return { ok: false, message: "Export permission is required to save CSV/Excel files." };
    }
    if (tableType === "materials" && rows.length > 50000) {
      return {
        ok: false,
        limitExceeded: true,
        message: "Fabric Material Code exports are limited to 50,000 records at a time. Please make the search or filter more specific.",
      };
    }
    const extension = format === "csv" ? "csv" : "xlsx";
    const safeBaseName = String(defaultName || "Zbom-Download").replace(/[<>:"/\\|?*\x00-\x1f]/g, "-");
    const downloadsFolder = app.getPath("downloads");
    let counter = 0;
    let filePath;
    do {
      const candidateBase = path.join(downloadsFolder, `${safeBaseName}${counter || ""}`);
      filePath = `${candidateBase}.${extension}`;
      counter += 1;
    } while (fs.existsSync(filePath) || fs.existsSync(filePath.replace(/\.(xlsx|csv)$/i, ".xlsx")) || fs.existsSync(filePath.replace(/\.(xlsx|csv)$/i, ".csv")));
    let sheet;
    if (tableType === "diaWiseSummary" && format !== "csv") {
      const columns = rows.length ? Object.keys(rows[0]) : [];
      const totalLines = (Array.isArray(diaTotals) ? diaTotals : []).reduce((lines, item, index) => {
        if (index % 6 === 0) lines.push([]);
        lines[lines.length - 1].push(`${item.label}: ${Math.round((Number(item.value) || 0) / 1000)} Ton`);
        return lines;
      }, []);
      const report = [["DIA WISE SUMMARY"], [`Generated: ${new Date().toLocaleString("en-GB")}`], [`Knitting Balance Qty: ${Math.round((Number(knittingBalance) || 0) / 1000)} Ton`], [], ...totalLines, [], columns, ...rows.map((row) => columns.map((column) => row[column] ?? ""))];
      sheet = XLSX.utils.aoa_to_sheet(report);
      const detailHeaderRow = 5 + totalLines.length;
      const endColumn = Math.max(0, columns.length - 1);
      const endRow = Math.max(detailHeaderRow, report.length - 1);
      sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: Math.min(endColumn, 5) } }];
      sheet["!cols"] = columns.map((column) => ({ wch: Math.min(28, Math.max(11, String(column).length + 3)) }));
      sheet["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
      sheet["!margins"] = { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
      sheet["!printArea"] = `A1:${XLSX.utils.encode_col(endColumn)}${endRow + 1}`;
      sheet["!autofilter"] = { ref: `A${detailHeaderRow + 1}:${XLSX.utils.encode_col(endColumn)}${endRow + 1}` };
    } else {
      sheet = XLSX.utils.json_to_sheet(rows);
    }
    if (format === "csv") {
      await fsp.writeFile(filePath, `\ufeff${XLSX.utils.sheet_to_csv(sheet)}`, "utf8");
    } else {
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, tableType === "diaWiseSummary" ? "Dia Wise Summary" : "Search Results");
      XLSX.writeFile(workbook, filePath);
    }
    return { ok: true, path: filePath };
  });
  ipcMain.handle("erp:open-path", (_event, target) => shell.openPath(target));
  ipcMain.handle("erp:get-remembered-login", async () => getSavedAccountsLocally());
  ipcMain.handle("erp:get-saved-accounts", async () => getSavedAccountsLocally());
  ipcMain.handle("erp:save-account-locally", async (_event, payload) => saveAccountLocally(payload));
  ipcMain.handle("erp:remove-saved-account", async (_event, username) => removeSavedAccountLocally(username));
  ipcMain.handle("erp:record-usage", async () => ({ ok: true }));
  ipcMain.handle("erp:get-characteristics-state", async (event) => {
    const access = await accessForEvent(event);
    const targetDir = userFile("characteristics");
    const sessionFile = path.join(targetDir, "session.json");
    const session = await readJson(sessionFile, {});
    const readyFileExists = session.readyFile ? fs.existsSync(path.join(targetDir, session.readyFile)) : false;
    const ready = Boolean(session.ready || readyFileExists);
    return {
      masterReady: true,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: ["master_user", "admin", "main_admin"].includes(access.role),
      materialFile: session.materialFile || null,
      generatedAt: session.generatedAt || null,
      importedRows: session.importedRows || 0,
      matchedRows: session.matchedRows || session.importedRows || 0,
      unmatchedRows: 0,
      readyFile: session.readyFile || null,
      ready,
    };
  });
  ipcMain.handle("erp:upload-characteristics-master", async (event) => {
    const access = await accessForEvent(event);
    if (!["master_user", "admin", "main_admin"].includes(access.role)) {
      return { ok: false, message: "Access Restricted. Please log in as Master User or Admin to access this module." };
    }
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select Characteristics Master File",
      properties: ["openFile"],
      filters: [{ name: "Excel files", extensions: ["xlsx", "xls"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const targetDir = userFile("characteristics");
    await fsp.mkdir(targetDir, { recursive: true });
    await fsp.copyFile(source, path.join(targetDir, "Characteristics Master.xlsx"));
    const published = await publishSharedMasterFile("characteristicsUpload", path.join(targetDir, "Characteristics Master.xlsx"));
    const sessionFile = path.join(targetDir, "session.json");
    const previous = await readJson(sessionFile, {});
    const next = { ...previous, masterFile: path.basename(source), masterUpdatedAt: new Date().toISOString() };
    await writeJsonAtomic(sessionFile, next);
    return { ok: published.ok, message: published.ok ? `Characteristics Master updated: ${path.basename(source)}` : published.message, state: { ...next, masterReady: true, canManageMaster: true, canImport: true, ready: Boolean(next.ready) } };
  });
  ipcMain.handle("erp:import-characteristics-material", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const access = authorization.access;
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select Material Code File for Characteristics",
      properties: ["openFile"],
      filters: [{ name: "Excel/CSV files", extensions: ["xlsx", "xls", "csv"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const matRows = readWorkbook(source);
    if (!Array.isArray(matRows) || !matRows.length) {
      return { ok: false, message: "No valid rows were found in the selected Material Code file." };
    }

    const now = new Date();
    const day = String(now.getDate()).padStart(2, "0");
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const year = now.getFullYear();
    const dateStr = `${day}.${month}.${year}`;

    const targetDir = userFile("characteristics");
    await fsp.mkdir(targetDir, { recursive: true });
    const readyFileName = `Characteristics Upload File - Ready-${dateStr}.xlsx`;
    const fullPath = path.join(targetDir, readyFileName);
    const generated = await generateCharacteristicsReadyWorkbook(matRows, fullPath);

    const sessionFile = path.join(targetDir, "session.json");
    const nextState = {
      masterReady: true,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: true,
      materialFile: path.basename(source),
      generatedAt: new Date().toISOString(),
      importedRows: generated.rowCount,
      matchedRows: generated.rowCount,
      unmatchedRows: 0,
      readyFile: readyFileName,
      ready: true,
    };
    await writeJsonAtomic(sessionFile, nextState);
    return { ok: true, message: `Characteristics processing complete: ${generated.rowCount} SAP-ready rows generated.`, state: nextState };
  });
  ipcMain.handle("erp:download-characteristics-ready", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const targetDir = userFile("characteristics");
    const session = await readJson(path.join(targetDir, "session.json"), {});
    if (!session.ready || !session.readyFile) return { ok: false, message: "No ready Characteristics file is available." };
    const source = path.join(targetDir, session.readyFile);
    if (!fs.existsSync(source)) return { ok: false, message: "Characteristics Ready file not found." };
    const downloadsFolder = app.getPath("downloads");
    const parsed = path.parse(session.readyFile);
    let destination = path.join(downloadsFolder, session.readyFile);
    let copyNumber = 1;
    while (fs.existsSync(destination)) {
      destination = path.join(downloadsFolder, `${parsed.name}-${copyNumber}${parsed.ext}`);
      copyNumber += 1;
    }
    await fsp.copyFile(source, destination);
    return { ok: true, path: destination, message: `Characteristics Upload File downloaded: ${path.basename(destination)}` };
  });
  ipcMain.handle("erp:clear-characteristics-session", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const targetDir = userFile("characteristics");
    const sessionFile = path.join(targetDir, "session.json");
    const previous = await readJson(sessionFile, {});
    if (previous.readyFile) {
      try { await fsp.unlink(path.join(targetDir, path.basename(previous.readyFile))); } catch {}
    }
    const nextState = {
      masterReady: true,
      canManageMaster: false,
      canImport: true,
      materialFile: null,
      generatedAt: null,
      importedRows: 0,
      matchedRows: 0,
      unmatchedRows: 0,
      readyFile: null,
      ready: false,
    };
    await writeJsonAtomic(sessionFile, nextState);
    return { ok: true, message: "Previous Material Code input and Characteristics Ready file were removed.", state: nextState };
  });
  ipcMain.handle("erp:get-routing-state", async (event) => {
    const access = await accessForEvent(event);
    const targetDir = userFile("routing");
    const session = await readJson(path.join(targetDir, "session.json"), {});
    const masterPath = path.join(targetDir, "Routing Master.xlsx");
    const builtInMaster = path.join(__dirname, "seed-data", "Routing Upload File Master.xlsx");
    return {
      masterReady: fs.existsSync(masterPath) || fs.existsSync(builtInMaster),
      masterFile: session.masterFile || (fs.existsSync(masterPath) ? "Routing Master.xlsx" : "Built-in Routing Master"),
      masterUpdatedAt: session.masterUpdatedAt || null,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: ["master_user", "admin", "main_admin"].includes(access.role),
      materialFile: session.materialFile || null,
      generatedAt: session.generatedAt || null,
      importedRows: session.importedRows || 0,
      matchedRows: session.matchedRows || 0,
      unmatchedRows: session.unmatchedRows || 0,
      unmatchedKeys: session.unmatchedKeys || [],
      readyFile: session.readyFile || null,
      ready: Boolean(session.ready),
    };
  });
  ipcMain.handle("erp:upload-routing-master", async (event) => {
    const access = await accessForEvent(event);
    if (!["master_user", "admin", "main_admin"].includes(access.role)) {
      return { ok: false, message: "Access Restricted. Please log in as Master User or Admin to access this module." };
    }
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select Routing Master File",
      properties: ["openFile"],
      filters: [{ name: "Excel files", extensions: ["xlsx", "xls"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const targetDir = userFile("routing");
    await fsp.mkdir(targetDir, { recursive: true });
    await fsp.copyFile(source, path.join(targetDir, "Routing Master.xlsx"));
    const published = await publishSharedMasterFile("routingUploadFormat", path.join(targetDir, "Routing Master.xlsx"));
    const sessionFile = path.join(targetDir, "session.json");
    const previous = await readJson(sessionFile, {});
    const next = { ...previous, masterFile: path.basename(source), masterUpdatedAt: new Date().toISOString() };
    await writeJsonAtomic(sessionFile, next);
    return { ok: published.ok, message: published.ok ? `Routing Master updated and published to all users: ${path.basename(source)}` : published.message, state: { ...next, masterReady: true, canManageMaster: true, canImport: true, ready: Boolean(next.ready) } };
  });
  ipcMain.handle("erp:import-routing-material", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const access = authorization.access;
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select Material Code File for Routing",
      properties: ["openFile"],
      filters: [{ name: "Excel/CSV files", extensions: ["xlsx", "xls", "csv"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const sourceRows = readWorkbook(source);
    if (!Array.isArray(sourceRows) || !sourceRows.length) return { ok: false, message: "No valid Material Code rows were found." };

    const targetDir = userFile("routing");
    const customMaster = path.join(targetDir, "Routing Master.xlsx");
    const seedMaster = path.join(__dirname, "seed-data", "Routing Upload File Master.xlsx");
    let masterPath = seedMaster;
    if (fs.existsSync(customMaster)) {
      try {
        const testWb = XLSX.readFile(customMaster, { cellFormula: true, cellDates: true });
        const testSheet = testWb.Sheets[testWb.SheetNames[0]];
        const testMatrix = XLSX.utils.sheet_to_json(testSheet, { header: 1, defval: "", raw: false });
        const headers = (testMatrix[1] || []).map((c) => String(c || "").toLowerCase());
        if (headers.some((h) => h.includes("work centre") || h.includes("control key") || h.includes("sutro link") || h.includes("operation description"))) {
          masterPath = customMaster;
        } else {
          try { fs.unlinkSync(customMaster); } catch {}
        }
      } catch {}
    }
    if (!fs.existsSync(masterPath)) return { ok: false, message: "Routing Master is not available. Admin must upload the fixed Routing format file." };

    const masterWorkbook = XLSX.readFile(masterPath, { cellFormula: true, cellDates: true });
    const masterSheet = masterWorkbook.Sheets[masterWorkbook.SheetNames[0]];
    const masterMatrix = XLSX.utils.sheet_to_json(masterSheet, { header: 1, defval: "", raw: false });
    const routingLookup = new Map();
    for (const masterRow of masterMatrix.slice(2)) {
      const materialType = String(masterRow[1] ?? "").trim();
      const group = String(masterRow[2] ?? "").trim();
      const fabricType = String(masterRow[3] ?? "").trim();
      const rawLink = `${materialType}${group}${fabricType}`;
      const link = rawLink.replace(/\s+/g, "").toLowerCase();
      if (!link || /^coming soon/i.test(link) || routingLookup.has(link)) continue;
      routingLookup.set(link, {
        unit: String(masterRow[4] ?? "").trim(),
        workCentre: String(masterRow[5] ?? "").trim(),
        controlKey: String(masterRow[6] ?? "").trim(),
        operationDescription: String(masterRow[7] ?? "").trim(),
        baseQuantity: String(masterRow[8] ?? "").trim(),
      });
    }

    const routingRows = [];
    const unmatchedKeys = [];
    for (const row of sourceRows) {
      const materialNumberText = String(sourceMaterialValue(row, ["MATERIAL NUMBER", "MATERIAL CODE", "Material Number", "Material Code", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"])).trim();
      if (!materialNumberText) continue;
      const materialType = String(sourceMaterialValue(row, ["MATERIAL TYPE", "Material Type"])).trim();
      const group = String(sourceMaterialValue(row, ["GROUP", "MATERIAL GROUP", "Material Group", "MATKL"])).trim();
      const fabricType = String(sourceMaterialValue(row, ["FABRIC TYPE", "Fabric Type"])).trim();
      const link = `${materialType}${group}${fabricType}`.replace(/\s+/g, "").toLowerCase();
      const matched = routingLookup.get(link);
      if (!matched) {
        unmatchedKeys.push({ materialNumber: materialNumberText, key: link || "Blank link" });
        continue;
      }
      const materialNumber = /^\d+$/.test(materialNumberText) && Number.isSafeInteger(Number(materialNumberText))
        ? Number(materialNumberText)
        : materialNumberText;
      const baseQuantity = /^[-+]?\d+(?:\.\d+)?$/.test(matched.baseQuantity)
        ? Number(matched.baseQuantity)
        : matched.baseQuantity;
      routingRows.push([
        materialNumber,
        8110,
        "",
        matched.unit,
        "01.01.2017",
        `x${String(routingRows.length + 1).padStart(7, "0")}`,
        matched.workCentre,
        matched.controlKey,
        matched.operationDescription,
        baseQuantity,
        "",
        "      ",
        24,
        "H",
      ]);
    }
    const importedRows = sourceRows.filter((row) => String(sourceMaterialValue(row, ["MATERIAL NUMBER", "MATERIAL CODE", "Material Number", "Material Code", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE"])).trim()).length;
    if (!importedRows) return { ok: false, message: "A Material Number or Material Code column is required." };

    const headers = ["Material Number", "Plant", "Material Description", "Material Base Unit of Measure", "Valid from", "Operation Number", "Work Centre", "Control Key", "Oeration Description", "Base Quantity", "Time 1", "UoM of Time1", "Time 2", "UoM of Time2"];
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...routingRows]);
    sheet["!cols"] = [10.21, 5.21, 8.93, 6.5, 9.93, 10.36, 11.79, 7.79, 13.93, 6.5, 3.93, 3.64, 6.5, 6.93].map((wch) => ({ wch }));
    sheet["!autofilter"] = { ref: "A1:N1" };
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Routing Upload Format");

    const now = new Date();
    const dateStr = `${String(now.getDate()).padStart(2, "0")}.${String(now.getMonth() + 1).padStart(2, "0")}.${now.getFullYear()}`;
    await fsp.mkdir(targetDir, { recursive: true });
    const readyFile = `Routing Upload File - Ready-${dateStr}.xlsx`;
    const sessionFile = path.join(targetDir, "session.json");
    const previous = await readJson(sessionFile, {});
    const next = {
      ...previous,
      masterFile: previous.masterFile || "Built-in Routing Master",
      canManageMaster: hasPermission(access, "system_control"),
      canImport: true,
      materialFile: path.basename(source),
      generatedAt: new Date().toISOString(),
      importedRows,
      matchedRows: routingRows.length,
      unmatchedRows: unmatchedKeys.length,
      unmatchedKeys,
      readyFile: unmatchedKeys.length ? null : readyFile,
      ready: unmatchedKeys.length === 0,
    };
    if (unmatchedKeys.length) {
      try { await fsp.unlink(path.join(targetDir, readyFile)); } catch {}
    } else {
      XLSX.writeFile(workbook, path.join(targetDir, readyFile));
    }
    await writeJsonAtomic(sessionFile, next);
    return {
      ok: unmatchedKeys.length === 0,
      message: unmatchedKeys.length
        ? `${unmatchedKeys.length} Routing link(s) were not found in the Master file. Admin must update the Routing Master.`
        : `Routing processing completed: ${routingRows.length} SAP-ready rows generated.`,
      state: { ...next, masterReady: true },
    };
  });
  ipcMain.handle("erp:download-routing-ready", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const targetDir = userFile("routing");
    const session = await readJson(path.join(targetDir, "session.json"), {});
    if (!session.ready || !session.readyFile) return { ok: false, message: "No Routing Ready file is available." };
    const source = path.join(targetDir, path.basename(session.readyFile));
    if (!fs.existsSync(source)) return { ok: false, message: "Routing Ready file was not found." };
    const downloadsFolder = app.getPath("downloads");
    const parsed = path.parse(session.readyFile);
    let destination = path.join(downloadsFolder, session.readyFile);
    let copyNumber = 1;
    while (fs.existsSync(destination)) {
      destination = path.join(downloadsFolder, `${parsed.name}-${copyNumber}${parsed.ext}`);
      copyNumber += 1;
    }
    await fsp.copyFile(source, destination);
    return { ok: true, path: destination, message: `Routing Upload File downloaded: ${path.basename(destination)}` };
  });
  ipcMain.handle("erp:clear-routing-session", async (event, kind = "routing") => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    if (kind !== "zkng") {
      const targetDir = userFile("routing");
      const sessionFile = path.join(targetDir, "session.json");
      const previous = await readJson(sessionFile, {});
      if (previous.readyFile) {
        try { await fsp.unlink(path.join(targetDir, path.basename(previous.readyFile))); } catch {}
      }
      const next = {
        masterReady: true,
        masterFile: previous.masterFile || "Built-in Routing Master",
        masterUpdatedAt: previous.masterUpdatedAt || null,
        canManageMaster: false,
        canImport: true,
        materialFile: null,
        generatedAt: null,
        importedRows: 0,
        matchedRows: 0,
        unmatchedRows: 0,
        unmatchedKeys: [],
        readyFile: null,
        ready: false,
      };
      await writeJsonAtomic(sessionFile, next);
      return { ok: true, message: "Previous Material Code input and Routing Ready file were removed.", state: next };
    }
    const targetDir = userFile("zkng");
    const sessionFile = path.join(targetDir, "session.json");
    const previous = await readJson(sessionFile, {});
    const generatedFiles = [previous.materialReadyFile, previous.routingReadyFile]
      .filter(Boolean)
      .map((file) => path.join(targetDir, path.basename(file)));
    await Promise.all(generatedFiles.map((file) => fsp.unlink(file).catch(() => {})));
    const nextState = {
      masterReady: true,
      zkngMasterReady: true,
      canManageMaster: false,
      canImport: true,
      masterFile: previous.masterFile || "ZKNG Master.xlsx",
      masterUpdatedAt: previous.masterUpdatedAt || null,
      materialFile: null,
      generatedAt: null,
      importedRows: 0,
      matchedRows: 0,
      unmatchedRows: 0,
      materialReadyFile: null,
      materialReady: false,
      ready: false,
    };
    await writeJsonAtomic(sessionFile, nextState);
    return { ok: true, message: "Previous Material Code input and ZKNG Ready file were removed.", state: nextState };
  });
  ipcMain.handle("erp:save-department-file", async () => ({ ok: true }));
  ipcMain.handle("erp:check-software-update", () => checkSoftwareUpdate());
  ipcMain.handle("erp:install-software-update", () => installSoftwareUpdate());
  ipcMain.handle("erp:reload-app", (event) => {
    event.sender.reload();
    return { ok: true };
  });
  ipcMain.handle("erp:admin-login", async (event, payload) => {
    const username = String(payload?.username || "").trim().toLowerCase();
    const result = await verifyAdminPassword(username, payload?.password);
    if (!result.accountExists || !result.active) return { ok: false, message: "This account is not active." };
    if (!result.valid) return { ok: false, message: "Incorrect username or password." };
    currentActiveSession = { username: result.access.username, lastActivityAt: Date.now() };
    authenticatedAdminSessions.set(event.sender.id, currentActiveSession);
    accountActivitySessions.set(event.sender.id, { username: result.access.username, lastHeartbeat: Date.now() });

    recordAccountActivity(result.access.username, { login: true }).catch((error) => {
      console.warn("Login activity could not be saved:", error.message);
    });

    if (payload?.remember !== false) {
      saveAccountLocally({
        username: result.access.username,
        password: payload.password,
        fullName: result.access.fullName,
        role: result.access.role,
      }).catch((error) => {
        console.warn("Remembered login could not be saved:", error.message);
      });
    }

    return { ok: true, access: result.access, message: `Signed in as ${result.access.fullName}.` };
  });
  ipcMain.handle("erp:register-account", async (_event, record) => registerUserAccount(record));
  ipcMain.handle("erp:logout", async (event) => {
    currentActiveSession = null;
    authenticatedAdminSessions.clear();
    accountActivitySessions.clear();
    const activityFlush = flushAccountActivity(event).catch((error) => {
      console.warn("Logout activity could not be flushed:", error.message);
    });
    void activityFlush;
    return {
      ok: true,
      authenticated: false,
      message: "Signed out successfully. You are now using Guest mode.",
    };
  });
  ipcMain.handle("erp:account-heartbeat", async (event) => {
    if (!authenticatedAdminSessions.has(event.sender.id)) return { ok: false };
    await flushAccountActivity(event);
    return { ok: true };
  });
  ipcMain.handle("erp:change-own-password", async (event, payload) => {
    const username = sessionForEvent(event)?.username;
    if (!username) return { ok: false, message: "Sign in before changing your password." };
    if (String(payload?.newPassword || "").length < 6) {
      return { ok: false, message: "New password must contain at least 6 characters." };
    }
    const verification = await verifyAdminPassword(username, payload?.currentPassword);
    if (!verification.valid) return { ok: false, message: "Current password is incorrect." };
    return setAdminPassword(username, payload.newPassword);
  });
  ipcMain.handle("erp:set-admin-password", async (event, payload) => {
    const access = await accessForEvent(event);
    if (!(access.role === "main_admin" || access.role === "admin") || !adminAuthorized(event, access)) {
      return { ok: false, message: "Admin authentication is required." };
    }
    const result = await setAdminPassword(payload.targetUsername, payload.password);
    if (result.ok) authenticatedAdminSessions.set(event.sender.id, { username: access.username, lastActivityAt: Date.now() });
    return result;
  });
  ipcMain.handle("erp:publish-software-update", async (event, version) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) return { ok: false, message: "Admin login is required." };
    return publishSoftwareUpdate(version, access);
  });
  ipcMain.handle("erp:save-sub-admin", async (event, record) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) return { ok: false, message: "Main Admin login is required." };
    return saveSubAdmin(record, access);
  });
  ipcMain.handle("erp:manage-sub-admin", async (event, action) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) return { ok: false, message: "Main Admin login is required." };
    return manageSubAdminAccount(action, access);
  });
  ipcMain.handle("erp:change-account-role", async (event, payload) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) return { ok: false, message: "Main Admin login is required." };
    return changeAccountRole(payload.username, payload.role, access);
  });
  ipcMain.handle("erp:review-sub-admin-request", async (event, payload) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) return { ok: false, message: "Main Admin login is required." };
    return reviewSubAdminRequest(payload.username, Boolean(payload.approved), access);
  });
  ipcMain.handle("erp:clear-master-section", async (event, section) => {
    const permission = await requireMasterControl(event);
    if (!permission.ok) return permission;
    return clearMasterSection(section);
  });
  ipcMain.handle("erp:upload-master-section", async (event, section) => {
    const permission = await requireMasterControl(event);
    if (!permission.ok) return permission;
    return uploadMasterSection(section);
  });
  ipcMain.handle("erp:upload-department-file", async (event, payload) => {
    const authorization = await requirePermission(event, "upload");
    if (!authorization.ok) return authorization;
    const department = typeof payload === "string" ? payload : payload?.department;
    const fileSlot = typeof payload === "string" ? "Department File" : payload?.fileSlot;
    return uploadDepartmentFile(department, fileSlot, authorization.access);
  });
  ipcMain.handle("erp:manage-department", async (event, action) => {
    const access = await accessForEvent(event);
    if (!adminAuthorized(event, access)) return { ok: false, message: "Main Admin login is required." };
    return manageDepartment(action, access);
  });
  ipcMain.handle("erp:choose-profile-image", async (event) => {
    const access = await accessForEvent(event);
    return chooseProfileImage(access.username);
  });
  ipcMain.handle("erp:download-resource", async (event, resourceId) => {
    const access = await accessForEvent(event);
    return downloadBundledResource(resourceId);
  });
  ipcMain.handle("erp:get-zkng-state", async (event) => {
    const access = await accessForEvent(event);
    const targetDir = userFile("zkng");
    const sessionFile = path.join(targetDir, "session.json");
    const session = await readJson(sessionFile, {});
    const readyFileExists = session.materialReadyFile ? fs.existsSync(path.join(targetDir, session.materialReadyFile)) : false;
    const materialReady = Boolean(session.materialReady || readyFileExists);
    return {
      masterReady: true,
      zkngMasterReady: true,
      zkngMasterFile: session.masterFile || "ZKNG Master.xlsx",
      zkngMasterUpdatedAt: session.masterUpdatedAt || null,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: ["master_user", "admin", "main_admin"].includes(access.role),
      materialFile: session.materialFile || null,
      generatedAt: session.generatedAt || null,
      importedRows: session.importedRows || 0,
      matchedRows: session.matchedRows || 0,
      unmatchedRows: session.unmatchedRows || 0,
      materialReadyFile: session.materialReadyFile || null,
      materialReady: materialReady,
      ready: materialReady,
    };
  });
  ipcMain.handle("erp:upload-zkng-master", async (event) => {
    const access = await accessForEvent(event);
    if (!["master_user", "admin", "main_admin"].includes(access.role)) return { ok: false, message: "Access Restricted. Please log in as Master User or Admin to access this module." };
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select ZKNG Master File",
      properties: ["openFile"],
      filters: [{ name: "Excel files", extensions: ["xlsx", "xls"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const destination = userFile(path.join("zkng", "ZKNG Master.xlsx"));
    await fsp.mkdir(path.dirname(destination), { recursive: true });
    await fsp.copyFile(source, destination);
    const published = await publishSharedMasterFile("zkngMaterialUpload", destination);
    const sessionFile = userFile(path.join("zkng", "session.json"));
    const previous = await readJson(sessionFile, {});
    const next = { ...previous, masterFile: path.basename(source), masterUpdatedAt: new Date().toISOString() };
    await writeJsonAtomic(sessionFile, next);
    return { ok: published.ok, message: published.ok ? `ZKNG Master replaced and published to all users: ${path.basename(source)}` : published.message, state: { ...next, masterReady: true, zkngMasterReady: true, canManageMaster: true, canImport: true, ready: Boolean(next.ready), materialReady: Boolean(next.materialReady) } };
  });
  ipcMain.handle("erp:import-zkng-material", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const access = authorization.access;
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select Material Code File for ZKNG",
      properties: ["openFile"],
      filters: [{ name: "Excel/CSV files", extensions: ["xlsx", "xls", "csv"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const matRows = readWorkbook(source);
    if (!Array.isArray(matRows) || !matRows.length) {
      return { ok: false, message: "No valid rows were found in the selected Material Code file." };
    }

    const dateStr = formatDDMMYYYY(new Date());
    const targetDir = userFile("zkng");
    await fsp.mkdir(targetDir, { recursive: true });
    const customMaster = path.join(targetDir, "ZKNG Master.xlsx");
    const builtInMaster = path.join(__dirname, "seed-data", "ZKNG Material Upload File.xlsx");
    const templatePath = fs.existsSync(customMaster) ? customMaster : (fs.existsSync(builtInMaster) ? builtInMaster : path.join(__dirname, "seed-data", "ZKNG Master.xlsx"));
    if (!fs.existsSync(templatePath)) {
      return { ok: false, message: "ZKNG master template is not available. Please upload ZKNG Material Upload File.xlsx first." };
    }
    const matXlsxName = `ZKNG Material Upload File - Ready-${dateStr}.xlsx`;
    const fullMatPath = path.join(targetDir, matXlsxName);
    const generated = await generateZkngReadyWorkbook(templatePath, matRows, fullMatPath);

    const sessionFile = path.join(targetDir, "session.json");
    const nextState = {
      masterReady: true,
      zkngMasterReady: true,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: true,
      materialFile: path.basename(source),
      generatedAt: new Date().toISOString(),
      importedRows: generated.rowCount,
      matchedRows: generated.rowCount,
      unmatchedRows: 0,
      materialReadyFile: matXlsxName,
      materialReady: true,
      ready: true,
    };
    await writeJsonAtomic(sessionFile, nextState);
    return { ok: true, message: `Processing completed: ${generated.rowCount} rows added to the ZKNG Material Upload Ready file.`, state: nextState };
  });
  ipcMain.handle("erp:download-zkng-ready", async (event, includeRouting = false) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const targetDir = userFile("zkng");
    const session = await readJson(path.join(targetDir, "session.json"), {});
    if (!session.materialReady || !session.materialReadyFile) {
      return { ok: false, message: "No ready ZKNG Material Upload file available." };
    }
    const downloadsFolder = app.getPath("downloads");
    const srcPath = path.join(targetDir, session.materialReadyFile);
    if (!fs.existsSync(srcPath)) return { ok: false, message: "ZKNG Material Upload file not found." };

    const now = new Date();
    const day = String(now.getDate()).padStart(2, "0");
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const year = now.getFullYear();
    const dateStr = `${day}.${month}.${year}`;

    const baseName = `ZKNG Material Upload File - Ready-${dateStr}`;
    let copyNumber = 0;
    let destFileName;
    let destPath;
    do {
      destFileName = `${baseName}${copyNumber ? `-${copyNumber}` : ""}.xlsx`;
      destPath = path.join(downloadsFolder, destFileName);
      copyNumber += 1;
    } while (fs.existsSync(destPath));

    await fsp.copyFile(srcPath, destPath);
    return {
      ok: true,
      path: destPath,
      paths: [destPath],
      message: `ZKNG Material Upload File downloaded: ${destFileName}`,
    };
  });
  ipcMain.handle("erp:get-lsmw-bom-state", async (event) => {
    const access = await accessForEvent(event);
    const sessionFile = userFile(path.join("lsmw-bom", "session.json"));
    const session = await readJson(sessionFile, {});
    return {
      masterReady: true,
      masterFile: session.masterFile || "LSMW BOM Upload Master File.xlsx",
      masterUpdatedAt: session.masterUpdatedAt || null,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: ["master_user", "admin", "main_admin"].includes(access.role),
      materialFile: session.materialFile || null,
      generatedAt: session.generatedAt || null,
      importedRows: session.importedRows || 0,
      headerRowsCount: session.headerRowsCount || 0,
      itemRowsCount: session.itemRowsCount || 0,
      generatedFiles: session.generatedFiles || [],
      ready: Boolean(session.ready),
    };
  });
  ipcMain.handle("erp:upload-lsmw-bom-master", async (event) => {
    const access = await accessForEvent(event);
    if (!["master_user", "admin", "main_admin"].includes(access.role)) return { ok: false, message: "Access Restricted. Please log in as Master User or Admin to access this module." };
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select LSMW BOM Master File",
      properties: ["openFile"],
      filters: [{ name: "Excel files", extensions: ["xlsx", "xls"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const destination = userFile(path.join("lsmw-bom", "LSMW BOM Upload Master File.xlsx"));
    await fsp.mkdir(path.dirname(destination), { recursive: true });
    await fsp.copyFile(source, destination);
    const published = await publishSharedMasterFile("lsmwBomUpload", destination);
    const sessionFile = userFile(path.join("lsmw-bom", "session.json"));
    const previous = await readJson(sessionFile, {});
    const next = { ...previous, masterFile: path.basename(source), masterUpdatedAt: new Date().toISOString() };
    await writeJsonAtomic(sessionFile, next);
    return { ok: published.ok, message: published.ok ? `LSMW BOM Master replaced and published to all users: ${path.basename(source)}` : published.message, state: { ...next, masterReady: true, canManageMaster: true, canImport: true, ready: Boolean(next.ready) } };
  });
  ipcMain.handle("erp:import-lsmw-bom-material", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const access = authorization.access;
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Select Material Code File for LSMW BOM",
      properties: ["openFile"],
      filters: [{ name: "Excel/CSV files", extensions: ["xlsx", "xls", "csv"] }],
    });
    if (result.canceled || !result.filePaths.length) return { ok: false, canceled: true };
    const source = result.filePaths[0];
    const matRows = readWorkbook(source);
    if (!Array.isArray(matRows) || !matRows.length) {
      return { ok: false, message: "No valid rows were found in the selected Material Code file." };
    }

    const now = new Date();
    const day = String(now.getDate()).padStart(2, "0");
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const year = now.getFullYear();
    const dateStr = `${day}.${month}.${year}`;

    const codeAliases = ["MATERIAL CODE", "Material Code", "Material Number", "MATERIAL NUMBER", "Material", "MATNR", "GREY MATERIAL", "Grey Material", "Fabric Code", "FABRIC CODE", "CODE", "Code", "ITEM CODE", "Item Code", "MATERIAL"];
    const yarnAliases = ["YARN CODE", "Yarn Code", "BOM component", "BOM Component", "Yarn", "YARN", "YARN COUNT", "Yarn Count", "COMPONENT", "Component"];
    const qtyAliases = ["YARN %", "Yarn %", "Quantity", "QUANTITY", "QTY", "Qty", "PERCENT", "Percent", "%"];

    function getRowMaterialCode(r) {
      const found = String(sourceMaterialValue(r, codeAliases)).trim();
      if (found) return found;
      const values = Object.values(r || {});
      for (const val of values) {
        const text = String(val || "").trim();
        if (text && text.toLowerCase() !== "null" && text.toLowerCase() !== "undefined") return text;
      }
      return "";
    }

    const uniqueHeaderMap = new Map();
    matRows.forEach((r) => {
      const code = getRowMaterialCode(r);
      if (!code) return;
      if (!uniqueHeaderMap.has(code)) {
        const matType = String(sourceMaterialValue(r, ["MATERIAL TYPE", "Material Type", "MTART"])).trim();
        const fabType = String(sourceMaterialValue(r, ["FABRIC TYPE", "Fabric Type"])).trim().toUpperCase();
        const materialGroup = Number(String(sourceMaterialValue(r, ["MATERIAL GROUP", "GROUP", "Material Group", "MATKL"])).trim());
        // The BOM master rules distinguish flat-knitting groups from circular
        // knitting groups.  Material Type is populated for some flat items,
        // but it must not override the circular group quantity rules.
        let baseQty = 99;
        if (materialGroup >= 160013 && materialGroup <= 160019) {
          baseQty = 1000;
        } else if (fabType === "YARN DYED") {
          baseQty = 94;
        } else if (materialGroup === 160005 || materialGroup === 160011) {
          baseQty = 99;
        } else {
          baseQty = 99.15;
        }
        uniqueHeaderMap.set(code, {
          "Material Number": isNaN(Number(code)) ? code : Number(code),
          Plant: 8110,
          "BOM Usage": 1,
          "BOM status": 1,
          "Base quantity": baseQty,
          "Valid-From Date": "01.01.2017",
        });
      }
    });
    const headerRows = Array.from(uniqueHeaderMap.values());

    const itemRows = matRows.map((r) => {
      const code = getRowMaterialCode(r);
      const yarnCode = String(sourceMaterialValue(r, yarnAliases)).trim();
      const quantityText = String(sourceMaterialValue(r, qtyAliases)).trim();
      const parsedQuantity = Number(quantityText.replace(/,/g, ""));
      const quantity = quantityText === "" ? "" : Number.isFinite(parsedQuantity) ? parsedQuantity : quantityText;
      if (!code) return null;
      return {
        "Material Number": code,
        Plant: 8110,
        " BOM component": yarnCode,
        "BOM Item Text (Line 1)": "",
        " Quantity": quantity,
        " Unit": "KG",
        " Item category": "L",
        " Scrap Percent": "",
      };
    }).filter(Boolean);

    const targetDir = userFile("lsmw-bom");
    await fsp.mkdir(targetDir, { recursive: true });

    const headXlsxName = `Header_BOM-${dateStr}.xlsx`;
    const headTxtName = `Header_BOM-${dateStr}.txt`;
    const itemXlsxName = `Item_BOM-${dateStr}.xlsx`;
    const itemTxtName = `Item_BOM-${dateStr}.txt`;

    // 1. Header_BOM.xlsx — same column order and display format as the approved reference.
    const headHeaders = ["Material Number", "Plant", "BOM Usage", "BOM status", "Base quantity", "Valid-From Date", "", "", ""];
    const headMatrix = [headHeaders, ...headerRows.map((r) => [
      r["Material Number"], r.Plant, r["BOM Usage"], r["BOM status"], r["Base quantity"], r["Valid-From Date"], "", "", "",
    ])];
    const headSheet = XLSX.utils.aoa_to_sheet(headMatrix);
    headSheet["!cols"] = [{ wch: 14 }, { wch: 9 }, { wch: 11 }, { wch: 11 }, { wch: 14 }, { wch: 15 }, { wch: 9 }, { wch: 9 }, { wch: 9 }];
    headSheet["!autofilter"] = { ref: `A1:F${headerRows.length + 1}` };
    for (let row = 2; row <= headerRows.length + 1; row += 1) headSheet[`E${row}`].z = "0.000";
    const headWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(headWb, headSheet, "Sheet1");
    XLSX.writeFile(headWb, path.join(targetDir, headXlsxName));

    // 2. Header_BOM.txt
    const headTxtLines = [`${headHeaders.slice(0, 6).join("\t")}\t\t\t`];
    headerRows.forEach((r) => {
      const bq = Number(r["Base quantity"]).toFixed(3);
      headTxtLines.push(`${[r["Material Number"], r.Plant, r["BOM Usage"], r["BOM status"], bq, r["Valid-From Date"]].join("\t")}\t\t\t`);
    });
    await fsp.writeFile(path.join(targetDir, headTxtName), `${headTxtLines.join("\r\n")}\r\n`, "utf8");

    // 3. Item_BOM.xlsx
    const itemHeaders = ["Material Number", "Plant", " BOM component", "BOM Item Text (Line 1)", " Quantity", " Unit", " Item category", " Scrap Percent"];
    const itemMatrix = [itemHeaders, ...itemRows.map((r) => [
      r["Material Number"], r.Plant, r[" BOM component"], "", r[" Quantity"], r[" Unit"], r[" Item category"], "",
    ])];
    const itemSheet = XLSX.utils.aoa_to_sheet(itemMatrix);
    itemSheet["!cols"] = [{ wch: 14 }, { wch: 13 }, { wch: 16 }, { wch: 19 }, { wch: 15 }, { wch: 11 }, { wch: 16 }, { wch: 14 }];
    itemSheet["!autofilter"] = { ref: `A1:H${itemRows.length + 1}` };
    for (let row = 2; row <= itemRows.length + 1; row += 1) {
      if (itemSheet[`E${row}`]?.t === "n") itemSheet[`E${row}`].z = "0.000";
    }
    const itemWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(itemWb, itemSheet, "Sheet1");
    XLSX.writeFile(itemWb, path.join(targetDir, itemXlsxName));

    // 4. Item_BOM.txt
    const itemTxtLines = [itemHeaders.join("\t")];
    itemRows.forEach((r) => {
      const quantity = typeof r[" Quantity"] === "number" ? r[" Quantity"].toFixed(3) : r[" Quantity"];
      itemTxtLines.push([r["Material Number"], r.Plant, r[" BOM component"], "", quantity, r[" Unit"], r[" Item category"], ""].join("\t"));
    });
    await fsp.writeFile(path.join(targetDir, itemTxtName), `${itemTxtLines.join("\r\n")}\r\n`, "utf8");

    const sessionFile = path.join(targetDir, "session.json");
    const nextState = {
      masterReady: true,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: true,
      materialFile: path.basename(source),
      generatedAt: new Date().toISOString(),
      importedRows: matRows.length,
      headerRowsCount: headerRows.length,
      itemRowsCount: itemRows.length,
      generatedFiles: [headXlsxName, headTxtName, itemXlsxName, itemTxtName],
      ready: true,
    };
    await writeJsonAtomic(sessionFile, nextState);
    return { ok: true, message: `LSMW BOM processing complete: Generated 4 ready files (${headerRows.length} Header & ${itemRows.length} Item rows).`, state: nextState };
  });
  ipcMain.handle("erp:download-lsmw-bom-ready", async (event, fileType = "all") => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const targetDir = userFile("lsmw-bom");
    const session = await readJson(path.join(targetDir, "session.json"), {});
    if (!session.ready || !Array.isArray(session.generatedFiles) || !session.generatedFiles.length) {
      return { ok: false, message: "No ready LSMW BOM files available." };
    }
    const downloadsFolder = app.getPath("downloads");
    let outputFolder = downloadsFolder;
    if (fileType === "all") {
      const now = new Date();
      const dateStr = `${String(now.getDate()).padStart(2, "0")}.${String(now.getMonth() + 1).padStart(2, "0")}.${now.getFullYear()}`;
      const baseFolderName = `BOM Ready File-${dateStr}`;
      outputFolder = path.join(downloadsFolder, baseFolderName);
      let folderNumber = 1;
      while (fs.existsSync(outputFolder)) {
        outputFolder = path.join(downloadsFolder, `${baseFolderName}-${folderNumber}`);
        folderNumber += 1;
      }
      await fsp.mkdir(outputFolder, { recursive: true });
    }
    const copiedFiles = [];
    for (const fileName of session.generatedFiles) {
      if (fileType !== "all") {
        const expected = fileType.toLowerCase();
        const extension = path.extname(expected);
        const stem = path.basename(expected, extension);
        if (!fileName.toLowerCase().startsWith(`${stem}-`) || path.extname(fileName).toLowerCase() !== extension) continue;
      }
      const srcPath = path.join(targetDir, fileName);
      const parsed = path.parse(fileName);
      let destPath = path.join(outputFolder, fileName);
      let copyNumber = 1;
      while (fs.existsSync(destPath)) {
        destPath = path.join(outputFolder, `${parsed.name}-${copyNumber}${parsed.ext}`);
        copyNumber += 1;
      }
      if (fs.existsSync(srcPath)) {
        await fsp.copyFile(srcPath, destPath);
        copiedFiles.push(destPath);
      }
    }
    if (!copiedFiles.length) return { ok: false, message: "Requested LSMW BOM file not found." };
    return {
      ok: true,
      path: fileType === "all" ? outputFolder : copiedFiles[0],
      paths: copiedFiles,
      message: fileType === "all"
        ? `Downloaded all 4 LSMW BOM files to ${path.basename(outputFolder)}.`
        : `Downloaded ${path.basename(copiedFiles[0])} to Downloads.`,
    };
  });
  ipcMain.handle("erp:clear-lsmw-bom-session", async (event) => {
    const authorization = await requireSecuredProcessor(event);
    if (!authorization.ok) return authorization;
    const access = await accessForEvent(event);
    const sessionFile = userFile(path.join("lsmw-bom", "session.json"));
    const previous = await readJson(sessionFile, {});
    const targetDir = userFile("lsmw-bom");
    for (const fileName of previous.generatedFiles || []) {
      const safeName = path.basename(String(fileName));
      if (safeName === fileName) {
        try { await fsp.unlink(path.join(targetDir, safeName)); } catch {}
      }
    }
    const nextState = {
      masterReady: true,
      canManageMaster: hasPermission(access, "system_control"),
      canImport: true,
      materialFile: null,
      generatedAt: null,
      importedRows: 0,
      headerRowsCount: 0,
      itemRowsCount: 0,
      generatedFiles: [],
      ready: false,
    };
    await writeJsonAtomic(sessionFile, nextState);
    return { ok: true, message: "LSMW BOM session cleared.", state: nextState };
  });
  createWindow();
});

app.on("window-all-closed", () => {
  previewWatcher?.close();
  if (process.platform !== "darwin") app.quit();
});
