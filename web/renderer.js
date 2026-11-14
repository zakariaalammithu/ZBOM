const state = {
  view: "dashboard",
  snapshot: { materials: [], summary: [], diaWiseSummary: [], diaWiseTotals: {}, finishType: [], sources: {}, updatedAt: null },
  config: {},
  access: { role: "guest", fullName: "Guest User", employeeId: "", departments: [], subAdmins: [], permissions: ["view"] },
  departments: [],
  departmentStatuses: {},
  departmentFiles: [],
  usageAnalytics: [],
  viewStartedAt: Date.now(),
  authenticated: false,
  profileImageDataUrl: "",
  profileImagesByUser: {},
  codeRequirementPreview: { rows: [], file: "", sheetName: "" },
  appVersion: "0.0.0",
  filteredMaterials: [],
  filteredSummary: [],
  filteredDiaWiseSummary: [],
  filteredFinishType: [],
  excelFilters: {
    materials: new Map(),
    summary: new Map(),
    finishType: new Map(),
  },
  summaryValueSelections: new Map(),
  materialValueSelections: new Map(),
  selectedRows: {
    materials: new Set(),
    summary: new Set(),
    finishType: new Set(),
  },
  currentPage: {
    materials: 1,
    summary: 1,
    finishType: 1,
  },
  lastDownloadedFile: "",
  columnWidths: (() => {
    try {
      return JSON.parse(localStorage.getItem("zbom-column-widths") || '{"materials":{},"summary":{}}');
    } catch {
      return { materials: {}, summary: {} };
    }
  })(),
  openExcelFilter: null,
  updating: false,
  summaryCategorySelected: new Set(),
  summaryOptionsReady: false,
  routing: {},
  zkng: {},
  lsmwBom: {},
  characteristics: {},
  pendingDepartmentUpload: null,
};

const SUMMARY_CATEGORIES = [
  "34x24", "32x24", "L/RIB", "28x24", "36x24", "RIB", "28gg", "30x24",
  "L/S/J", "YARN DYED", "INTERLOCK", "FLEECE", "40X24", "38X24", "20gg",
  "14gg", "22gg", "26x24", "24x24", "NEED DIA",
];
const DIA_BALANCE_COLUMNS = [
  "Interlock", "Rib", "L/Rib", "22x24", "23x24", "24x24", "26x24",
  "28x24", "30x24", "32x24", "34x24", "36x24", "38x24", "40x24", "14gg",
  "20gg", "22gg", "28gg", "Fleece", "L/S/J", "YARN DYED", "Need Dia",
];
state.summaryCategorySelected = new Set(SUMMARY_CATEGORIES);

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
state.columnWidths.materials ||= {};
state.columnWidths.summary ||= {};
state.columnWidths.finishType ||= {};
const previewMode = new URLSearchParams(window.location.search).get("preview") === "1";
if (previewMode) $("#previewBadge")?.classList.remove("hidden");

function showNotice(message, type = "info") {
  const element = $("#notice");
  element.textContent = message;
  element.className = `notice ${type}`;
  window.clearTimeout(showNotice.timer);
  showNotice.timer = window.setTimeout(() => element.classList.add("hidden"), 7000);
}

function formatDate(value) {
  if (!value) return "Never";
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");
  const seconds = String(d.getSeconds()).padStart(2, "0");
  if (d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0) {
    return `${day}.${month}.${year}`;
  }
  return `${day}.${month}.${year}, ${hours}:${minutes}:${seconds}`;
}

function formatFileSize(value) {
  const bytes = Number(value) || 0;
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function debounce(callback, wait = 180) {
  let timer;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => callback(...args), wait);
  };
}

function formatUsageTime(totalSeconds) {
  const seconds = Math.max(0, Number(totalSeconds) || 0);
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  if (hours) return `${hours}h ${minutes}m`;
  if (minutes) return `${minutes}m`;
  return seconds ? "< 1m" : "Not recorded";
}

function numberFrom(row, names) {
  for (const name of names) {
    if (row[name] !== undefined && row[name] !== "") {
      const value = Number(String(row[name]).replace(/,/g, ""));
      if (Number.isFinite(value)) return value;
    }
  }
  return 0;
}

function diaNumber(row, column) {
  const value = Number(String(row[column] ?? "").replace(/,/g, ""));
  return Number.isFinite(value) ? value : 0;
}

function formatTons(value) {
  const wholeTons = Math.round(value / 1000);
  return `${wholeTons} Ton`;
}

function dateFrom(row) {
  const value = row["Kniting Compolite Date"]
    || row["Knitting Complite Date"]
    || row["Order Sheet Received Date"]
    || row["Last Yarn Allocation Date "]
    || row["Plan Date"]
    || row.Date
    || row.DATE
    || row["Delivery Date"]
    || "";
  if (!value) return "";
  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? String(value).slice(0, 10) : parsed.toISOString().slice(0, 10);
}

function normalizedSearchText(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function searchValueMatches(value, query) {
  const text = normalizedSearchText(value);
  const term = normalizedSearchText(query);
  if (!term) return true;
  if (/^\d+(?:\.\d+)?$/.test(term)) {
    const compactNumber = text.replace(/,/g, "");
    if (/^\d+(?:\.\d+)?$/.test(compactNumber)) return compactNumber === term;
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return new RegExp(`(^|[^0-9.])${escaped}([^0-9.]|$)`).test(text);
  }
  return text.includes(term)
    || text.replace(/\s+/g, "").includes(term.replace(/\s+/g, ""));
}

function searchTerms(query) {
  const terms = [];
  const expression = /"([^"]+)"|(\S+)/g;
  let match;
  while ((match = expression.exec(String(query || ""))) !== null) {
    const term = normalizedSearchText(match[1] || match[2]);
    if (term) terms.push(term);
  }
  return terms;
}

function rowMatchesQuery(row, query) {
  const terms = searchTerms(query);
  if (!terms.length) return true;
  const values = Object.values(row);
  return terms.every((term) => values.some((value) => searchValueMatches(value, term)));
}

function valueMatchesAllTerms(value, query) {
  return searchTerms(query).every((term) => searchValueMatches(value, term));
}

function filterRows(rows, query, from = "", to = "") {
  return rows.filter((row) => {
    const textMatch = rowMatchesQuery(row, query);
    const date = dateFrom(row);
    const fromMatch = !from || !date || date >= from;
    const toMatch = !to || !date || date <= to;
    return textMatch && fromMatch && toMatch;
  });
}

function excelValue(row, column) {
  const value = String(row[column] ?? "").trim();
  return value || "(Blanks)";
}

function applyExcelFilters(rows, tableType) {
  const filters = state.excelFilters[tableType];
  if (!filters?.size) return rows;
  return rows.filter((row) =>
    [...filters.entries()].every(([column, excluded]) =>
      !excluded.has(excelValue(row, column)),
    ),
  );
}

function sourceRowsFor(tableType) {
  if (tableType === "materials") return state.snapshot.materials || [];
  if (tableType === "summary") return state.snapshot.summary || [];
  if (tableType === "finishType") return state.snapshot.finishType || [];
  return [];
}

function defaultColumnWidth(column) {
  const preferred = {
    "Short Name": 180,
    "Full Name": 420,
    Remarks: 360,
    "Grey Material": 120,
    DESCRIPTION: 220,
    "MATERIAL TYPE": 120,
    GROUP: 100,
    "FABRIC TYPE": 130,
    FABRICATION: 150,
    "DESIGN/JACQUARD DESIGN": 180,
    "FABRIC COMPOSITION": 160,
    GSM: 80,
    "FINISH TYPE": 110,
    "YARN COUNT & PROCESS-1": 190,
    "FIBRE TYPE": 120,
    "SPECIAL TREATMENT": 150,
    "COLOR CODE / COLOR": 150,
    "COLOR COMBO": 130,
    "REPET SIZE": 110,
  };
  return preferred[column] || Math.min(190, Math.max(100, column.length * 7 + 34));
}

function columnWidth(tableType, column) {
  const saved = state.columnWidths[tableType]?.[column];
  if (tableType === "finishType") {
    const readableMinimum = { "Short Name": 180, "Full Name": 420, Remarks: 360 }[column];
    if (readableMinimum) return Math.max(Number(saved) || 0, readableMinimum);
  }
  return saved || defaultColumnWidth(column);
}

const textMeasureCanvas = document.createElement("canvas");
const textMeasureContext = textMeasureCanvas.getContext("2d");

function shrinkToFitFontSize(text, width, baseSize = 9, reservedSpace = 18) {
  const available = Math.max(10, width - reservedSpace);
  textMeasureContext.font = `${baseSize}px "Segoe UI", Arial, sans-serif`;
  const measured = textMeasureContext.measureText(String(text || "")).width;
  if (measured <= available) return baseSize;
  return Math.max(6, Number((baseSize * available / measured).toFixed(2)));
}

function resizeCellText(cell, width) {
  const isHeader = cell.tagName === "TH";
  cell.style.fontSize = `${shrinkToFitFontSize(
    cell.dataset.fitText || cell.textContent,
    width,
    isHeader ? 9.5 : 9,
    isHeader ? 36 : 18,
  )}px`;
}

function saveColumnWidths() {
  localStorage.setItem("zbom-column-widths", JSON.stringify(state.columnWidths));
}

function updateFrozenColumns(table, tableType, columns) {
  if (tableType !== "materials") return;
  table.querySelectorAll(".select-column").forEach((cell) => {
    cell.classList.add("frozen-column", "frozen-select");
    cell.style.left = "0px";
  });
  let left = 38;
  columns.slice(0, 2).forEach((column, index) => {
    table.querySelectorAll(`[data-column-index="${index}"]`).forEach((cell) => {
      cell.classList.add("frozen-column");
      cell.classList.toggle("freeze-boundary", index === 1);
      cell.style.left = `${left}px`;
    });
    left += columnWidth(tableType, column);
  });
}

function startColumnResize(event, table, tableType, column, columnIndex) {
  event.preventDefault();
  event.stopPropagation();
  const startX = event.clientX;
  const startWidth = columnWidth(tableType, column);
  const col = table.querySelector(`col[data-column-index="${columnIndex}"]`);
  const cells = () => table.querySelectorAll(`[data-column-index="${columnIndex}"]`);

  const onMove = (moveEvent) => {
    const width = Math.max(55, Math.min(460, startWidth + moveEvent.clientX - startX));
    state.columnWidths[tableType][column] = width;
    if (col) col.style.width = `${width}px`;
    cells().forEach((cell) => {
      cell.style.width = `${width}px`;
      resizeCellText(cell, width);
    });
    const totalWidth = 38 + [...table.querySelectorAll("col[data-column-index]")]
      .reduce((sum, item) => sum + Number.parseFloat(item.style.width || "0"), 0);
    table.style.width = `${totalWidth}px`;
    const sourceRows = sourceRowsFor(tableType);
    const columns = sourceRows?.length ? Object.keys(sourceRows[0]) : [];
    updateFrozenColumns(table, tableType, columns);
  };

  const onUp = () => {
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("resizing-column");
    saveColumnWidths();
  };

  document.body.classList.add("resizing-column");
  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp);
}

function renderTable(table, rows, tableType) {
  const sourceRows = sourceRowsFor(tableType);
  const columns = sourceRows?.length ? Object.keys(sourceRows[0]) : [];
  const pageSize = 100;
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  state.currentPage[tableType] = Math.min(Math.max(1, state.currentPage[tableType]), totalPages);
  const pageStart = (state.currentPage[tableType] - 1) * pageSize;
  const visibleRows = rows.slice(pageStart, pageStart + pageSize);
  const selected = state.selectedRows[tableType];
  table.innerHTML = "";
  const colgroup = document.createElement("colgroup");
  const selectCol = document.createElement("col");
  selectCol.style.width = "38px";
  colgroup.appendChild(selectCol);
  columns.forEach((column, index) => {
    const col = document.createElement("col");
    col.dataset.columnIndex = String(index);
    col.style.width = `${columnWidth(tableType, column)}px`;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);
  table.style.width = `${38 + columns.reduce((sum, column) => sum + columnWidth(tableType, column), 0)}px`;
  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  const selectHead = document.createElement("th");
  selectHead.className = "select-column";
  const selectAll = document.createElement("input");
  selectAll.type = "checkbox";
  const selectedVisible = visibleRows.filter((row) => selected.has(row)).length;
  selectAll.checked = visibleRows.length > 0 && selectedVisible === visibleRows.length;
  selectAll.indeterminate = selectedVisible > 0 && selectedVisible < visibleRows.length;
  selectAll.title = "Select all visible rows";
  selectAll.addEventListener("change", () => {
    visibleRows.forEach((row) => {
      if (selectAll.checked) selected.add(row);
      else selected.delete(row);
    });
    renderTable(table, rows, tableType);
    updateExportLabels(tableType);
  });
  selectHead.appendChild(selectAll);
  headRow.appendChild(selectHead);
  columns.forEach((column, columnIndex) => {
    const cell = document.createElement("th");
    const width = columnWidth(tableType, column);
    cell.dataset.columnIndex = String(columnIndex);
    cell.dataset.fitText = column;
    cell.style.width = `${width}px`;
    resizeCellText(cell, width);
    const button = document.createElement("button");
    button.className = "table-filter-button";
    if (state.excelFilters[tableType]?.has(column)) button.classList.add("filtered");
    button.dataset.tableType = tableType;
    button.dataset.column = column;
    const label = document.createElement("span");
    label.textContent = column;
    const arrow = document.createElement("i");
    arrow.textContent = "▼";
    button.append(label, arrow);
    cell.appendChild(button);
    const resizer = document.createElement("span");
    resizer.className = "column-resizer";
    resizer.title = "Drag to resize column";
    resizer.addEventListener("mousedown", (event) =>
      startColumnResize(event, table, tableType, column, columnIndex),
    );
    cell.appendChild(resizer);
    headRow.appendChild(cell);
  });
  head.appendChild(headRow);
  table.appendChild(head);

  const body = document.createElement("tbody");
  for (const row of visibleRows) {
    const tr = document.createElement("tr");
    if (selected.has(row)) tr.classList.add("selected");
    const selectCell = document.createElement("td");
    selectCell.className = "select-column";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selected.has(row);
    checkbox.title = "Select this row";
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) selected.add(row);
      else selected.delete(row);
      tr.classList.toggle("selected", checkbox.checked);
      updateExportLabels(tableType);
    });
    selectCell.appendChild(checkbox);
    tr.appendChild(selectCell);
    columns.forEach((column, columnIndex) => {
      const cell = document.createElement("td");
      const width = columnWidth(tableType, column);
      cell.dataset.columnIndex = String(columnIndex);
      const cellText = String(row[column] ?? "");
      cell.dataset.fitText = cellText;
      cell.style.width = `${width}px`;
      resizeCellText(cell, width);
      cell.textContent = cellText;
      cell.title = cell.textContent;
      tr.appendChild(cell);
    });
    body.appendChild(tr);
  }
  table.appendChild(body);
  updateFrozenColumns(table, tableType, columns);
  renderPagination(tableType, rows.length, pageSize);
}

function renderPagination(tableType, totalRows, pageSize) {
  const page = state.currentPage[tableType];
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  const prefix = tableType === "materials" ? "material" : tableType;
  const container = $(`#${prefix}Pagination`);
  const meta = $(`#${prefix}PageMeta`);
  const firstRecord = totalRows ? (page - 1) * pageSize + 1 : 0;
  const lastRecord = Math.min(page * pageSize, totalRows);
  meta.textContent = `${firstRecord.toLocaleString()}–${lastRecord.toLocaleString()} of ${totalRows.toLocaleString()} · 100 per page`;
  container.innerHTML = "";

  const addButton = (label, targetPage, options = {}) => {
    const button = document.createElement("button");
    button.textContent = label;
    button.disabled = options.disabled;
    button.classList.toggle("active", options.active);
    button.addEventListener("click", () => {
      state.currentPage[tableType] = targetPage;
      if (tableType === "materials") renderMaterials();
      else if (tableType === "summary") renderSummary();
      else renderFinishType();
      const scroll = $(`#${prefix}Table`).closest(".table-scroll");
      if (scroll) scroll.scrollTop = 0;
    });
    container.appendChild(button);
  };

  addButton("« First Page", 1, { disabled: page <= 1 });
  addButton("‹ Previous", page - 1, { disabled: page <= 1 });
  const start = Math.max(1, Math.min(page - 3, totalPages - 6));
  const end = Math.min(totalPages, start + 6);
  for (let number = start; number <= end; number += 1) {
    addButton(String(number), number, { active: number === page });
  }
  addButton("Next ›", page + 1, { disabled: page >= totalPages });
  const summary = document.createElement("span");
  summary.textContent = `Page ${page.toLocaleString()} of ${totalPages.toLocaleString()}`;
  container.appendChild(summary);
}

function renderCodeRequirementPreview() {
  const preview = state.codeRequirementPreview || {};
  const rows = preview.rows || [];
  const table = $("#codeRequirementPreviewTable");
  table.innerHTML = "";
  if (!rows.length) {
    $("#codeRequirementPreviewInfo").textContent = preview.message || "Template preview is not available.";
    return;
  }
  const columns = Object.keys(rows[0]);
  const widths = columns.map((column, index) => {
    const workbookWidth = Number(preview.columnWidths?.[index]);
    if (workbookWidth > 0) return Math.max(80, Math.min(300, Math.round(workbookWidth * 1.35)));
    const longest = Math.max(
      column.length,
      ...rows.map((row) => String(row[column] ?? "").length),
    );
    return Math.min(250, Math.max(105, longest * 6 + 24));
  });
  const colgroup = document.createElement("colgroup");
  widths.forEach((width) => {
    const col = document.createElement("col");
    col.style.width = `${width}px`;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);
  table.style.width = `${widths.reduce((sum, width) => sum + width, 0)}px`;
  const head = document.createElement("thead");
  const headerRow = document.createElement("tr");
  columns.forEach((column, index) => {
    const cell = document.createElement("th");
    cell.textContent = column;
    cell.title = column;
    cell.dataset.fitText = column;
    resizeCellText(cell, widths[index]);
    headerRow.appendChild(cell);
  });
  head.appendChild(headerRow);
  table.appendChild(head);
  const body = document.createElement("tbody");
  rows.forEach((row) => {
    const tableRow = document.createElement("tr");
    columns.forEach((column, index) => {
      const cell = document.createElement("td");
      const value = String(row[column] ?? "");
      cell.textContent = value;
      cell.title = value;
      cell.dataset.fitText = value;
      resizeCellText(cell, widths[index]);
      tableRow.appendChild(cell);
    });
    body.appendChild(tableRow);
  });
  table.appendChild(body);
  $("#codeRequirementPreviewInfo").textContent =
    `${preview.file} · Sheet: ${preview.sheetName} · Range: ${preview.worksheetRange || "—"} · ${columns.length} columns · ${rows.length} data rows`;
}

function renderRoutingState() {
  const routing = state.routing || {};
  const serviceReady = routing.masterReady;
  $("#routingMasterStatus").textContent = serviceReady ? "Ready" : "Not Available";
  $("#routingMasterStatus").classList.toggle("not-ready", !serviceReady);
  $("#routingMasterDate").textContent = routing.masterUpdatedAt ? `Template updated ${formatDate(routing.masterUpdatedAt)}` : "Secured routing template";
  $("#routingAdminMaster").classList.toggle("hidden", !routing.canManageMaster);
  $("#routingMasterAdminInfo").textContent = routing.masterFile
    ? `${routing.masterFile} · Updated ${formatDate(routing.masterUpdatedAt)}`
    : "Only Admin can view or replace this secured backend workbook.";
  $("#importRoutingMaterial").disabled = !routing.canImport;
  $("#clearRoutingMaterial").disabled = false;
  $("#importRoutingMaterial").title = !routing.canImport ? "Access Restricted. Please log in as Master User or Admin." : "";
  $("#routingMaterialInfo").textContent = routing.materialFile
    ? `${routing.materialFile} · Generated ${formatDate(routing.generatedAt)}`
    : "No Material Code file selected.";
  const hasImport = Boolean(routing.materialFile);
  $("#routingResults").classList.toggle("hidden", !hasImport);
  $("#routingImportedRows").textContent = Number(routing.importedRows || 0).toLocaleString();
  $("#routingMatchedRows").textContent = Number(routing.matchedRows || 0).toLocaleString();
  $("#routingUnmatchedRows").textContent = Number(routing.unmatchedRows || 0).toLocaleString();
  const unmatched = routing.unmatchedKeys || [];
  $("#routingUnmatched").classList.toggle("hidden", !unmatched.length);
  $("#routingReviewMessage").textContent = routing.canManageMaster ? "Admin review required for the following routing references." : "Some routing rows need Admin review before download.";
  $("#routingUnmatchedList").classList.toggle("hidden", !routing.canManageMaster);
  $("#routingUnmatchedList").replaceChildren(...unmatched.map((item) => {
    const chip = document.createElement("span");
    chip.textContent = `${item.materialNumber}: ${item.key || "Blank link"}`;
    return chip;
  }));
  $("#downloadRoutingReady").disabled = !routing.ready;
  $("#routingReadyInfo").textContent = routing.ready
    ? `${Number(routing.importedRows || 0).toLocaleString()} structured Routing rows ready · Excel format`
    : routing.unmatchedRows
      ? `${Number(routing.unmatchedRows).toLocaleString()} rows need Admin review before the Ready file can be downloaded.`
      : "Process a Material Code file to prepare the Routing Upload Ready file.";
}

function renderZkngState() {
  const zkng = state.zkng || {};
  const serviceReady = zkng.masterReady !== false;
  if ($("#zkngServiceStatus")) {
    $("#zkngServiceStatus").textContent = serviceReady ? "Ready" : "Not Available";
    $("#zkngServiceStatus").classList.toggle("not-ready", !serviceReady);
  }
  if ($("#zkngServiceDate")) {
    $("#zkngServiceDate").textContent = zkng.zkngMasterUpdatedAt ? `Template updated ${formatDate(zkng.zkngMasterUpdatedAt)}` : "Secured SAP template";
  }
  const isFullAdmin = state.access.role === "admin" || state.access.role === "main_admin";
  if ($("#zkngAdminMaster")) $("#zkngAdminMaster").classList.toggle("hidden", !isFullAdmin);
  if ($("#zkngMasterAdminInfo")) $("#zkngMasterAdminInfo").textContent = zkng.zkngMasterFile ? `${zkng.zkngMasterFile} · Updated ${formatDate(zkng.zkngMasterUpdatedAt)}` : "ZKNG Material Master is not available.";
  if ($("#importZkngMaterial")) $("#importZkngMaterial").disabled = !zkng.canImport;
  if ($("#clearZkngMaterial")) $("#clearZkngMaterial").disabled = false;
  if ($("#zkngMaterialInfo")) $("#zkngMaterialInfo").textContent = zkng.materialFile ? `${zkng.materialFile} · Processed ${formatDate(zkng.generatedAt)}` : "No Material Code file selected.";
  if ($("#zkngResults")) $("#zkngResults").classList.remove("hidden");
  if ($("#zkngImportedRows")) $("#zkngImportedRows").textContent = Number(zkng.importedRows || 0).toLocaleString();
  if ($("#zkngMatchedRows")) $("#zkngMatchedRows").textContent = Number(zkng.matchedRows || 0).toLocaleString();
  if ($("#zkngUnmatchedRows")) $("#zkngUnmatchedRows").textContent = Number(zkng.unmatchedRows || 0).toLocaleString();
  if ($("#downloadZkngReady")) $("#downloadZkngReady").disabled = !zkng.materialReady;
  if ($("#downloadBothReady")) $("#downloadBothReady").disabled = !(zkng.materialReady && zkng.ready);
  if ($("#zkngReadyInfo")) $("#zkngReadyInfo").textContent = zkng.materialReady ? `${Number(zkng.importedRows || 0).toLocaleString()} fixed-value rows ready · Direct SAP upload format` : "Process a Material Code file to prepare the ZKNG Ready file.";
}

async function refreshZkngState() {
  try { state.zkng = await window.erp.getZkngState(); renderZkngState(); } catch (error) { showNotice(`ZKNG service unavailable: ${error.message}`, "warning"); }
}

function renderLsmwBomState() {
  const lsmwBom = state.lsmwBom || {};
  const serviceReady = lsmwBom.masterReady !== false;
  if ($("#lsmwBomServiceStatus")) {
    $("#lsmwBomServiceStatus").textContent = serviceReady ? "Ready" : "Not Available";
    $("#lsmwBomServiceStatus").classList.toggle("not-ready", !serviceReady);
  }
  if ($("#lsmwBomServiceDate")) {
    $("#lsmwBomServiceDate").textContent = lsmwBom.masterUpdatedAt ? `Rules updated ${formatDate(lsmwBom.masterUpdatedAt)}` : "Secured BOM rules";
  }
  if ($("#lsmwBomAdminMaster")) $("#lsmwBomAdminMaster").classList.toggle("hidden", !lsmwBom.canManageMaster);
  if ($("#lsmwBomMasterAdminInfo")) $("#lsmwBomMasterAdminInfo").textContent = lsmwBom.masterFile ? `${lsmwBom.masterFile} · Updated ${formatDate(lsmwBom.masterUpdatedAt)}` : "LSMW BOM Master is ready.";
  if ($("#importLsmwBomMaterial")) $("#importLsmwBomMaterial").disabled = !lsmwBom.canImport;
  if ($("#clearLsmwBomMaterial")) $("#clearLsmwBomMaterial").disabled = false;
  if ($("#lsmwBomMaterialInfo")) $("#lsmwBomMaterialInfo").textContent = lsmwBom.materialFile ? `${lsmwBom.materialFile} · Processed ${formatDate(lsmwBom.generatedAt)}` : "No Material Code file selected.";
  if ($("#lsmwBomResults")) $("#lsmwBomResults").classList.toggle("hidden", !lsmwBom.materialFile);
  if ($("#lsmwBomImportedRows")) $("#lsmwBomImportedRows").textContent = Number(lsmwBom.importedRows || 0).toLocaleString();
  if ($("#lsmwBomHeaderRows")) $("#lsmwBomHeaderRows").textContent = Number(lsmwBom.headerRowsCount || 0).toLocaleString();
  if ($("#lsmwBomItemRows")) $("#lsmwBomItemRows").textContent = Number(lsmwBom.itemRowsCount || 0).toLocaleString();
  const ready = Boolean(lsmwBom.ready);
  ["downloadAllLsmwBom", "downloadHeaderBomXlsx", "downloadHeaderBomTxt", "downloadItemBomXlsx", "downloadItemBomTxt"].forEach((id) => {
    const btn = $(`#${id}`);
    if (btn) btn.disabled = !ready;
  });
  if ($("#lsmwBomReadyInfo")) $("#lsmwBomReadyInfo").textContent = ready ? `Ready: ${Number(lsmwBom.headerRowsCount || 0).toLocaleString()} Header rows · ${Number(lsmwBom.itemRowsCount || 0).toLocaleString()} Item rows · 4 files` : "Process a Material Code file to prepare all four LSMW BOM files.";
}

async function refreshLsmwBomState() {
  try {
    if (typeof window.erp.getLsmwBomState === "function") {
      state.lsmwBom = await window.erp.getLsmwBomState();
    } else {
      state.lsmwBom = { masterReady: true, canManageMaster: state.access?.role === "admin" || state.access?.role === "main_admin" };
    }
    renderLsmwBomState();
  } catch (error) {
    showNotice(`LSMW BOM status could not load: ${error.message}`, "warning");
  }
}

async function refreshCharacteristicsState() {
  try {
    if (typeof window.erp.getCharacteristicsState === "function") {
      state.characteristics = await window.erp.getCharacteristicsState();
    } else {
      state.characteristics = { masterReady: false, canManageMaster: false, canImport: false, ready: false };
    }
    renderCharacteristicsState();
  } catch (error) {
    showNotice(`Characteristics status could not load: ${error.message}`, "warning");
  }
}

function renderCharacteristicsState() {
  const characteristics = state.characteristics || {};
  const serviceReady = characteristics.masterReady !== false;
  if ($("#characteristicsStatus")) {
    $("#characteristicsStatus").textContent = serviceReady ? "Ready" : "Not Available";
    $("#characteristicsStatus").classList.toggle("not-ready", !serviceReady);
  }
  const isFullAdmin = state.access.role === "admin" || state.access.role === "main_admin";
  if ($("#characteristicsAdminMaster")) $("#characteristicsAdminMaster").classList.toggle("hidden", !isFullAdmin);
  if ($("#importCharacteristicsMaterial")) $("#importCharacteristicsMaterial").disabled = !characteristics.canImport;
  if ($("#clearCharacteristicsMaterial")) $("#clearCharacteristicsMaterial").disabled = false;
  if ($("#characteristicsMaterialInfo")) $("#characteristicsMaterialInfo").textContent = characteristics.materialFile ? `${characteristics.materialFile} · Processed ${formatDate(characteristics.generatedAt)}` : "No Material Code file selected.";
  if ($("#characteristicsResults")) $("#characteristicsResults").classList.toggle("hidden", !characteristics.materialFile);
  if ($("#characteristicsImportedRows")) $("#characteristicsImportedRows").textContent = Number(characteristics.importedRows || 0).toLocaleString();
  if ($("#characteristicsMatchedRows")) $("#characteristicsMatchedRows").textContent = Number(characteristics.matchedRows || 0).toLocaleString();
  if ($("#downloadCharacteristicsReady")) $("#downloadCharacteristicsReady").disabled = !characteristics.ready;
  if ($("#characteristicsReadyInfo")) $("#characteristicsReadyInfo").textContent = characteristics.ready ? `${Number(characteristics.importedRows || 0).toLocaleString()} SAP-ready characteristics rows ready · Excel format` : "Process a Material Code file to prepare the Characteristics Upload Ready file.";
}

async function refreshRoutingState() {
  try {
    state.routing = await window.erp.getRoutingState();
    renderRoutingState();
  } catch (error) {
    showNotice(`Routing status could not load: ${error.message}`, "warning");
  }
}

function updateExportLabels(tableType) {
  const count = state.selectedRows[tableType].size;
  const section = `#${tableType}`;
  const hasFilter = tableType === "materials"
    ? Boolean(
        $("#materialSearch").value.trim() ||
        state.materialValueSelections.size ||
        state.excelFilters.materials.size
      )
    : tableType === "summary" ? Boolean(
        $$("[data-summary-column]").some((input) => input.value.trim()) ||
        state.summaryCategorySelected.size !== SUMMARY_CATEGORIES.length ||
        state.summaryValueSelections.size ||
        state.excelFilters.summary.size
      ) : Boolean($("#finishTypeSearch").value.trim() || state.excelFilters.finishType.size);
  const resultCount = tableType === "materials"
    ? state.filteredMaterials.length
    : tableType === "summary"
      ? state.filteredSummary.length
      : state.filteredFinishType.length;
  $$(`${section} .export`).forEach((button) => {
    const base = button.dataset.format === "csv" ? "CSV" : "Excel ↓";
    const exportCount = tableType === "materials" && hasFilter ? resultCount : count;
    const exportLabel = tableType === "materials" && hasFilter
      ? `${base} (${exportCount.toLocaleString()} matching)`
      : exportCount ? `${base} (${exportCount} selected)` : base;
    button.textContent = exportLabel;
    button.disabled = tableType === "finishType"
      ? resultCount === 0
      : count === 0 && (!hasFilter || resultCount === 0);
  });
}

function renderFinishType() {
  const query = $("#finishTypeSearch").value;
  state.filteredFinishType = applyExcelFilters(
    filterRows(state.snapshot.finishType || [], query),
    "finishType",
  );
  $("#finishTypeMatches").textContent = `${state.filteredFinishType.length.toLocaleString()} matching records`;
  renderTable($("#finishTypeTable"), state.filteredFinishType, "finishType");
  updateExportLabels("finishType");
}

function renderMaterials() {
  const query = $("#materialSearch").value;
  const searchedRows = filterRows(state.snapshot.materials || [], query).filter((row) =>
    [...state.materialValueSelections.entries()].every(([column, selected]) =>
      selected.has(excelValue(row, column))),
  );
  state.filteredMaterials = applyExcelFilters(searchedRows, "materials");
  $("#materialMatches").textContent = `${state.filteredMaterials.length.toLocaleString()} matching records`;
  renderTable($("#materialTable"), state.filteredMaterials, "materials");
  const hasFilter = Boolean(query.trim() || state.materialValueSelections.size || state.excelFilters.materials.size);
  $$("#materials .export").forEach((button) => {
    button.disabled = !hasFilter || !state.filteredMaterials.length;
  });
  updateMaterialValueFilterButtons();
  renderActiveMaterialFilters();
  updateExportLabels("materials");
}

function materialRowsForFilterOptions(excludedColumn = "") {
  let rows = filterRows(state.snapshot.materials || [], $("#materialSearch").value).filter((row) =>
    [...state.materialValueSelections.entries()].every(([column, selected]) =>
      column === excludedColumn || selected.has(excelValue(row, column))),
  );
  if (state.excelFilters.materials?.size) {
    rows = rows.filter((row) => [...state.excelFilters.materials.entries()].every(([column, excluded]) =>
      column === excludedColumn || !excluded.has(excelValue(row, column))));
  }
  return rows;
}

function pruneDependentMaterialSelections(protectedColumn) {
  for (const [column, selected] of [...state.materialValueSelections.entries()]) {
    if (column === protectedColumn) continue;
    const available = new Set(materialRowsForFilterOptions(column).map((row) => excelValue(row, column)));
    for (const value of [...selected]) {
      if (!available.has(value)) selected.delete(value);
    }
    if (!selected.size) state.materialValueSelections.delete(column);
  }
}

function finishTypeRowsForFilterOptions(excludedColumn = "") {
  let rows = filterRows(state.snapshot.finishType || [], $("#finishTypeSearch").value);
  if (state.excelFilters.finishType?.size) {
    rows = rows.filter((row) => [...state.excelFilters.finishType.entries()].every(([column, excluded]) =>
      column === excludedColumn || !excluded.has(excelValue(row, column))));
  }
  return rows;
}

function updateMaterialValueFilterButtons() {
  $$("#materials .material-value-filter-button").forEach((button) => {
    const selected = state.materialValueSelections.get(button.dataset.column);
    button.classList.toggle("active", Boolean(selected?.size));
    button.textContent = selected?.size ? `${selected.size.toLocaleString()} selected` : "Search and select...";
  });
}

function renderActiveMaterialFilters() {
  const chips = [];
  for (const [column, selected] of state.materialValueSelections.entries()) {
    selected.forEach((value) => chips.push({ column, value }));
  }
  const panel = $("#materialActiveFilters");
  const container = $("#materialActiveFilterChips");
  panel.classList.toggle("hidden", chips.length === 0);
  container.replaceChildren(...chips.map(({ column, value }) => {
    const chip = document.createElement("span");
    chip.append(document.createTextNode(value));
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "×";
    remove.title = `Remove ${value} from ${column}`;
    remove.dataset.materialFilterColumn = column;
    remove.dataset.materialFilterValue = value;
    chip.append(remove);
    return chip;
  }));
}

function renderSummary() {
  const columnFilters = $$("[data-summary-column]")
    .map((input) => ({
      column: input.dataset.summaryColumn,
      value: input.value.trim(),
    }))
    .filter((filter) => filter.value);
  const selectedCategories = state.summaryCategorySelected;
  const categoryFilterActive = selectedCategories.size !== SUMMARY_CATEGORIES.length;
  const searchedRows = (state.snapshot.summary || []).filter((row) =>
    columnFilters.every((filter) => valueMatchesAllTerms(row[filter.column], filter.value))
    && [...state.summaryValueSelections.entries()].every(([column, selected]) =>
      selected.has(excelValue(row, column)))
    && (!categoryFilterActive
      || [...selectedCategories].some((category) => rowMatchesSummaryCategory(row, category))),
  );
  state.filteredSummary = applyExcelFilters(
    searchedRows,
    "summary",
  );
  $("#summaryMatches").textContent = `${state.filteredSummary.length.toLocaleString()} matching records`;
  renderTable($("#summaryTable"), state.filteredSummary, "summary");
  const running = state.filteredSummary.reduce((sum, row) => sum + numberFrom(row, ["Knitted Qty."]), 0);
  const balance = state.filteredSummary.reduce((sum, row) => sum + numberFrom(row, ["Balance Qty."]), 0);
  $("#runningTotal").textContent = running.toLocaleString();
  $("#balanceTotal").textContent = balance.toLocaleString();
  updateSummaryFieldFilterButtons();
  renderSummaryCategoryOptions();
  renderActiveSummaryFilters();
  const filtered = Boolean(
    columnFilters.length
    || categoryFilterActive
    || state.excelFilters.summary.size
    || state.summaryValueSelections.size
  );
  $$("#summary .export").forEach((button) => {
    button.disabled = !filtered || !state.filteredSummary.length;
  });
  updateExportLabels("summary");
}


// ---- Production Calculation (Circular Knitting) ----
// Formulae from the user's Production_Calculation.xls (sheet "S.J or Pique"):
//   Production / shift (kg) = PI * Dia * Gauge * StitchLength * Feeders * RPM * 60 * shiftHours * efficiency
//                             / (25.4 * 36 * 840 * YarnCount * 2.2046)
//   Counter (rev)           = Production * 25.4 * 36 * 840 * YarnCount * 2.2046
//                             / (PI * Dia * Gauge * Feeders * StitchLength)
//   Production from counter = PI * Dia * Gauge * StitchLength * Feeders * Counter
//                             / (25.4 * 36 * 840 * YarnCount * 2.2046)
// The two counter conversions in the sheet carry no efficiency factor (theoretical,
// 100% efficiency); only the per-shift production uses the efficiency.
// The same standard formula applies to every circular-knitting fabric, so a fabric
// added later (Rib, Interlock, ...) reuses it with its own machine rows.
const PROD_PI = 3.1416;
const PROD_MM_PER_INCH = 25.4;
const PROD_INCH_PER_YARD = 36;
const PROD_YARDS_PER_HANK = 840;
const PROD_LB_PER_KG = 2.2046;
const PROD_BUILTIN_FABRICS = ["Single Jersey", "Pique"];
const PROD_SUGGESTED_FABRICS = ["Rib", "Interlock"];
const PROD_STORAGE_KEY = "zbom-production-calculation-v1";
const PROD_SAMPLE_ROW = { machine: "1", count: 50, stitchLength: 2.2, dia: 30, gauge: 28, feeders: 96, rpm: 25 };
let prodCalc = loadProdCalc();

function prodDefaultState() {
  return {
    fabric: "Single Jersey",
    shiftHours: 8,
    efficiency: 90,
    targetProduction: 10,
    counterInput: 1100,
    addedFabrics: [],
    rows: { "Single Jersey": [{ ...PROD_SAMPLE_ROW }], Pique: [{ ...PROD_SAMPLE_ROW }] },
  };
}
function loadProdCalc() {
  try {
    const raw = window.localStorage.getItem(PROD_STORAGE_KEY);
    if (!raw) return prodDefaultState();
    const parsed = JSON.parse(raw);
    const base = prodDefaultState();
    const merged = { ...base, ...parsed, rows: { ...base.rows, ...(parsed.rows || {}) } };
    merged.addedFabrics = (parsed.addedFabrics || []).filter((name) => typeof name === "string" && name.trim() && !PROD_BUILTIN_FABRICS.includes(name));
    PROD_BUILTIN_FABRICS.forEach((name) => { if (!Array.isArray(merged.rows[name]) || !merged.rows[name].length) merged.rows[name] = [{ ...PROD_SAMPLE_ROW }]; });
    merged.addedFabrics.forEach((name) => { if (!Array.isArray(merged.rows[name]) || !merged.rows[name].length) merged.rows[name] = [{ ...PROD_SAMPLE_ROW }]; });
    return merged;
  } catch (e) {
    return prodDefaultState();
  }
}
function saveProdCalc() {
  try { window.localStorage.setItem(PROD_STORAGE_KEY, JSON.stringify(prodCalc)); } catch (e) {}
}
function prodEsc(value) {
  return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function prodAvailableFabrics() { return [...PROD_BUILTIN_FABRICS, ...prodCalc.addedFabrics]; }
function prodRows() { return prodCalc.rows[prodCalc.fabric] || []; }
function prodNum(value) { const n = parseFloat(value); return Number.isFinite(n) ? n : NaN; }
function prodProductionPerShift(row) {
  const count = prodNum(row.count), sl = prodNum(row.stitchLength), dia = prodNum(row.dia);
  const gauge = prodNum(row.gauge), feeders = prodNum(row.feeders), rpm = prodNum(row.rpm);
  if ([count, sl, dia, gauge, feeders, rpm].some((v) => !Number.isFinite(v) || v <= 0)) return NaN;
  const eff = prodCalc.efficiency / 100, hours = prodCalc.shiftHours;
  if (!Number.isFinite(eff) || eff <= 0 || !Number.isFinite(hours) || hours <= 0) return NaN;
  return (PROD_PI * dia * gauge * sl * feeders * rpm * 60 * hours * eff)
    / (PROD_MM_PER_INCH * PROD_INCH_PER_YARD * PROD_YARDS_PER_HANK * count * PROD_LB_PER_KG);
}
function prodCounterFromProduction(row, kg) {
  const count = prodNum(row.count), sl = prodNum(row.stitchLength), dia = prodNum(row.dia);
  const gauge = prodNum(row.gauge), feeders = prodNum(row.feeders);
  if ([count, sl, dia, gauge, feeders].some((v) => !Number.isFinite(v) || v <= 0) || !Number.isFinite(kg) || kg < 0) return NaN;
  return (kg * PROD_MM_PER_INCH * PROD_INCH_PER_YARD * PROD_YARDS_PER_HANK * count * PROD_LB_PER_KG)
    / (PROD_PI * dia * gauge * feeders * sl);
}
function prodProductionFromCounter(row, counter) {
  const count = prodNum(row.count), sl = prodNum(row.stitchLength), dia = prodNum(row.dia);
  const gauge = prodNum(row.gauge), feeders = prodNum(row.feeders);
  if ([count, sl, dia, gauge, feeders].some((v) => !Number.isFinite(v) || v <= 0) || !Number.isFinite(counter) || counter < 0) return NaN;
  return (PROD_PI * dia * gauge * sl * feeders * counter)
    / (PROD_MM_PER_INCH * PROD_INCH_PER_YARD * PROD_YARDS_PER_HANK * count * PROD_LB_PER_KG);
}
function prodFmt(value, decimals = 2) {
  return Number.isFinite(value) ? value.toLocaleString("en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) : "—";
}
function prodKg(value) {
  return Number.isFinite(value) ? `${prodFmt(value)} kg` : "—";
}
function prodFillFabricSelect() {
  const select = $("#prodFabricSelect");
  if (!select) return;
  const available = prodAvailableFabrics();
  const pending = PROD_SUGGESTED_FABRICS.filter((name) => !available.includes(name));
  let html = `<optgroup label="Ready">` + available.map((name) => `<option value="${prodEsc(name)}">${prodEsc(name)}</option>`).join("") + `</optgroup>`;
  if (pending.length) html += `<optgroup label="Not added yet">` + pending.map((name) => `<option value="${prodEsc(name)}">${prodEsc(name)} — not added yet</option>`).join("") + `</optgroup>`;
  select.innerHTML = html;
  select.value = prodCalc.fabric;
  if (select.value !== prodCalc.fabric) select.value = available[0];
}
function renderProductionCalculation() {
  prodFillFabricSelect();
  const shiftInput = $("#prodShiftHours"), effInput = $("#prodEfficiency");
  if (shiftInput) shiftInput.value = prodCalc.shiftHours;
  if (effInput) effInput.value = prodCalc.efficiency;
  const ready = prodAvailableFabrics().includes(prodCalc.fabric);
  $("#prodCalcBody")?.classList.toggle("hidden", !ready);
  const notice = $("#prodFabricNotice");
  if (notice) {
    notice.classList.toggle("hidden", ready);
    if (!ready) {
      notice.innerHTML = `Production calculation for <b>${prodEsc(prodCalc.fabric)}</b> is not added yet. <button type="button" class="secondary" id="prodAddThisFabric">Add ${prodEsc(prodCalc.fabric)} Calculation</button>`;
      $("#prodAddThisFabric")?.addEventListener("click", () => prodAddFabric(prodCalc.fabric));
    }
  }
  if (ready) {
    renderProdTable();
    renderProdReverse();
  }
}
function renderProdTable() {
  const body = $("#prodTableBody");
  if (!body) return;
  const rows = prodRows();
  body.innerHTML = rows.map((row, i) => `
    <tr data-row="${i}">
      <td><input data-field="machine" value="${prodEsc(row.machine ?? "")}" inputmode="numeric"></td>
      <td><input data-field="count" type="number" min="0" step="any" value="${prodEsc(row.count ?? "")}"></td>
      <td><input data-field="stitchLength" type="number" min="0" step="any" value="${prodEsc(row.stitchLength ?? "")}"></td>
      <td><input data-field="dia" type="number" min="0" step="any" value="${prodEsc(row.dia ?? "")}"></td>
      <td><input data-field="gauge" type="number" min="0" step="any" value="${prodEsc(row.gauge ?? "")}"></td>
      <td><input data-field="feeders" type="number" min="0" step="1" value="${prodEsc(row.feeders ?? "")}"></td>
      <td><input data-field="rpm" type="number" min="0" step="any" value="${prodEsc(row.rpm ?? "")}"></td>
      <td class="prod-result" data-result="${i}">${prodKg(prodProductionPerShift(row))}</td>
      <td><button type="button" class="secondary prod-remove-row" data-row="${i}" title="Remove machine">✕</button></td>
    </tr>`).join("");
  prodUpdateTotals();
}
function prodUpdateTotals() {
  const rows = prodRows();
  const total = rows.reduce((sum, row) => sum + (prodProductionPerShift(row) || 0), 0);
  const totalEl = $("#prodTotal");
  if (totalEl) totalEl.textContent = `${prodFmt(total)} kg`;
  const matches = $("#prodMatches");
  if (matches) matches.textContent = `${rows.length} machine${rows.length === 1 ? "" : "s"}`;
}
function prodRecalcRow(i) {
  const row = prodRows()[i];
  const cell = document.querySelector(`#prodTableBody [data-result="${i}"]`);
  if (cell && row) cell.textContent = prodKg(prodProductionPerShift(row));
  prodUpdateTotals();
  prodUpdateReverseResults();
}
function prodReverseRow(selectId) {
  const select = $(selectId);
  const idx = select ? parseInt(select.value, 10) : 0;
  return prodRows()[Number.isFinite(idx) ? idx : 0];
}
function renderProdReverse() {
  const rows = prodRows();
  ["#prodCounterMachine", "#prodProductionMachine"].forEach((id) => {
    const select = $(id);
    if (!select) return;
    const prev = select.value;
    select.innerHTML = rows.map((row, i) => `<option value="${i}">M/C ${prodEsc(row.machine ?? i + 1)}</option>`).join("");
    select.value = prev && rows[+prev] ? prev : "0";
  });
  prodUpdateReverseResults();
}
function prodUpdateReverseResults() {
  const kg = prodNum($("#prodTargetProduction")?.value);
  const counter = prodNum($("#prodCounterInput")?.value);
  const cRow = prodReverseRow("#prodCounterMachine");
  const pRow = prodReverseRow("#prodProductionMachine");
  const cResult = $("#prodCounterResult");
  if (cResult && cRow) {
    const v = prodCounterFromProduction(cRow, Number.isFinite(kg) ? kg : prodCalc.targetProduction);
    cResult.textContent = Number.isFinite(v) ? `${prodFmt(v, 0)} rev` : "—";
  }
  const pResult = $("#prodProductionResult");
  if (pResult && pRow) {
    const v = prodProductionFromCounter(pRow, Number.isFinite(counter) ? counter : prodCalc.counterInput);
    pResult.textContent = Number.isFinite(v) ? `${prodFmt(v)} kg` : "—";
  }
}
function prodAddFabric(name) {
  const clean = String(name || "").trim();
  if (!clean) { showNotice("Enter a fabric name to add it.", "warning"); return; }
  if (prodAvailableFabrics().includes(clean)) {
    prodCalc.fabric = clean;
    saveProdCalc();
    renderProductionCalculation();
    return;
  }
  prodCalc.addedFabrics.push(clean);
  prodCalc.rows[clean] = [{ ...PROD_SAMPLE_ROW }];
  prodCalc.fabric = clean;
  saveProdCalc();
  renderProductionCalculation();
  showNotice(`${clean} production calculation added.`, "success");
}

function renderDiaWiseSummary() {
  const source = state.snapshot.diaWiseSummary || [];
  const term = normalizedSearchText($("#diaSummarySearch")?.value || "");
  state.filteredDiaWiseSummary = term
    ? source.filter((row) => normalizedSearchText(Object.values(row).join(" ")).includes(term))
    : source;
  const rows = state.filteredDiaWiseSummary;
  $("#diaSummaryMatches").textContent = `${rows.length.toLocaleString()} matching records`;
  const totalFor = (column) => term
    ? rows.reduce((sum, row) => sum + diaNumber(row, column), 0)
    : diaNumber(state.snapshot.diaWiseTotals || {}, column);
  $("#diaKnittingBalance").textContent = formatTons(totalFor("Knitting Balance"));

  const balanceGrid = $("#diaBalanceGrid");
  balanceGrid.replaceChildren(...DIA_BALANCE_COLUMNS.map((column) => {
    const value = totalFor(column);
    const card = document.createElement("article");
    card.className = value < 0 ? "negative" : value > 0 ? "positive" : "zero";
    const label = document.createElement("small");
    label.textContent = column;
    const total = document.createElement("strong");
    total.textContent = formatTons(value);
    card.append(label, total);
    return card;
  }));

  const table = $("#diaSummaryTable");
  table.innerHTML = "";
  const columns = source.length ? Object.keys(source[0]) : [];
  if (!columns.length) {
    const body = document.createElement("tbody");
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.textContent = "Dia Wise Summary sheet is not available. Press Update Now after adding the latest Circular Knitting workbook.";
    row.appendChild(cell);
    body.appendChild(row);
    table.appendChild(body);
    return;
  }
  const widths = columns.map((column) => {
    const longest = Math.max(column.length, ...rows.slice(0, 100).map((row) => String(row[column] ?? "").length));
    return Math.min(230, Math.max(90, longest * 6 + 24));
  });
  const colgroup = document.createElement("colgroup");
  widths.forEach((width) => {
    const col = document.createElement("col");
    col.style.width = `${width}px`;
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);
  table.style.width = `${widths.reduce((sum, width) => sum + width, 0)}px`;
  const head = document.createElement("thead");
  const headerRow = document.createElement("tr");
  columns.forEach((column) => {
    const cell = document.createElement("th");
    cell.textContent = column;
    cell.title = column;
    headerRow.appendChild(cell);
  });
  head.appendChild(headerRow);
  table.appendChild(head);
  const body = document.createElement("tbody");
  rows.forEach((row) => {
    const tableRow = document.createElement("tr");
    columns.forEach((column) => {
      const cell = document.createElement("td");
      const value = row[column] ?? "";
      cell.textContent = String(value);
      cell.title = cell.textContent;
      if (DIA_BALANCE_COLUMNS.includes(column) || ["Order Qty", "Knitting QTY.", "Knitting Balance", "Yarn Balance"].includes(column)) {
        const numeric = diaNumber(row, column);
        cell.classList.add("numeric-cell");
        if (numeric < 0) cell.classList.add("negative-value");
      }
      tableRow.appendChild(cell);
    });
    body.appendChild(tableRow);
  });
  table.appendChild(body);
}

function compactFilterValues(values) {
  if (values.length <= 8) return values.join(", ");
  return `${values.slice(0, 6).join(", ")} +${values.length - 6} more`;
}

function renderActiveSummaryFilters() {
  const chips = [];
  for (const [column, selected] of state.summaryValueSelections.entries()) {
    selected.forEach((value) => chips.push({
      key: column,
      value,
      label: value,
    }));
  }
  if (state.summaryCategorySelected.size !== SUMMARY_CATEGORIES.length) {
    [...state.summaryCategorySelected].forEach((value) => chips.push({
      key: "__category__",
      value,
      label: value,
    }));
  }
  const panel = $("#summaryActiveFilters");
  const container = $("#summaryActiveFilterChips");
  panel.classList.toggle("hidden", chips.length === 0);
  container.replaceChildren(...chips.map(({ key, value, label }) => {
    const chip = document.createElement("span");
    chip.append(document.createTextNode(label));
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "×";
    remove.title = value
      ? `Remove ${value} from ${key}`
      : `Clear ${key === "__category__" ? "Machine / Dia / Fabric Type" : key} filter`;
    remove.dataset.summaryFilterKey = key;
    if (value) remove.dataset.summaryFilterValue = value;
    chip.append(remove);
    return chip;
  }));
}

function updateSummaryFieldFilterButtons() {
  $$("#summary .summary-field-filter-button").forEach((button) => {
    const selected = state.summaryValueSelections.get(button.dataset.column);
    button.classList.toggle("active", Boolean(selected?.size));
    button.textContent = selected?.size
      ? `${selected.size.toLocaleString()} selected`
      : "Search and select...";
    button.title = selected?.size
      ? `${selected.size} value(s) selected. Click to change.`
      : `Select multiple ${button.dataset.column} values`;
  });
}

function populateSummarySearchOptions() {
  const mappings = [
    ["Buyer", "summaryBuyerOptions"],
    ["Order NO", "summaryOrderOptions"],
    ["SO-Item", "summarySoItemOptions"],
    ["Style", "summaryStyleOptions"],
    ["Material Number", "summaryMaterialOptions"],
    ["Material Description", "summaryDescriptionOptions"],
  ];
  for (const [column, listId] of mappings) {
    const list = $(`#${listId}`);
    if (!list) continue;
    const values = [...new Set(
      (state.snapshot.summary || [])
        .map((row) => String(row[column] ?? "").trim())
        .filter(Boolean),
    )].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    list.replaceChildren(...values.map((value) => {
      const option = document.createElement("option");
      option.value = value;
      return option;
    }));
  }
}

function rowMatchesSummaryCategory(row, category) {
  const machineDia = String(row["M/C Dia"] ?? "").trim();
  const searchable = Object.values(row).map((value) => String(value ?? "")).join(" ");
  if (category === "NEED DIA") return !machineDia || machineDia === "0";
  if (category === "L/RIB") return /l\s*\/\s*rib/i.test(searchable);
  if (category === "RIB") return /\brib\b/i.test(searchable) && !/l\s*\/\s*rib/i.test(searchable);
  if (category === "L/S/J") return /l\s*\/\s*s\s*\/\s*j/i.test(searchable);
  return searchValueMatches(searchable, category);
}

function summaryRowsForFilterOptions(excludedColumn = "") {
  const columnFilters = $$('[data-summary-column]')
    .map((input) => ({ column: input.dataset.summaryColumn, value: input.value.trim() }))
    .filter((filter) => filter.value && filter.column !== excludedColumn);
  const categoryFilterActive = excludedColumn !== "__category__"
    && state.summaryCategorySelected.size !== SUMMARY_CATEGORIES.length;
  let rows = (state.snapshot.summary || []).filter((row) =>
    columnFilters.every((filter) => valueMatchesAllTerms(row[filter.column], filter.value))
    && [...state.summaryValueSelections.entries()].every(([column, selected]) =>
      column === excludedColumn || selected.has(excelValue(row, column)))
    && (!categoryFilterActive
      || [...state.summaryCategorySelected].some((category) => rowMatchesSummaryCategory(row, category))),
  );
  if (state.excelFilters.summary?.size) {
    rows = rows.filter((row) => [...state.excelFilters.summary.entries()].every(([column, excluded]) =>
      column === excludedColumn || !excluded.has(excelValue(row, column))));
  }
  return rows;
}

function pruneDependentSummarySelections(protectedColumn) {
  for (const [column, selected] of [...state.summaryValueSelections.entries()]) {
    if (column === protectedColumn) continue;
    const available = new Set(summaryRowsForFilterOptions(column).map((row) => excelValue(row, column)));
    for (const value of [...selected]) {
      if (!available.has(value)) selected.delete(value);
    }
    if (!selected.size) state.summaryValueSelections.delete(column);
  }
  if (protectedColumn !== "__category__"
    && state.summaryCategorySelected.size !== SUMMARY_CATEGORIES.length) {
    const relatedRows = summaryRowsForFilterOptions("__category__");
    const availableCategories = new Set(SUMMARY_CATEGORIES.filter((category) =>
      relatedRows.some((row) => rowMatchesSummaryCategory(row, category))));
    state.summaryCategorySelected = new Set(
      [...state.summaryCategorySelected].filter((category) => availableCategories.has(category)),
    );
    if (!state.summaryCategorySelected.size) {
      state.summaryCategorySelected = new Set(SUMMARY_CATEGORIES);
    }
  }
}

function renderSummaryCategoryOptions() {
  const container = $("#summaryCategoryOptions");
  const relatedRows = summaryRowsForFilterOptions("__category__");
  const availableCategories = SUMMARY_CATEGORIES.filter((category) =>
    relatedRows.some((row) => rowMatchesSummaryCategory(row, category)));
  container.replaceChildren(...availableCategories.map((category) => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = category;
    checkbox.checked = state.summaryCategorySelected.has(category);
    const span = document.createElement("span");
    span.textContent = category;
    label.append(checkbox, span);
    return label;
  }));
  const selected = state.summaryCategorySelected.size;
  $("#summaryCategoryButton").textContent = selected === SUMMARY_CATEGORIES.length
    ? "All results ▾"
    : selected === 0
      ? "No category selected ▾"
      : `${selected} categories selected ▾`;
  updateSummaryCategoryToggleLabel();
}

function updateSummaryCategoryToggleLabel() {
  const checkboxes = $$("#summaryCategoryOptions input");
  const allChecked = checkboxes.length > 0 && checkboxes.every((checkbox) => checkbox.checked);
  $("#summaryCategoryToggleAll").textContent = allChecked ? "Clear All" : "Select All";
}

function currentExcelSource() {
  if (!state.openExcelFilter) return [];
  return sourceRowsFor(state.openExcelFilter.tableType);
}

function renderExcelValueList() {
  if (!state.openExcelFilter) return;
  const { column, pendingExcluded, pendingSelected, allValues, summaryTopMode, valueFilterTable } = state.openExcelFilter;
  const term = $("#excelValueSearch").value.trim().toLowerCase();
  const matching = term
    ? allValues.filter((value) => value.toLowerCase().includes(term))
    : allValues;
  const visible = matching.slice(0, 400);
  const list = $("#excelValueList");
  list.innerHTML = "";

  for (const value of visible) {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = summaryTopMode ? pendingSelected.has(value) : !pendingExcluded.has(value);
    checkbox.addEventListener("change", () => {
      if (summaryTopMode) {
        if (checkbox.checked) pendingSelected.add(value);
        else pendingSelected.delete(value);
        const selections = valueFilterTable === "materials"
          ? state.materialValueSelections
          : state.summaryValueSelections;
        if (pendingSelected.size) selections.set(column, new Set(pendingSelected));
        else selections.delete(column);
        if (valueFilterTable === "materials") pruneDependentMaterialSelections(column);
        else pruneDependentSummarySelections(column);
        refreshFilteredTable(valueFilterTable);
      } else if (checkbox.checked) {
        pendingExcluded.delete(value);
      } else {
        pendingExcluded.add(value);
      }
      $("#excelSelectAll").checked = pendingExcluded.size === 0;
      $("#excelSelectAll").indeterminate = pendingExcluded.size > 0 && pendingExcluded.size < allValues.length;
      updateExcelSelectAllLabel();
      renderExcelPendingSelection();
    });
    const text = document.createElement("span");
    text.textContent = value;
    label.append(checkbox, text);
    list.appendChild(label);
  }

  if ($("#excelFilterNote")) {
    $("#excelFilterNote").textContent =
      matching.length > visible.length
        ? `Showing first ${visible.length.toLocaleString()} of ${matching.length.toLocaleString()} values. Type to search more.`
        : `${matching.length.toLocaleString()} unique values`;
  }
  if ($("#excelSelectAll")) $("#excelSelectAll").checked = summaryTopMode ? false : pendingExcluded.size === 0;
  if ($("#excelSelectAll")) $("#excelSelectAll").indeterminate = summaryTopMode ? false : pendingExcluded.size > 0 && pendingExcluded.size < allValues.length;
  updateExcelSelectAllLabel();
  renderExcelPendingSelection();
}

function updateExcelSelectAllLabel() {
  const menu = $("#excelFilterMenu");
  const label = menu.querySelector(".excel-select-all span");
  if (!label) return;
  label.textContent = menu.classList.contains("summary-top-excel-filter")
    ? ($("#excelSelectAll").checked ? "Clear All" : "Select All")
    : "(Select All)";
}

function renderExcelPendingSelection() {
  const container = $("#excelPendingSelection");
  const menu = $("#excelFilterMenu");
  if (!state.openExcelFilter || menu.classList.contains("summary-top-excel-filter")) {
    container.classList.add("hidden");
    container.replaceChildren();
    return;
  }
  const { allValues, pendingExcluded } = state.openExcelFilter;
  const included = allValues.filter((value) => !pendingExcluded.has(value));
  const showIncluded = included.length <= pendingExcluded.size;
  const values = showIncluded ? included : [...pendingExcluded];
  const heading = document.createElement("b");
  if (pendingExcluded.size === 0) heading.textContent = "All results selected";
  else if (included.length === 0) heading.textContent = "No value selected";
  else heading.textContent = showIncluded ? `Selected (${included.length})` : `Excluded (${pendingExcluded.size})`;
  const chips = document.createElement("div");
  const visible = values.slice(0, 10);
  visible.forEach((value) => {
    const chip = document.createElement("span");
    chip.append(document.createTextNode(value));
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "×";
    remove.dataset.pendingFilterValue = value;
    remove.dataset.pendingFilterMode = showIncluded ? "included" : "excluded";
    chip.append(remove);
    chips.append(chip);
  });
  if (values.length > visible.length) {
    const more = document.createElement("em");
    more.textContent = `+${values.length - visible.length} more`;
    chips.append(more);
  }
  container.classList.remove("hidden");
  container.replaceChildren(heading, chips);
}

function closeExcelFilter() {
  $("#excelFilterMenu").classList.add("hidden");
  state.openExcelFilter = null;
}

function openExcelFilter(button) {
  const tableType = button.dataset.tableType;
  const column = button.dataset.column;
  const summaryTopMode = button.classList.contains("summary-field-filter-button");
  const valueFilterTable = button.dataset.valueFilter || (summaryTopMode ? "summary" : "");
  const sourceRows = valueFilterTable === "materials"
    ? materialRowsForFilterOptions(column)
    : summaryTopMode
      ? summaryRowsForFilterOptions(column)
      : tableType === "materials"
        ? materialRowsForFilterOptions(column)
        : tableType === "summary"
          ? summaryRowsForFilterOptions(column)
          : finishTypeRowsForFilterOptions(column);
  const allValues = [...new Set(sourceRows.map((row) => excelValue(row, column)))]
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  const saved = state.excelFilters[tableType].get(column);
  state.openExcelFilter = {
    tableType,
    column,
    allValues,
    pendingExcluded: new Set(saved || []),
    pendingSelected: new Set(summaryTopMode
      ? ((valueFilterTable === "materials" ? state.materialValueSelections : state.summaryValueSelections).get(column) || [])
      : []),
    summaryTopMode,
    valueFilterTable,
  };
  $("#excelFilterTitle").textContent = column;
  $("#excelValueSearch").value = "";
  const menu = $("#excelFilterMenu");
  menu.classList.toggle("summary-top-excel-filter", summaryTopMode);
  if ($("#applyExcelFilter")) $("#applyExcelFilter").classList.toggle("hidden", summaryTopMode);
  if ($("#cancelExcelFilter")) $("#cancelExcelFilter").classList.toggle("hidden", summaryTopMode);
  menu.classList.remove("hidden");
  const rect = button.getBoundingClientRect();
  const menuWidth = summaryTopMode ? 330 : 242;
  const left = Math.min(Math.max(8, rect.left), window.innerWidth - menuWidth - 8);
  const menuHeight = summaryTopMode ? 455 : 340;
  const top = Math.min(rect.bottom + 4, window.innerHeight - menuHeight);
  menu.style.left = `${left}px`;
  menu.style.top = `${Math.max(8, top)}px`;
  renderExcelValueList();
  $("#excelValueSearch").focus();
}

function refreshFilteredTable(tableType) {
  state.currentPage[tableType] = 1;
  if (tableType === "materials") renderMaterials();
  else if (tableType === "summary") renderSummary();
  else renderFinishType();
}

function renderState() {
  const { snapshot } = state;
  state.selectedRows.materials.clear();
  state.selectedRows.summary.clear();
  state.selectedRows.finishType.clear();
  $("#materialCount").textContent = (snapshot.materials || []).length.toLocaleString();
  $("#finishTypeCount").textContent = (snapshot.finishType || []).length.toLocaleString();
  $("#summaryCount").textContent = (snapshot.summary || []).length.toLocaleString();
  state.summaryOptionsReady = false;
  $("#lastUpdate").textContent = formatDate(snapshot.updatedAt);
  $("#dataStatus").textContent = snapshot.updatedAt ? "Available offline" : "Update required";
  $("#sideUpdated").textContent = snapshot.updatedAt ? `Updated ${formatDate(snapshot.updatedAt)}` : "No local update";
  $("#materialSource").textContent = snapshot.sources?.materials
    ? `${snapshot.sources.materials.filename} · ${formatDate(snapshot.sources.materials.modifiedAt)}`
    : "Waiting for first update";
  $("#summarySource").textContent = snapshot.sources?.summary
    ? `${snapshot.sources.summary.filename} · ${formatDate(snapshot.sources.summary.modifiedAt)}`
    : "Waiting for first update";
  $("#finishTypeSource").textContent = snapshot.sources?.finishType
    ? `${snapshot.sources.finishType.filename} · ${formatDate(snapshot.sources.finishType.modifiedAt)}`
    : "Waiting for first update";
  $("#materialFileUpdated").textContent = snapshot.sources?.materials
    ? formatDate(snapshot.sources.materials.modifiedAt)
    : "Not updated yet";
  $("#materialFileName").textContent = snapshot.sources?.materials?.filename || "No source file";
  $("#summaryFileUpdated").textContent = snapshot.sources?.summary
    ? formatDate(snapshot.sources.summary.modifiedAt)
    : "Not updated yet";
  $("#summaryFileName").textContent = snapshot.sources?.summary?.filename || "No source file";
  $("#finishTypeFileUpdated").textContent = snapshot.sources?.finishType
    ? formatDate(snapshot.sources.finishType.modifiedAt)
    : "Not updated yet";
  $("#finishTypeFileName").textContent = snapshot.sources?.finishType?.filename || "No source file";
  if (state.view === "materials") renderMaterials();
  if (state.view === "summary") {
    populateSummarySearchOptions();
    state.summaryOptionsReady = true;
    renderSummary();
  }
  if (state.view === "diaWiseSummary") renderDiaWiseSummary();
  if (state.view === "finishType") renderFinishType();
  renderOverviewStatus();
  updateHeaderLastUpdated();
}

function renderOverviewStatus() {
  const list = $("#overviewDepartmentStatus");
  if (!list) return;
  list.innerHTML = "";
  const moduleRows = [
    {
      name: "Fabric Material Code",
      file: state.snapshot.sources?.materials?.filename,
      updatedAt: state.snapshot.sources?.materials?.modifiedAt,
      view: "materials",
    },
    {
      name: "Finish Type Short Name",
      file: state.snapshot.sources?.finishType?.filename,
      updatedAt: state.snapshot.sources?.finishType?.modifiedAt,
      view: "finishType",
    },
  ];
  const departmentRows = (state.departments || []).map((department, index) => {
    const status = state.departmentStatuses?.[department];
    const summarySource = index === 0 ? state.snapshot.sources?.summary : null;
    return {
      name: department,
      file: status?.file || summarySource?.filename,
      updatedAt: status?.uploadedAt || summarySource?.modifiedAt,
      view: index === 0 ? "summary" : "department",
      department,
    };
  });
  const rows = [...moduleRows, ...departmentRows];
  const availableDepartments = departmentRows.filter((row) => row.file).length;
  $("#departmentFilesCount").textContent = `${availableDepartments} / ${departmentRows.length}`;
  $("#overviewStatusSummary").textContent = `${rows.filter((row) => row.file).length} of ${rows.length} modules ready`;
  rows.forEach((item) => {
    const row = document.createElement("button");
    row.className = `overview-status-row ${item.file ? "ready" : "waiting"}`;
    const name = document.createElement("span");
    name.innerHTML = `<i></i><b></b>`;
    name.querySelector("b").textContent = item.name;
    const file = document.createElement("span");
    file.textContent = item.file || "No file uploaded";
    const updated = document.createElement("span");
    updated.textContent = item.updatedAt ? formatDate(item.updatedAt) : "Not updated yet";
    const status = document.createElement("em");
    status.textContent = item.file ? "Available" : "Pending";
    row.append(name, file, updated, status);
    row.addEventListener("click", () => {
      if (item.department) state.department = item.department;
      changeView(item.view);
    });
    list.appendChild(row);
  });
}

function setConnection(online, message) {
  $("#connection").textContent = online ? "● LAN available" : "● Offline";
  $("#connection").className = `connection ${online ? "online" : "offline"}`;
  $("#sideStatus").className = `side-status ${online ? "" : "offline"}`;
  $("#sideStatus b").textContent = online ? "LAN available" : "Offline data mode";
  if (message) $("#sideUpdated").textContent = message;
}

// The four upload modules used to be restricted to Master User / Admin.
// Since 2026-11-12 they are free for everyone (client-side generation,
// no login — see web-erp.js and ONLINE_PLAN.md), so the set is empty.
// To restrict a module again later, add its view id back here.
const RESTRICTED_MODULES = new Set([]);
const RESTRICTED_ROLES = new Set(["master_user", "admin", "main_admin"]);

function canAccessRestrictedModule() {
  return RESTRICTED_ROLES.has(state.access?.role);
}

function showAccessRestricted() {
  const overlay = $("#accessRestrictedOverlay");
  if (overlay) overlay.classList.remove("hidden");
  else showNotice("Access Restricted. Please log in as Master User or Admin to access this module.", "warning");
}

function setMaterialMenuExpanded(expanded) {
  const group = $("#materialSubmodules");
  const parent = $("#materialModuleToggle");
  if (!group || !parent) return;
  group.classList.toggle("hidden", !expanded);
  parent.classList.toggle("expanded", expanded);
  parent.setAttribute("aria-expanded", String(expanded));
}

function changeView(view) {
  if (!view) return;
  try {
    if (RESTRICTED_MODULES.has(view) && !canAccessRestrictedModule()) {
      showAccessRestricted();
      return;
    }
    const previousView = state.view;
    const seconds = Math.max(1, Math.round((Date.now() - (state.viewStartedAt || Date.now())) / 1000));
    const previousLabel = {
      dashboard: "Overview", materials: "Fabric Material Code", codeRequirement: "Code Requirement File",
      routingUploadFormat: "Routing Upload File", characteristicsUpload: "Characteristics Upload File", finishType: "Finish Type Short Name",
      summary: "Circular Knitting", diaWiseSummary: "Dia Wise Summary", productionCalculation: "Production Calculation", department: state.department || "Department Workspace",
      administration: "Administration", settings: "Settings",
    }[previousView] || previousView;
    if (previousView && window.erp && typeof window.erp.recordUsage === "function") {
      try {
        window.erp.recordUsage({ module: previousLabel, seconds, file: state.departmentStatuses?.[state.department]?.file || "" }).catch(() => {});
      } catch (e) {}
    }
    state.view = view;
    state.viewStartedAt = Date.now();
    const materialChildViews = new Set(["codeRequirement", "finishType", "zkngMaterialUpload", "lsmwBomUpload", "routingUploadFormat", "characteristicsUpload"]);
    if (materialChildViews.has(view) || view === "materials") setMaterialMenuExpanded(true);
    $$(".view").forEach((element) => element.classList.toggle("active", element.id === view));
    $$("nav button, .submodule-nav").forEach((button) => button.classList.toggle(
      "active",
      button.dataset.view === view && (view !== "department" || button.dataset.department === state.department),
    ));
    const pageTitle = $("#pageTitle");
    if (pageTitle) {
      pageTitle.textContent = {
        dashboard: "Overview",
        materials: "Fabric Material Code",
        codeRequirement: "Code Requirement File",
        routingUploadFormat: "Routing Upload File",
        zkngMaterialUpload: "ZKNG Material Upload File",
        lsmwBomUpload: "LSMW BOM Upload File",
        characteristicsUpload: "Characteristics Upload File",
        finishType: "Finish Type Short Name",
        summary: state.departments[0] || "Circular Knitting",
        diaWiseSummary: "Dia Wise Summary",
        productionCalculation: "Production Calculation",
        department: state.department || "Department Workspace",
        administration: "Administration",
        settings: "Settings",
      }[view] || "Overview";
    }
    if (view === "dashboard") renderOverviewStatus();
    if (view === "department") {
      if ($("#departmentWorkspaceTitle")) $("#departmentWorkspaceTitle").textContent = state.department;
      renderDepartmentStatus(state.department);
    }
    if (view === "materials") renderMaterials();
    if (view === "codeRequirement") renderCodeRequirementPreview();
    if (view === "summary") {
      if (!state.summaryOptionsReady) {
        populateSummarySearchOptions();
        state.summaryOptionsReady = true;
      }
      renderSummary();
    }
    if (view === "diaWiseSummary") renderDiaWiseSummary();
    if (view === "productionCalculation") renderProductionCalculation();
    if (view === "finishType") renderFinishType();
    if (view === "routingUploadFormat") refreshRoutingState();
    if (view === "zkngMaterialUpload") refreshZkngState();
    if (view === "lsmwBomUpload") refreshLsmwBomState();
    if (view === "characteristicsUpload") refreshCharacteristicsState();
    updateHeaderLastUpdated(view);
  } catch (error) {
    console.error("changeView error:", error);
  }
}

function updateHeaderLastUpdated(view = state.view) {
  const sourceId = {
    materials: "materialFileUpdated",
    codeRequirement: "materialFileUpdated",
    routingUploadFormat: "materialFileUpdated",
    zkngMaterialUpload: "materialFileUpdated",
    lsmwBomUpload: "materialFileUpdated",
    finishType: "finishTypeFileUpdated",
    summary: "summaryFileUpdated",
    diaWiseSummary: "summaryFileUpdated",
    department: "departmentFileUpdated",
  }[view];
  const headerUpdate = $("#headerLastUpdated");
  if (!sourceId) {
    headerUpdate.classList.add("hidden");
    return;
  }
  const value = $(`#${sourceId}`)?.textContent || "Not updated yet";
  headerUpdate.querySelector("strong").textContent = value;
  headerUpdate.classList.remove("hidden");
}

function renderDepartmentStatus(department) {
  const status = state.departmentStatuses?.[department];
  $("#departmentUpdateTitle").textContent = department || "Department";
  $("#departmentFileUpdated").textContent = status?.uploadedAt ? formatDate(status.uploadedAt) : "Not updated yet";
  $("#departmentFileName").textContent = status?.file || "No file uploaded";
  if (state.view === "department") updateHeaderLastUpdated("department");
}

function initials(name) {
  return String(name || "User")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0])
    .join("")
    .toUpperCase();
}

function fillDepartmentSelect(select, departments) {
  select.innerHTML = "";
  departments.forEach((department) => {
    const option = document.createElement("option");
    option.value = department;
    option.textContent = department;
    select.appendChild(option);
  });
}

function openDepartmentView(department) {
  state.department = department;
  changeView("department");
}

function renderDepartmentNavigation() {
  const primary = state.departments[0] || "Circular Knitting";
  $("#primarySummaryNav span").textContent = primary;
  $("#dashboardSummaryName").firstChild.textContent = primary;
  $("#summaryMetricName").textContent = `${primary} records`;
  $("#summarySourceName").textContent = primary;
  $("#summaryUpdateTitle").textContent = primary;
  const nav = $("#departmentNav");
  nav.innerHTML = "";
  state.departments.slice(1).forEach((department) => {
    const button = document.createElement("button");
    button.dataset.view = "department";
    button.dataset.department = department;
    const icon = document.createElement("i");
    icon.textContent = "▤";
    button.append(icon, document.createTextNode(department));
    button.addEventListener("click", () => openDepartmentView(department));
    nav.appendChild(button);
  });
}

function renderSubAdmins() {
  const list = $("#subAdminList");
  list.innerHTML = "";
  if (!state.access.subAdmins?.length) {
    const empty = document.createElement("small");
    empty.textContent = "No Sub Admin has been assigned yet.";
    list.appendChild(empty);
    return;
  }
  state.access.subAdmins.forEach((admin) => {
    const row = document.createElement("div");
    const identity = document.createElement("span");
    const name = document.createElement("b");
    const detail = document.createElement("small");
    name.textContent = admin.fullName;
    detail.textContent = `${admin.employeeId} · ${admin.designation || "No designation"} · Login ID: ${admin.employeeId}`;
    identity.append(name, detail);
    const department = document.createElement("em");
    department.textContent = admin.department;
    row.append(identity, department);
    list.appendChild(row);
  });
}

function renderAdminDashboard() {
  renderUsageAnalytics();
  const accountRoleLabels = { guest: "Guest User", standard_user: "Standard User", power_user: "Power User", master_user: "Master User", department_head: "Department Head", admin: "Admin", main_admin: "Admin" };
  const subAdmins = state.access.subAdmins || [];
  $("#adminTotalUsers").textContent = String(1 + subAdmins.length);
  $("#adminActiveUsers").textContent = String(1 + subAdmins.filter((admin) => admin.active !== false).length);
  $("#adminDepartmentCount").textContent = String((state.departments || []).length);
  $("#adminUploadedFiles").textContent = String(state.departmentFiles.length);
  const body = $("#adminUserTable");
  if (!body) return;
  body.innerHTML = "";
  const protectedAdmin = state.access.protectedMainAdmin || {
    fullName: "Zakaria Alam Mithu", employeeId: "31001625", username: "31001625",
    designation: "CTM", department: "All Departments", passwordSet: false,
  };
  const accounts = [{
    fullName: protectedAdmin.fullName, employeeId: protectedAdmin.employeeId,
    windowsUsername: protectedAdmin.username, roleLabel: "Main Admin",
    designation: protectedAdmin.designation, department: protectedAdmin.department,
    recoveryEmail: protectedAdmin.recoveryEmail, passwordSet: protectedAdmin.passwordSet,
    active: true, main: true,
  }, ...subAdmins.map((admin) => ({
    ...admin,
    roleLabel: accountRoleLabels[admin.role] || "Standard User",
    main: false,
  }))];
  accounts.forEach((account) => {
    const row = document.createElement("tr");
    [account.fullName, account.employeeId, account.designation, account.roleLabel,
      account.department, account.passwordSet ? "Set" : "Not set"].forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = value || "";
      cell.title = value || "";
      row.appendChild(cell);
    });
    if (!account.main) {
      const roleCell = row.children[3];
      roleCell.textContent = "";
      const roleSelect = document.createElement("select");
      roleSelect.className = "inline-role-select";
      [
        ["admin", "Admin"],
        ["department_head", "Department Head"],
        ["master_user", "Master User"],
        ["power_user", "Power User"],
        ["standard_user", "Standard User"],
        ["guest", "Guest User"],
      ].forEach(([value, label]) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        roleSelect.appendChild(option);
      });
      roleSelect.value = account.role || "standard_user";
      const roleSave = document.createElement("button");
      roleSave.type = "button";
      roleSave.className = "inline-role-save";
      roleSave.textContent = "Save";
      roleSave.disabled = true;
      roleSelect.addEventListener("change", () => {
        roleSave.disabled = roleSelect.value === (account.role || "standard_user");
      });
      roleSave.addEventListener("click", async () => {
        const previousRole = account.role || "standard_user";
        roleSelect.disabled = true;
        roleSave.disabled = true;
        roleSave.textContent = "Saving…";
        const result = await changeAccountRole(account, roleSelect.value);
        if (!result.ok) roleSelect.value = previousRole;
        roleSelect.disabled = false;
        roleSave.textContent = "Save";
        roleSave.disabled = roleSelect.value === previousRole;
      });
      const roleEditor = document.createElement("div");
      roleEditor.className = "inline-role-editor";
      roleEditor.append(roleSelect, roleSave);
      roleCell.appendChild(roleEditor);
    }
    const status = document.createElement("td");
    const statusText = account.active === false
      ? "Inactive"
      : account.requestStatus === "pending"
        ? "Active · Sub-Admin Pending"
        : "Active";
    status.innerHTML = `<span class="account-status${account.active === false ? " inactive" : account.requestStatus === "pending" ? " pending" : ""}">${statusText}</span>`;
    status.title = statusText;
    row.appendChild(status);
    [account.lastSeenAt ? formatDate(account.lastSeenAt) : "Never",
      formatUsageTime(account.totalUsageSeconds),
      String(Number(account.loginCount) || 0)].forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = value;
      cell.title = value;
      row.appendChild(cell);
    });
    const control = document.createElement("td");
    control.className = "account-controls";
    const passwordControl = document.createElement("button");
    passwordControl.className = "table-action";
    passwordControl.textContent = "Password";
    passwordControl.addEventListener("click", () => {
      $("#passwordAdminAccount").value = account.windowsUsername;
      $("#newAdminPassword").value = "";
      $("#confirmAdminPassword").value = "";
      $("#adminPasswordPanel").scrollIntoView({ behavior: "smooth", block: "start" });
      window.setTimeout(() => $("#newAdminPassword").focus(), 350);
    });
    if (account.main) {
      const protectedLabel = document.createElement("span");
      protectedLabel.textContent = account.recoveryEmail
        ? `Permanent protected account · Recovery: ${account.recoveryEmail}`
        : "Permanent protected account";
      protectedLabel.title = "Role, department, status and account identity cannot be changed by another user.";
      control.append(protectedLabel);
      if (state.access.role === "main_admin") control.append(passwordControl);
    } else {
      const edit = document.createElement("button");
      edit.className = "table-action";
      edit.textContent = "Edit All";
      edit.addEventListener("click", () => {
        $("#subAdminName").value = account.fullName || "";
        $("#subAdminEmployeeId").value = account.employeeId || "";
        $("#subAdminMobile").value = account.mobile || "";
        $("#subAdminTelephone").value = account.telephone || "";
        $("#subAdminDesignation").value = account.designation || "";
        $("#subAdminCompany").value = account.company || "SFL";
        $("#subAdminRole").value = account.role || "standard_user";
        $("#subAdminDepartment").value = account.department || "";
        $("#subAdminDepartment").disabled = ["admin", "master_user"].includes(account.role || "standard_user");
        $("#subAdminPassword").value = "";
        const allowed = Array.isArray(account.filePermissions) ? account.filePermissions : ["materials", "codeRequirement", "finishType", "routingUploadFormat", "summary", "diaWiseSummary", "department"];
        $$("#subAdminFilePermissions input").forEach((input) => { input.checked = allowed.includes(input.value); });
        $("#subAdminName").scrollIntoView({ behavior: "smooth", block: "center" });
      });
      const toggle = document.createElement("button");
      toggle.className = "table-action";
      toggle.textContent = account.active === false ? "Activate" : "Deactivate";
      toggle.addEventListener("click", () => controlSubAdmin({ mode: "toggle", username: account.windowsUsername, active: account.active === false, button: toggle }));
      const remove = document.createElement("button");
      remove.className = "table-action danger";
      remove.textContent = "Remove";
      remove.addEventListener("click", () => {
        if (window.confirm(`Remove ${account.fullName}'s account?`)) {
          controlSubAdmin({ mode: "delete", username: account.windowsUsername, button: remove });
        }
      });
      control.append(edit, passwordControl, toggle, remove);
      if (account.requestStatus === "pending") {
        const approve = document.createElement("button");
        approve.className = "table-action approve";
        approve.textContent = "Approve Sub-Admin";
        approve.addEventListener("click", () => reviewSubAdminRequest(account, true, approve));
        const reject = document.createElement("button");
        reject.className = "table-action danger";
        reject.textContent = "Reject";
        reject.addEventListener("click", () => reviewSubAdminRequest(account, false, reject));
        control.prepend(approve, reject);
      }
    }
    row.appendChild(control);
    body.appendChild(row);
  });
}

function renderUsageAnalytics() {
  const body = $("#usageAnalyticsTable");
  if (!body) return;
  const records = state.usageAnalytics || [];
  $("#usageUserCount").textContent = `${records.length} USER${records.length === 1 ? "" : "S"}`;
  body.innerHTML = "";
  if (!records.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 8;
    cell.textContent = "Usage will appear after users start opening modules and files.";
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }
  records.forEach((record) => {
    const modules = Object.entries(record.modules || {}).sort((a, b) => (Number(b[1].seconds) || 0) - (Number(a[1].seconds) || 0));
    const [moduleName, moduleInfo = {}] = modules[0] || ["No module recorded", {}];
    const row = document.createElement("tr");
    const modulePercent = Number(record.totalSeconds) ? ((Number(moduleInfo.seconds) || 0) / Number(record.totalSeconds)) * 100 : 0;
    const values = [record.rank, record.fullName || record.username, String(record.role || "").replaceAll("_", " "), formatUsageTime(record.totalSeconds), Number(record.percentage || 0).toFixed(1), record.totalUses || 0, `${moduleName} · ${modulePercent.toFixed(1)}% · ${formatUsageTime(moduleInfo.seconds)}${moduleInfo.lastFile ? ` · ${moduleInfo.lastFile}` : ""}`, record.lastUsedAt ? formatDate(record.lastUsedAt) : "Never"];
    values.forEach((value, index) => {
      const cell = document.createElement("td");
      if (index === 0) cell.className = "usage-rank";
      if (index === 4) {
        cell.className = "usage-percent";
        cell.innerHTML = `${value}%<span><i style="width:${Math.min(100, Number(value))}%"></i></span>`;
      } else cell.textContent = value;
      cell.title = String(value);
      row.appendChild(cell);
    });
    body.appendChild(row);
  });
}

function renderDepartmentFiles() {
  const body = $("#departmentFilesTable");
  if (!body) return;
  $("#departmentFileCount").textContent = `${state.departmentFiles.length} FILE${state.departmentFiles.length === 1 ? "" : "S"}`;
  const lastSaved = state.departmentFiles[0];
  $("#uploadLastSaved").textContent = lastSaved
    ? `Last saved upload: ${lastSaved.fileSlot || lastSaved.file} · ${formatDate(lastSaved.uploadedAt)}`
    : "Last saved upload: No file saved yet";
  body.innerHTML = "";
  const files = state.departmentFiles.length ? state.departmentFiles : [{ empty: true }];
  files.forEach((file) => {
    const row = document.createElement("tr");
    const values = file.empty ? ["No uploaded files found.", "", "", "", "", "", ""] : [file.department, file.fileSlot || "Department File", file.file, file.extension || "File", formatFileSize(file.size), file.uploadedBy, formatDate(file.uploadedAt)];
    values.forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = value || "";
      cell.title = value || "";
      row.appendChild(cell);
    });
    body.appendChild(row);
  });
  renderUploadFileSlots();
}

function renderUploadFileSlots() {
  const select = $("#uploadFileSlot");
  if (!select) return;
  const scope = $("#uploadScope").value;
  const department = $("#uploadDepartment").value;
  $("#uploadDepartmentField").classList.toggle("hidden", scope !== "department");
  const previous = select.value;
  const isAdmin = Boolean(state.access?.permissions?.includes("system_control"));
  const fixed = scope === "material"
    ? ["LSMW BOM Upload File", "ZKNG Material Upload File", "Routing Upload File", "Fabric Material Code", "Finish Type Short Name", "Code Requirement File"]
    : ["Department File"];
  const area = scope === "material" ? "Material Related Files" : department;
  const automatic = state.departmentFiles
    .filter((file) => file.department === area && file.fileSlot)
    .map((file) => file.fileSlot)
    .filter((name) => String(name).toLowerCase() !== "routing upload format - material code");
  const names = [...new Set([...fixed, ...automatic])];
  select.replaceChildren(...names.map((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    return option;
  }));
  const addNew = document.createElement("option");
  addNew.value = "__new__";
  addNew.textContent = "+ Add New File Name";
  select.appendChild(addNew);
  select.value = names.includes(previous) || previous === "__new__" ? previous : names[0];
  $("#uploadNewFileSlotField").classList.toggle("hidden", select.value !== "__new__");
}

function renderUploadHeaderCheck(check) {
  const panel = $("#uploadHeaderCheck");
  panel.innerHTML = "";
  if (!check) {
    panel.classList.add("hidden");
    $("#saveDepartmentFile").disabled = false;
    return;
  }
  panel.classList.remove("hidden");
  const title = document.createElement("strong");
  title.textContent = `Header Column Check · Sheet: ${check.sheetName}`;
  const status = document.createElement("div");
  status.className = `upload-header-status${check.complete ? " ok" : ""}`;
  status.textContent = check.complete
    ? `All ${check.expected.length} required columns matched automatically.`
    : `${check.matched} of ${check.expected.length} matched. Select the missing columns manually.`;
  const mappedColumns = new Set((check.mappings || []).filter((index) => Number(index) >= 0));
  const extraColumns = check.headers
    .map((header, index) => ({ header, index }))
    .filter(({ index }) => !mappedColumns.has(index));
  const extras = document.createElement("div");
  extras.className = "upload-extra-columns";
  if (extraColumns.length) {
    const extraTitle = document.createElement("strong");
    extraTitle.textContent = "Extra columns · automatically removed when saved:";
    extras.appendChild(extraTitle);
    extraColumns.forEach(({ header, index }) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "upload-extra-chip";
      chip.dataset.columnIndex = String(index);
      chip.textContent = `${index < 26 ? String.fromCharCode(65 + index) : `Column ${index + 1}`}: ${header}  ×`;
      chip.title = "Remove this source column from the mapping list";
      chip.addEventListener("click", () => {
        document.querySelectorAll(`.upload-header-map option[value="${index}"]`).forEach((option) => {
          if (option.parentElement.value === String(index)) option.parentElement.value = "-1";
          option.remove();
        });
        chip.remove();
        updateUploadMappingStatus();
      });
      extras.appendChild(chip);
    });
  } else {
    extras.classList.add("clean");
    extras.textContent = "No extra columns found. The uploaded headers match the approved Material Code format.";
  }
  const grid = document.createElement("div");
  grid.className = "upload-header-grid";
  check.expected.forEach((expected, expectedIndex) => {
    const label = document.createElement("label");
    label.textContent = expected;
    const sample = document.createElement("div");
    sample.className = "upload-column-sample";
    const showSample = (columnIndex) => {
      if (Number(columnIndex) < 0) {
        sample.textContent = "Data: Select a source column to preview";
        sample.title = "";
        sample.classList.add("empty");
        return;
      }
      const columnData = check.columnSamples?.[Number(columnIndex)] || { filledRows: 0, samples: [] };
      const values = Array.isArray(columnData) ? columnData.map((value, index) => ({ row: index + 2, value })) : (columnData.samples || []);
      const filledRows = Array.isArray(columnData) ? values.length : Number(columnData.filledRows || 0);
      sample.textContent = values.length
        ? `${filledRows.toLocaleString()} filled rows · ${values.map((item) => `Row ${item.row}: ${item.value}`).join("  ·  ")}`
        : "0 filled rows · No value found in this column";
      sample.title = values.map((item) => `Row ${item.row}: ${item.value}`).join(" | ");
      sample.classList.toggle("empty", !values.length);
    };
    const select = document.createElement("select");
    select.className = "upload-header-map";
    select.dataset.expectedIndex = String(expectedIndex);
    const placeholder = document.createElement("option");
    placeholder.value = "-1";
    placeholder.textContent = "Select source column...";
    select.appendChild(placeholder);
    check.headers.forEach((header, columnIndex) => {
      const option = document.createElement("option");
      option.value = String(columnIndex);
      option.textContent = `${columnIndex < 26 ? String.fromCharCode(65 + columnIndex) : `Column ${columnIndex + 1}`}: ${header}`;
      select.appendChild(option);
    });
    select.value = String(check.mappings[expectedIndex] ?? -1);
    select.classList.toggle("missing", Number(select.value) < 0);
    showSample(select.value);
    select.addEventListener("change", () => {
      showSample(select.value);
      updateUploadMappingStatus();
    });
    grid.append(label, sample, select);
  });
  panel.append(title, status, extras, grid);
  updateUploadMappingStatus();
}

function uploadHeaderMappings() {
  return [...document.querySelectorAll(".upload-header-map")]
    .sort((a, b) => Number(a.dataset.expectedIndex) - Number(b.dataset.expectedIndex))
    .map((select) => Number(select.value));
}

function updateUploadMappingStatus() {
  const selects = [...document.querySelectorAll(".upload-header-map")];
  selects.forEach((select) => select.classList.toggle("missing", Number(select.value) < 0));
  const values = selects.map((select) => Number(select.value));
  const complete = !selects.length || (values.every((value) => value >= 0) && new Set(values).size === values.length);
  $("#saveDepartmentFile").disabled = !state.pendingDepartmentUpload || !complete;
  const status = $("#uploadHeaderCheck")?.querySelector(".upload-header-status");
  if (status && complete) {
    status.classList.add("ok");
    status.textContent = `All ${selects.length} required columns are mapped. Ready to Save.`;
  } else if (status && values.every((value) => value >= 0)) {
    status.classList.remove("ok");
    status.textContent = "A source column cannot be assigned more than once. Please correct the mapping.";
  }
}

async function controlSubAdmin(action) {
  const button = action.button;
  const request = { ...action };
  delete request.button;
  if (button) {
    button.disabled = true;
    button.dataset.originalText = button.textContent;
    button.textContent = action.mode === "delete" ? "Deleting..." : "Processing...";
  }
  try {
    const result = await window.erp.manageSubAdmin(request);
    showNotice(result.message, result.ok ? "success" : "warning");
    if (!result.ok) return;
    state.access.subAdmins = result.subAdmins;
    renderSubAdmins();
    renderDepartmentFiles();
    renderPasswordAccounts();
    renderAdminDashboard();
  } catch (error) {
    showNotice(`Account action failed: ${error.message}`, "warning");
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = button.dataset.originalText || "Process";
    }
  }
}

async function reviewSubAdminRequest(account, approved, actionButton = null) {
  if (actionButton) {
    actionButton.disabled = true;
    actionButton.textContent = approved ? "Approving..." : "Rejecting...";
  }
  try {
    const result = await window.erp.reviewSubAdminRequest({ username: account.employeeId, approved });
    showNotice(result.message, result.ok ? "success" : "warning");
    if (!result.ok) return;
    state.access.subAdmins = result.subAdmins;
    renderAccess();
  } catch (error) {
    showNotice(`Request action failed: ${error.message}`, "warning");
  } finally {
    if (actionButton) actionButton.disabled = false;
  }
}

async function changeAccountRole(account, role) {
  const result = await window.erp.changeAccountRole({
    username: account.employeeId,
    role,
  });
  showNotice(result.message, result.ok ? "success" : "warning");
  if (!result.ok) return result;
  state.access.subAdmins = result.subAdmins;
  renderAccess();
  return result;
}

function renderPasswordAccounts() {
  const select = $("#passwordAdminAccount");
  select.innerHTML = "";
  const mainOption = document.createElement("option");
  mainOption.value = state.access.username;
  mainOption.textContent = `Admin - ${state.access.fullName}`;
  select.appendChild(mainOption);
  (state.access.subAdmins || []).forEach((admin) => {
    const option = document.createElement("option");
    option.value = admin.windowsUsername;
    const roleLabel = ({ guest: "Guest User", standard_user: "Standard User", power_user: "Power User", master_user: "Master User", department_head: "Department Head", admin: "Admin" })[admin.role] || "Standard User";
    option.textContent = `${roleLabel} — ${admin.fullName} (${admin.department})`;
    select.appendChild(option);
  });
}

function renderProfileImage() {
  const hasImage = Boolean(state.profileImageDataUrl);
  $("#profileImage").classList.toggle("hidden", !hasImage);
  $("#profileInitials").classList.toggle("hidden", hasImage);
  if (hasImage) $("#profileImage").src = state.profileImageDataUrl;
}

function ensureMasterFileControls() {
  // The former top-level Admin Upload buttons are superseded by the secured
  // two-button Master File Control shown at the bottom of each section.
  $$(".admin-upload-btn").forEach((button) => button.remove());
  const sections = {
    materials: "Fabric Material Code",
    codeRequirement: "Code Requirement File",
    finishType: "Finish Type Short Name",
    routingUploadFormat: "Routing Upload File",
    zkngMaterialUpload: "ZKNG Material Upload File",
    lsmwBomUpload: "LSMW BOM Upload File",
    summary: "Circular Knitting Summary",
    diaWiseSummary: "Dia Wise Summary",
  };
  Object.entries(sections).forEach(([sectionId, label]) => {
    const section = document.getElementById(sectionId);
    if (!section || section.querySelector(`.master-file-control[data-master-section="${sectionId}"]`)) return;
    const control = document.createElement("div");
    control.className = "master-file-control hidden";
    control.dataset.masterSection = sectionId;
    control.innerHTML = `
      <div class="master-file-control-actions">
        <button type="button" class="secondary master-file-action" data-master-action="clear" data-master-section="${sectionId}">Clear</button>
        <button type="button" class="primary master-file-action" data-master-action="upload" data-master-section="${sectionId}">Upload</button>
      </div>`;
    section.appendChild(control);
  });
}

function renderAccess() {
  const isMainAdmin = state.access.role === "main_admin";
  const isFullAdmin = isMainAdmin || state.access.role === "admin";
  const isAuthorizedUploadRole = ["power_user", "master_user", "department_head", "admin", "main_admin"].includes(state.access.role) || (state.access.permissions || []).includes("upload");
  const canUpload = isAuthorizedUploadRole;
  const allowedFiles = Array.isArray(state.access.filePermissions) ? state.access.filePermissions : ["materials", "codeRequirement", "finishType", "routingUploadFormat", "zkngMaterialUpload", "lsmwBomUpload", "characteristicsUpload", "summary", "diaWiseSummary", "productionCalculation", "department"];

  $$(".admin-upload-btn, [data-admin-upload]").forEach((button) => {
    button.classList.toggle("hidden", !isAuthorizedUploadRole);
  });
  $$(".master-file-control").forEach((control) => {
    control.classList.toggle("hidden", !isFullAdmin);
  });
  if ($("#zkngAdminMaster")) {
    $("#zkngAdminMaster").classList.toggle("hidden", !isFullAdmin);
  }
  if ($("#openShare")) {
    $("#openShare").classList.toggle("hidden", !isFullAdmin);
  }
  $$('nav button[data-view], .submodule-nav').forEach((button) => {
    button.classList.remove("hidden");
  });
  $("#materialModuleToggle")?.classList.remove("hidden");
  $("#materialSubmodules")?.classList.remove("hidden");
  $("#zkngMaterialUploadNav")?.classList.remove("hidden");
  $("#lsmwBomUploadNav")?.classList.remove("hidden");
  $("#primarySummaryNav")?.classList.remove("hidden");
  $("#summarySubmodules")?.classList.remove("hidden");
  const roleLabels = {
    guest: "Guest User",
    user: "Standard User",
    standard_user: "Standard User",
    sub_admin: "Power User",
    power_user: "Power User",
    master_user: "Master User",
    department_head: "Department Head",
    admin: "Admin",
    main_admin: "Admin",
  };
  if ($("#profileName")) $("#profileName").textContent = state.access.fullName || state.access.username || "ZBOM User";
  if ($("#profileInitials")) $("#profileInitials").textContent = initials(state.access.fullName || state.access.username);
  if ($("#overviewUserName")) $("#overviewUserName").textContent = state.access.fullName || state.access.username || "User";
  const roleLabel = roleLabels[state.access.role] || "Standard User";
  const departmentLabel = state.access.departments?.length === state.departments?.length
    ? "All Departments"
    : state.access.departments?.[0] || "Not signed in";
  if ($("#profileRole")) $("#profileRole").textContent = `${roleLabel} · ${departmentLabel}`;
  if ($("#overviewUserRole")) $("#overviewUserRole").textContent = roleLabel;
  const logoutBtn = $("#logoutUser");
  if (logoutBtn) {
    logoutBtn.textContent = state.authenticated ? "Logout" : "Sign In / Create Account";
    logoutBtn.classList.toggle("signin-btn", !state.authenticated);
  }
  $("#changeProfilePhotoButton")?.classList.toggle("hidden", !state.authenticated);
  $("#administrationNav")?.classList.toggle("hidden", !(canUpload || isFullAdmin));
  if (!(canUpload || isFullAdmin)) {
    $("#administrationSubmodules")?.classList.add("hidden");
    $("#administrationNav")?.classList.remove("expanded");
  }
  $$(".main-admin-only").forEach((element) => element.classList.toggle("hidden", !isFullAdmin));
  $$('[data-admin-section]:not([data-admin-section="departmentUploadPanel"])').forEach((element) =>
    element.classList.toggle("hidden", !isFullAdmin));
  if ($("#adminRoleChip")) $("#adminRoleChip").textContent = roleLabel.toUpperCase();
  fillDepartmentSelect($("#uploadDepartment"), state.access.departments || []);
  fillDepartmentSelect($("#subAdminDepartment"), state.departments || []);
  fillDepartmentSelect($("#registerDepartment"), state.departments || []);
  fillDepartmentSelect($("#existingDepartment"), state.departments || []);
  renderDepartmentNavigation();
  const materialGroupAllowed = ["materials", "codeRequirement", "finishType", "routingUploadFormat", "zkngMaterialUpload", "lsmwBomUpload"].some((id) => allowedFiles.includes(id));
  $("#materialModuleToggle")?.classList.toggle("hidden", !materialGroupAllowed);
  $("#zkngMaterialUploadNav")?.classList.remove("hidden");
  $("#lsmwBomUploadNav")?.classList.remove("hidden");
  $("#characteristicsUploadNav")?.classList.remove("hidden");
  $("#primarySummaryNav")?.classList.toggle("hidden", !["summary", "diaWiseSummary", "productionCalculation"].some((id) => allowedFiles.includes(id)));
  $$('[data-view="department"]').forEach((button) => button.classList.toggle("hidden", !allowedFiles.includes("department")));
  renderSubAdmins();
  renderDepartmentFiles();
  if (isFullAdmin) renderPasswordAccounts();
  if (isFullAdmin) renderAdminDashboard();
  renderProfileImage();
  renderOverviewStatus();
}

async function refreshAccess() {
  const latest = await window.erp.getAccessState();
  state.access = latest.access;
  state.departments = latest.departments || [];
  state.departmentStatuses = latest.departmentStatuses || {};
  state.departmentFiles = latest.departmentFiles || [];
  state.usageAnalytics = latest.usageAnalytics || [];
  state.authenticated = latest.authenticated;
  if (latest.profileImageDataUrl !== undefined) {
    state.profileImagesByUser[state.access.username] = latest.profileImageDataUrl || "";
  }
  state.profileImageDataUrl = state.profileImagesByUser[state.access.username] || "";
  renderAccess();
  if (state.view === "department") renderDepartmentStatus(state.department);
  renderOverviewStatus();
}

async function checkSoftwareUpdate(silent = true) {
  try {
    const result = await window.erp.checkSoftwareUpdate();
    if (result.available) {
      $("#softwareUpdateTitle").textContent = `Update required · ZBOM ${result.manifest.version}`;
      $("#softwareUpdateText").textContent = `New features are ready. Click Update Now to install and restart · Published ${formatDate(result.manifest.publishedAt)}`;
      $("#softwareUpdate").classList.remove("hidden");
      if (!silent) showNotice(result.message, "success");
    } else {
      $("#softwareUpdate").classList.add("hidden");
      if (!silent) showNotice(result.message, result.ok ? "success" : "warning");
    }
    return result;
  } catch (error) {
    if (!silent) showNotice(`Update check failed: ${error.message}`, "warning");
    return null;
  }
}

async function updateData(force = false, silent = false) {
  if (state.updating) return null;
  state.updating = true;
  $("#updateNow").disabled = true;
  $("#updateNow").textContent = "Updating…";
  let updateResult = null;
  try {
    const result = await window.erp.update(force);
    updateResult = result;
    if (result.snapshot) {
      state.snapshot = result.snapshot;
      renderState();
    }
    await refreshAccess();
    setConnection(Boolean(result.ok), result.ok ? undefined : "Using saved offline data");
    if (!silent || result.changed) showNotice(result.message, result.ok ? "success" : "warning");
  } catch (error) {
    setConnection(false, "Using saved offline data");
    if (!silent) showNotice(`Update failed: ${error.message}`, "warning");
  } finally {
    state.updating = false;
    $("#updateNow").disabled = false;
    $("#updateNow").textContent = "Update Now";
  }
  return updateResult;
}

async function updateEverything() {
  const dataResult = await updateData(true, false);
  if (!dataResult) return;
  $("#updateNow").disabled = true;
  $("#updateNow").textContent = "Checking Software...";
  const software = await checkSoftwareUpdate(true);
  if (software?.available) {
    $("#updateNow").textContent = "Installing...";
    const install = await window.erp.installSoftwareUpdate();
    if (!install.ok) {
      showNotice(install.message || "Software update could not be installed.", "warning");
      $("#updateNow").disabled = false;
      $("#updateNow").textContent = "Update Now";
    }
    return;
  }
  $("#updateNow").disabled = false;
  $("#updateNow").textContent = "Update Now";
  showNotice(software?.ok ? "All data, permissions and software are up to date." : "Data updated. Software update folder is not available.", software?.ok ? "success" : "warning");
}

async function exportCurrent(format) {
  if (!state.authenticated) {
    showNotice("Create an account or sign in before downloading files.", "warning");
    openAdminLogin();
    return;
  }
  if (!(state.access.permissions || []).includes("export")) {
    showNotice("Guest User role has View Only access. Sign in as Standard User or higher for Data Export.", "warning");
    openAdminLogin();
    return;
  }
  const tableType = state.view === "summary"
    ? "summary"
    : state.view === "diaWiseSummary"
      ? "diaWiseSummary"
    : state.view === "finishType"
      ? "finishType"
      : "materials";
  const selected = [...(state.selectedRows[tableType] || [])];
  const materialSearchActive = tableType === "materials" && Boolean(
    $("#materialSearch").value.trim() ||
    state.materialValueSelections.size ||
    state.excelFilters.materials.size,
  );
  const rows = materialSearchActive
    ? state.filteredMaterials
    : selected.length
    ? selected
    : tableType === "diaWiseSummary"
      ? state.filteredDiaWiseSummary
    : tableType === "summary"
      ? state.filteredSummary
      : tableType === "finishType"
        ? state.filteredFinishType
        : state.filteredMaterials;
  if (tableType === "materials" && rows.length > 50000) {
    showNotice(
      `${rows.length.toLocaleString()} matching records found. Please make the search/filter more specific; one download can contain at most 50,000 records.`,
      "warning",
    );
    return;
  }
  const result = await window.erp.exportRows({
    rows,
    format,
    tableType,
    diaTotals: tableType === "diaWiseSummary" ? DIA_BALANCE_COLUMNS.map((column) => ({ label: column, value: rows.reduce((sum, row) => sum + diaNumber(row, column), 0) })) : undefined,
    knittingBalance: tableType === "diaWiseSummary" ? rows.reduce((sum, row) => sum + diaNumber(row, "Knitting Balance"), 0) : undefined,
    defaultName: tableType === "summary"
      ? `Zbom-${(state.departments[0] || "Circular Knitting").replace(/\s+/g, "")}`
      : tableType === "diaWiseSummary"
        ? "Zbom-Dia-Wise-Summary-Print"
      : tableType === "finishType"
        ? "Zbom-FinishTypeShortName"
        : "Zbom-Fabriccode",
  });
  if (result.ok) {
    state.lastDownloadedFile = result.path;
    $("#downloadedFileName").textContent = `${result.path.split(/[\\/]/).pop()} · ${rows.length.toLocaleString()} rows`;
    $("#downloadComplete").classList.remove("hidden");
    showNotice("Search results saved successfully.", "success");
  } else if (result.message) {
    showNotice(result.message, "warning");
  }
}

async function initialize() {
  ensureMasterFileControls();

  // Render default UI and activate views instantly (0.00ms delay!)
  renderAccess();
  renderState();
  renderCodeRequirementPreview();
  changeView("materials");

  // Asynchronous background state fetch
  window.erp.getState().then((initial) => {
    if (!initial) return;
    state.snapshot = initial.snapshot || state.snapshot;
    state.config = initial.config || state.config;
    state.appVersion = initial.appVersion || state.appVersion;
    state.access = initial.access || state.access;
    state.departments = initial.departments || state.departments;
    state.departmentStatuses = initial.departmentStatuses || state.departmentStatuses;
    state.departmentFiles = initial.departmentFiles || state.departmentFiles;
    state.usageAnalytics = initial.usageAnalytics || state.usageAnalytics;
    state.authenticated = initial.authenticated;
    state.profileImageDataUrl = initial.profileImageDataUrl || "";
    if (state.access?.username) state.profileImagesByUser[state.access.username] = state.profileImageDataUrl;
    state.codeRequirementPreview = initial.codeRequirementPreview || state.codeRequirementPreview;
    if ($("#currentAppVersion")) $("#currentAppVersion").textContent = `Current version ${initial.appVersion || "1.0.0"}`;
    if ($("#sharePath")) $("#sharePath").value = initial.config?.sharePath || "";
    if ($("#autoUpdate")) $("#autoUpdate").checked = initial.config?.autoUpdate !== false;

    renderAccess();
    renderState();
    renderCodeRequirementPreview();
  }).catch(() => {});

  // Non-blocking background data sync
  setConnection(false, "Checking shared folder…");
  updateData(false, true).catch(() => {});
  checkSoftwareUpdate(true).catch(() => {});

  window.setInterval(() => {
    if ($("#autoUpdate")?.checked) {
      updateData(false, true).catch(() => {});
      checkSoftwareUpdate(true).catch(() => {});
    }
  }, 60 * 1000);
}

async function renderSavedAccountOptions() {
  const group = $("#savedAccountsGroup");
  const select = $("#savedAccountSelect");
  if (!group || !select) return;
  
  let data = { savedAccounts: [], lastSelectedUsername: "" };
  try {
    if (window.erp && typeof window.erp.getSavedAccounts === "function") {
      data = await window.erp.getSavedAccounts();
    }
  } catch (err) {
    console.warn("getSavedAccounts error:", err);
  }

  const list = Array.isArray(data.savedAccounts) ? data.savedAccounts : [];
  if (!list.length) {
    group.classList.add("hidden");
    return;
  }

  group.classList.remove("hidden");
  select.innerHTML = `<option value="">-- Choose a saved account on this computer --</option>`;
  
  let preselected = false;
  list.forEach((acc) => {
    const opt = document.createElement("option");
    opt.value = acc.username;
    opt.dataset.password = acc.password || "";
    const roleTitle = ({ guest: "Guest User", standard_user: "Standard User", power_user: "Power User", master_user: "Master User", department_head: "Department Head", admin: "Admin", main_admin: "Main Admin" })[acc.role] || "Account";
    opt.textContent = `👤 ${acc.fullName} (${acc.username}) — ${roleTitle}`;
    if (acc.username === data.lastSelectedUsername || !preselected) {
      opt.selected = true;
      preselected = true;
    }
    select.appendChild(opt);
  });

  const selectedOpt = select.options[select.selectedIndex];
  if (selectedOpt && selectedOpt.value) {
    $("#adminLoginUsername").value = selectedOpt.value;
    $("#adminLoginPassword").value = selectedOpt.dataset.password || "";
  }
}

async function openAdminLogin() {
  const overlay = $("#adminLoginOverlay");
  if (overlay) overlay.classList.remove("hidden");

  const usernameInput = $("#adminLoginUsername");
  const passwordInput = $("#adminLoginPassword");
  const rememberCheckbox = $("#rememberLogin");
  const errorMsg = $("#adminLoginError");

  if (rememberCheckbox) rememberCheckbox.checked = true;
  if (errorMsg) errorMsg.classList.add("hidden");

  renderSavedAccountOptions().catch(() => {});

  setTimeout(() => {
    if (usernameInput) (usernameInput.value ? passwordInput : usernameInput).focus();
  }, 10);
}

function openAssignedDepartment() {
  const departments = Array.isArray(state.access.departments) ? state.access.departments.filter((name) => name && name !== "All Departments") : [];
  if (departments.length !== 1) {
    changeView("dashboard");
    return;
  }
  state.department = departments[0];
  if (state.department === "Circular Knitting") {
    changeView("summary");
  } else {
    changeView("department");
  }
}

async function confirmAdminLogin() {
  const username = $("#adminLoginUsername").value.trim();
  const password = $("#adminLoginPassword").value;
  const signInButton = $("#confirmAdminLogin");
  signInButton.disabled = true;
  signInButton.textContent = "Signing In...";
  let result;
  try {
    result = await window.erp.adminLogin({ username, password, remember: $("#rememberLogin").checked });
  } catch (error) {
    $("#adminLoginError").textContent = `Sign in failed: ${error.message}`;
    $("#adminLoginError").classList.remove("hidden");
    return;
  } finally {
    signInButton.disabled = false;
    signInButton.textContent = "Sign In";
  }
  if (!result.ok) {
    $("#adminLoginError").textContent = result.message;
    $("#adminLoginError").classList.remove("hidden");
    return;
  }
  // Instant UI state update on the spot (0.00ms delay!)
  state.authenticated = true;
  if (result.access) state.access = result.access;
  $("#adminLoginOverlay").classList.add("hidden");
  $("#adminLoginError").classList.add("hidden");
  renderAccess();
  openAssignedDepartment();
  showNotice(result.message, "success");

  refreshAccess().catch((error) => {
    console.warn("Post-login refresh error:", error);
  });
}

function openRegistration() {
  $("#adminLoginOverlay").classList.add("hidden");
  $("#registrationOverlay").classList.remove("hidden");
  $("#registrationError").classList.add("hidden");
  window.setTimeout(() => $("#registerFullName").focus(), 50);
}

function closeRegistration() {
  $("#registrationOverlay").classList.add("hidden");
  $("#adminLoginOverlay").classList.remove("hidden");
}

async function submitRegistration() {
  const password = $("#registerPassword").value;
  if (password !== $("#registerConfirmPassword").value) {
    $("#registrationError").textContent = "Password confirmation does not match.";
    $("#registrationError").classList.remove("hidden");
    return;
  }
  const record = {
    fullName: $("#registerFullName").value.trim(),
    employeeId: $("#registerEmployeeId").value.trim(),
    mobile: $("#registerMobile").value.trim(),
    telephone: $("#registerTelephone").value.trim(),
    designation: $("#registerDesignation").value.trim(),
    department: $("#registerDepartment").value,
    company: $("#registerCompany").value,
    password,
    requestSubAdmin: $("#registerRequestSubAdmin").checked,
  };
  $("#submitRegistration").disabled = true;
  $("#submitRegistration").textContent = "Creating...";
  let result;
  try {
    result = await window.erp.registerAccount(record);
  } catch (error) {
    result = { ok: false, message: `Account could not be created: ${error.message}` };
  } finally {
    $("#submitRegistration").disabled = false;
    $("#submitRegistration").textContent = "Create Account";
  }
  if (!result.ok) {
    $("#registrationError").textContent = result.message;
    $("#registrationError").classList.remove("hidden");
    return;
  }
  const login = await window.erp.adminLogin({ username: result.username, password, remember: true });
  if (login.ok) {
    await refreshAccess();
    $("#registrationOverlay").classList.add("hidden");
    $("#adminLoginOverlay").classList.add("hidden");
    openAssignedDepartment();
    showNotice(`${result.message} You are now signed in and can download files.`, "success");
    return;
  }
  $("#registrationOverlay").classList.add("hidden");
  $("#adminLoginOverlay").classList.remove("hidden");
  $("#adminLoginUsername").value = result.username;
  showNotice(`${result.message} Please sign in to continue.`, "success");
  $("#adminLoginUsername").focus();
}

async function downloadResource(resourceId) {
  if (!state.authenticated) {
    showNotice("Create an account or sign in before downloading this file.", "warning");
    openAdminLogin();
    return;
  }
  const result = await window.erp.downloadResource(resourceId);
  if (result.canceled) return;
  showNotice(result.message, result.ok ? "success" : "warning");
  if (!result.ok) return;
  state.lastDownloadedFile = result.path;
  $("#downloadedFileName").textContent = result.path.split(/[\\/]/).pop();
  $("#downloadComplete").classList.remove("hidden");
}

async function triggerAdminUpload(slot) {
  const isAuthorizedUploadRole = ["power_user", "master_user", "department_head", "admin", "main_admin"].includes(state.access.role) || (state.access.permissions || []).includes("upload");
  if (!isAuthorizedUploadRole) {
    showNotice("Access denied. Only Power User, Master User, Department Head, and Admin roles can use File Upload & Admin options.", "warning");
    openAdminLogin();
    return;
  }
  if (slot === "zkngMaterialUpload") {
    const btn = $("#uploadZkngMaster");
    if (btn) btn.click();
    else showNotice("Admin ZKNG Master Control active.", "info");
    return;
  }
  if (slot === "lsmwBomUpload") {
    const btn = $("#uploadLsmwBomMaster");
    if (btn) btn.click();
    else showNotice("Admin LSMW BOM Master Control active.", "info");
    return;
  }
  if (slot === "routingUploadFormat") {
    const btn = $("#uploadRoutingMaster");
    if (btn) btn.click();
    else showNotice("Admin Routing Master Control active.", "info");
    return;
  }
  
  const dept = (state.access.departments && state.access.departments.length > 0 && state.access.departments[0] !== "All Departments")
    ? state.access.departments[0]
    : (state.departments?.[0] || "Circular Knitting");
    
  const result = await window.erp.uploadDepartmentFile({ department: dept, slot });
  if (result && result.canceled) return;
  if (result && result.ok) {
    showNotice(result.message, "success");
    await refreshAccess();
  } else if (result && result.message) {
    showNotice(result.message, "warning");
  }
}

document.addEventListener("click", async (event) => {
  const button = event.target.closest(".master-file-action");
  if (!button) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  const allowed = ["power_user", "master_user", "department_head", "admin", "main_admin"].includes(state.access.role);
  if (!allowed) {
    showNotice("Only Power User, Master User, Department Head, and Admin can manage master files.", "warning");
    return;
  }
  const section = button.dataset.masterSection;
  const action = button.dataset.masterAction;
  const control = button.closest(".master-file-control");
  const status = control?.querySelector(".master-file-status");
  if (action === "clear" && !window.confirm("Clear/remove the old master file and its cached data for this section?")) return;
  button.disabled = true;
  const originalLabel = button.textContent;
  button.textContent = action === "clear" ? "Clearing..." : "Uploading...";
  if (status) {
    status.className = "master-file-status processing";
    status.textContent = action === "clear" ? "Processing: clearing old master file..." : "Processing file: copying, validating and importing...";
  }
  try {
    const result = action === "clear"
      ? await window.erp.clearMasterSection(section)
      : await window.erp.uploadMasterSection(section);
    if (result?.canceled) {
      if (status) {
        status.className = "master-file-status";
        status.textContent = "Upload canceled. No file was changed.";
      }
      return;
    }
    showNotice(result?.message || "Master-file action completed.", result?.ok ? "success" : "warning");
    if (status) {
      status.className = `master-file-status ${result?.ok ? "completed" : "failed"}`;
      status.textContent = result?.ok
        ? (action === "clear" ? "Completed: old master file and data cleared." : "Upload completed: new master file imported and activated.")
        : `Upload failed: ${result?.message || "The file could not be processed."}`;
    }
    if (result?.snapshot) {
      state.snapshot = result.snapshot;
      renderState();
      renderCodeRequirementPreview();
    } else if (result?.ok && action === "clear") {
      const refreshed = await window.erp.update(true);
      if (refreshed?.snapshot) {
        state.snapshot = refreshed.snapshot;
        renderState();
      }
    }
  } catch (error) {
    if (status) {
      status.className = "master-file-status failed";
      status.textContent = `Upload failed: ${error.message}`;
    }
    showNotice(`Master file could not be processed: ${error.message}`, "warning");
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
});

document.addEventListener("click", async (event) => {
  const btn = event.target.closest("button, aside button, nav button, .submodule-nav, [data-view], [data-go], [data-admin-upload]");
  if (!btn) return;
  if (btn.dataset.adminUpload) {
    await triggerAdminUpload(btn.dataset.adminUpload);
    return;
  }
  const targetView = btn.dataset.view;
  if (btn.dataset.go) {
    changeView(btn.dataset.go);
    return;
  }
  if (btn.dataset.resource) {
    await downloadResource(btn.dataset.resource);
    return;
  }
  if (targetView === "administration" && !state.authenticated) {
    openAdminLogin();
    return;
  }
  if (btn.dataset.department) {
    state.department = btn.dataset.department;
    openDepartmentView(btn.dataset.department);
    return;
  }
  if (btn.id === "materialModuleToggle") {
    const group = $("#materialSubmodules");
    const expanded = Boolean(group && group.classList.contains("hidden"));
    setMaterialMenuExpanded(expanded);
    return;
  }
  if (btn.id === "primarySummaryNav") {
    const group = $("#summarySubmodules");
    if (group) group.classList.remove("hidden");
    btn.classList.add("expanded");
  }
  if (btn.id === "administrationNav") {
    const group = $("#administrationSubmodules");
    if (group) {
      const opening = group.classList.contains("hidden");
      group.classList.toggle("hidden", !opening);
      btn.classList.toggle("expanded", opening);
    }
  }
  if (targetView) {
    changeView(targetView);
  }
  if (btn.dataset.adminSection) {
    window.setTimeout(() => document.getElementById(btn.dataset.adminSection)?.scrollIntoView({
      behavior: "smooth",
      block: "start",
    }), 30);
  }
});
$("#materialSearch")?.addEventListener("input", debounce(() => refreshFilteredTable("materials")));
$("#diaSummarySearch")?.addEventListener("input", debounce(renderDiaWiseSummary, 120));
// Production Calculation listeners
$("#prodFabricSelect")?.addEventListener("change", (event) => { prodCalc.fabric = event.target.value; saveProdCalc(); renderProductionCalculation(); });
$("#prodShiftHours")?.addEventListener("input", debounce(() => { const v = prodNum($("#prodShiftHours").value); if (Number.isFinite(v) && v > 0) { prodCalc.shiftHours = v; saveProdCalc(); renderProdTable(); renderProdReverse(); } }, 250));
$("#prodEfficiency")?.addEventListener("input", debounce(() => { const v = prodNum($("#prodEfficiency").value); if (Number.isFinite(v) && v > 0 && v <= 100) { prodCalc.efficiency = v; saveProdCalc(); renderProdTable(); renderProdReverse(); } }, 250));
$("#prodAddFabricToggle")?.addEventListener("click", () => { $("#prodAddFabricPanel")?.classList.toggle("hidden"); $("#prodNewFabricName")?.focus(); });
$("#prodAddFabricBtn")?.addEventListener("click", () => prodAddFabric($("#prodNewFabricName")?.value));
$("#prodNewFabricName")?.addEventListener("keydown", (event) => { if (event.key === "Enter") prodAddFabric(event.target.value); });
$$("#prodAddFabricPanel [data-add-fabric]").forEach((btn) => btn.addEventListener("click", () => prodAddFabric(btn.dataset.addFabric)));
$("#prodAddMachine")?.addEventListener("click", () => { const rows = prodRows(); rows.push({ ...PROD_SAMPLE_ROW, machine: String(rows.length + 1) }); saveProdCalc(); renderProdTable(); renderProdReverse(); });
$("#prodReset")?.addEventListener("click", () => { prodCalc.rows[prodCalc.fabric] = [{ ...PROD_SAMPLE_ROW }]; saveProdCalc(); renderProdTable(); renderProdReverse(); });
$("#prodTableBody")?.addEventListener("input", (event) => {
  const input = event.target.closest("input[data-field]");
  if (!input) return;
  const tr = input.closest("tr[data-row]");
  const row = prodRows()[Number(tr?.dataset.row)];
  if (!row) return;
  row[input.dataset.field] = input.value;
  saveProdCalc();
  prodRecalcRow(Number(tr.dataset.row));
  if (input.dataset.field === "machine") renderProdReverse();
});
$("#prodTableBody")?.addEventListener("click", (event) => {
  const btn = event.target.closest(".prod-remove-row");
  if (!btn) return;
  const rows = prodRows();
  if (rows.length <= 1) { showNotice("At least one machine row is needed.", "warning"); return; }
  rows.splice(Number(btn.dataset.row), 1);
  saveProdCalc();
  renderProdTable();
  renderProdReverse();
});
$("#prodCounterMachine")?.addEventListener("change", prodUpdateReverseResults);
$("#prodProductionMachine")?.addEventListener("change", prodUpdateReverseResults);
$("#prodTargetProduction")?.addEventListener("input", debounce(() => { const v = prodNum($("#prodTargetProduction").value); if (Number.isFinite(v)) { prodCalc.targetProduction = v; saveProdCalc(); } prodUpdateReverseResults(); }, 200));
$("#prodCounterInput")?.addEventListener("input", debounce(() => { const v = prodNum($("#prodCounterInput").value); if (Number.isFinite(v)) { prodCalc.counterInput = v; saveProdCalc(); } prodUpdateReverseResults(); }, 200));

$$("[data-material-column]").forEach((input) =>
  input.addEventListener("input", debounce(() => refreshFilteredTable("materials"))),
);
$("#toggleMaterialFilters")?.addEventListener("click", () => {
  const panel = $("#materialColumnFilters");
  const opening = panel.classList.contains("hidden");
  panel.classList.toggle("hidden", !opening);
  $("#materials").classList.toggle("filters-open", opening);
  $("#toggleMaterialFilters").classList.toggle("active", opening);
  $("#toggleMaterialFilters").textContent = opening ? "☷ Hide Filter" : "☷ Filter";
});
$("#clearMaterialFilters")?.addEventListener("click", () => {
  $("#materialSearch").value = "";
  $$("[data-material-column]").forEach((input) => {
    input.value = "";
  });
  state.excelFilters.materials.clear();
  state.materialValueSelections.clear();
  refreshFilteredTable("materials");
});
$("#clearAllMaterialFilters")?.addEventListener("click", () => {
  $("#materialSearch").value = "";
  $$("[data-material-column]").forEach((input) => {
    input.value = "";
  });
  state.excelFilters.materials.clear();
  state.materialValueSelections.clear();
  closeExcelFilter();
  refreshFilteredTable("materials");
  showNotice("All Fabric Material filters cleared.", "success");
});
$("#clearAllSummaryFilters")?.addEventListener("click", () => {
  $$("[data-summary-column]").forEach((input) => {
    input.value = "";
  });
  state.summaryCategorySelected = new Set(SUMMARY_CATEGORIES);
  renderSummaryCategoryOptions();
  state.excelFilters.summary.clear();
  state.summaryValueSelections.clear();
  closeExcelFilter();
  refreshFilteredTable("summary");
  showNotice(`All ${state.departments[0] || "Circular Knitting"} filters cleared.`, "success");
});
$("#clearFinishTypeFilters")?.addEventListener("click", () => {
  $("#finishTypeSearch").value = "";
  state.excelFilters.finishType.clear();
  closeExcelFilter();
  refreshFilteredTable("finishType");
  showNotice("All Finish Type Short Name filters cleared.", "success");
});
document.addEventListener("click", (event) => {
  const filterButton = event.target.closest(".table-filter-button");
  if (filterButton) {
    event.stopPropagation();
    openExcelFilter(filterButton);
    return;
  }
  if (!event.target.closest("#excelFilterMenu")) closeExcelFilter();
});
$$(".summary-field-filter-button").forEach((button) => {
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    openExcelFilter(button);
  });
});
$("#closeExcelFilter")?.addEventListener("click", closeExcelFilter);
$("#cancelExcelFilter")?.addEventListener("click", closeExcelFilter);
$("#excelValueSearch")?.addEventListener("input", renderExcelValueList);
$("#excelSelectAll")?.addEventListener("change", () => {
  if (!state.openExcelFilter) return;
  const { allValues, pendingExcluded } = state.openExcelFilter;
  pendingExcluded.clear();
  if (!$("#excelSelectAll")?.checked) {
    allValues.forEach((value) => pendingExcluded.add(value));
  }
  renderExcelValueList();
});
$("#excelPendingSelection")?.addEventListener("click", (event) => {
  const remove = event.target.closest("[data-pending-filter-value]");
  if (!remove || !state.openExcelFilter) return;
  const value = remove.dataset.pendingFilterValue;
  if (remove.dataset.pendingFilterMode === "included") {
    state.openExcelFilter.pendingExcluded.add(value);
  } else {
    state.openExcelFilter.pendingExcluded.delete(value);
  }
  renderExcelValueList();
});
$("#applyExcelFilter")?.addEventListener("click", () => {
  if (!state.openExcelFilter) return;
  const { tableType, column, pendingExcluded } = state.openExcelFilter;
  if (pendingExcluded.size) state.excelFilters[tableType].set(column, new Set(pendingExcluded));
  else state.excelFilters[tableType].delete(column);
  closeExcelFilter();
  refreshFilteredTable(tableType);
});
$$("[data-summary-column]").forEach((input) =>
  input.addEventListener("input", debounce(() => refreshFilteredTable("summary"), 150)),
);
$("#summaryCategoryButton")?.addEventListener("click", (event) => {
  event.stopPropagation();
  const menu = $("#summaryCategoryMenu");
  const opening = menu.classList.contains("hidden");
  menu.classList.toggle("hidden");
  if (opening) {
    $("#summaryCategorySearch").value = "";
    $$("#summaryCategoryOptions label").forEach((label) => label.classList.remove("hidden"));
    $("#summaryCategorySearch").focus();
  }
});
$("#summaryCategoryMenu")?.addEventListener("click", (event) => event.stopPropagation());
$("#summaryCategoryToggleAll")?.addEventListener("click", () => {
  const checkboxes = $$("#summaryCategoryOptions input");
  const allChecked = checkboxes.every((checkbox) => checkbox.checked);
  checkboxes.forEach((checkbox) => {
    checkbox.checked = !allChecked;
  });
  updateSummaryCategoryToggleLabel();
});
$("#summaryCategoryOptions")?.addEventListener("change", updateSummaryCategoryToggleLabel);
$("#summaryCategorySearch")?.addEventListener("input", () => {
  const term = normalizedSearchText($("#summaryCategorySearch").value);
  $$("#summaryCategoryOptions label").forEach((label) => {
    label.classList.toggle("hidden", Boolean(term) && !normalizedSearchText(label.textContent).includes(term));
  });
});
$("#summaryCategoryApply")?.addEventListener("click", () => {
  state.summaryCategorySelected = new Set(
    $$("#summaryCategoryOptions input:checked").map((input) => input.value),
  );
  renderSummaryCategoryOptions();
  $("#summaryCategoryMenu").classList.add("hidden");
  refreshFilteredTable("summary");
});
$("#summaryActiveFilterChips")?.addEventListener("click", (event) => {
  const remove = event.target.closest("[data-summary-filter-key]");
  if (!remove) return;
  const key = remove.dataset.summaryFilterKey;
  if (key === "__category__") {
    if (remove.dataset.summaryFilterValue) {
      state.summaryCategorySelected.delete(remove.dataset.summaryFilterValue);
    } else {
      state.summaryCategorySelected = new Set(SUMMARY_CATEGORIES);
    }
    renderSummaryCategoryOptions();
  } else if (remove.dataset.summaryFilterValue) {
    const selected = state.summaryValueSelections.get(key);
    selected?.delete(remove.dataset.summaryFilterValue);
    if (!selected?.size) state.summaryValueSelections.delete(key);
    pruneDependentSummarySelections(key);
  } else {
    state.summaryValueSelections.delete(key);
  }
  refreshFilteredTable("summary");
});
$("#materialActiveFilterChips")?.addEventListener("click", (event) => {
  const remove = event.target.closest("[data-material-filter-column]");
  if (!remove) return;
  const column = remove.dataset.materialFilterColumn;
  const selected = state.materialValueSelections.get(column);
  selected?.delete(remove.dataset.materialFilterValue);
  if (!selected?.size) state.materialValueSelections.delete(column);
  pruneDependentMaterialSelections(column);
  refreshFilteredTable("materials");
});
document.addEventListener("click", (event) => {
  if (!event.target.closest(".summary-multi-filter")) $("#summaryCategoryMenu").classList.add("hidden");
});
renderSummaryCategoryOptions();
$("#finishTypeSearch")?.addEventListener("input", debounce(() => refreshFilteredTable("finishType"), 120));
$("#updateNow")?.addEventListener("click", updateEverything);
$("#forceUpdate")?.addEventListener("click", () => updateData(true));
$("#browseFolder")?.addEventListener("click", async () => {
  const folder = await window.erp.chooseFolder();
  if (folder) $("#sharePath").value = folder;
});
$("#saveSettings")?.addEventListener("click", async () => {
  state.config = await window.erp.saveSettings({
    sharePath: $("#sharePath").value.trim(),
    autoUpdate: $("#autoUpdate").checked,
  });
  showNotice("Settings saved on this computer.", "success");
});
$("#openShare")?.addEventListener("click", () => window.erp.openPath($("#sharePath").value.trim()));
$("#openDownloadedFile")?.addEventListener("click", () => {
  if (state.lastDownloadedFile) window.erp.openPath(state.lastDownloadedFile);
  $("#downloadComplete").classList.add("hidden");
});
$("#closeDownloadNotice")?.addEventListener("click", () => $("#downloadComplete").classList.add("hidden"));
$("#checkSoftwareUpdate")?.addEventListener("click", () => checkSoftwareUpdate(false));
$("#installSoftwareUpdate")?.addEventListener("click", async () => {
  $("#installSoftwareUpdate").disabled = true;
  $("#installSoftwareUpdate").textContent = "Verifying…";
  const result = await window.erp.installSoftwareUpdate();
  if (!result.ok) {
    showNotice(result.message, "warning");
    $("#installSoftwareUpdate").disabled = false;
    $("#installSoftwareUpdate").textContent = "Update Now & Restart";
  }
});
$("#publishSoftwareUpdate")?.addEventListener("click", async () => {
  const version = $("#publishVersion").value.trim();
  const result = await window.erp.publishSoftwareUpdate(version);
  if (result.canceled) return;
  showNotice(result.message, result.ok ? "success" : "warning");
  if (result.ok) checkSoftwareUpdate(true);
});
$("#uploadScope")?.addEventListener("change", () => {
  const departmentScope = $("#uploadScope").value === "department";
  $("#uploadDepartmentField").classList.toggle("hidden", !departmentScope);
  renderUploadFileSlots();
});
$("#uploadDepartment")?.addEventListener("change", renderUploadFileSlots);
$("#uploadFileSlot")?.addEventListener("change", () => {
  const custom = $("#uploadFileSlot").value === "__new__";
  $("#uploadNewFileSlotField").classList.toggle("hidden", !custom);
  if (custom) $("#uploadNewFileSlot").focus();
});
$("#uploadDepartmentFile")?.addEventListener("click", async () => {
  const scope = $("#uploadScope").value;
  const department = $("#uploadDepartment").value;
  const selectedSlot = $("#uploadFileSlot").value;
  const fileSlot = selectedSlot === "__new__" ? $("#uploadNewFileSlot").value.trim() : selectedSlot.trim();
  if (scope === "department" && !department) {
    showNotice("No department upload permission is assigned to this user.", "warning");
    return;
  }
  if (!fileSlot) {
    showNotice("Enter a separate File Slot / Document Name first.", "warning");
    (selectedSlot === "__new__" ? $("#uploadNewFileSlot") : $("#uploadFileSlot")).focus();
    return;
  }
  $("#uploadDepartmentFile").disabled = true;
  $("#uploadDepartmentFile").textContent = "Uploading…";
  try {
    const result = await window.erp.uploadDepartmentFile({ department, fileSlot, scope });
    if (result.canceled) return;
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok && result.pending) {
      state.pendingDepartmentUpload = { pendingId: result.pendingId, scope, department, fileSlot, headerCheck: result.headerCheck };
      $("#uploadResult").textContent = `${result.file} · Selected for ${fileSlot} · Not saved yet`;
      $("#uploadResult").classList.remove("hidden", "completed");
      renderUploadHeaderCheck(result.headerCheck);
    }
  } catch (error) {
    showNotice(`Upload failed: ${error.message}`, "warning");
  } finally {
    $("#uploadDepartmentFile").disabled = false;
    $("#uploadDepartmentFile").textContent = "Select One File & Upload";
  }
});
$("#saveDepartmentFile")?.addEventListener("click", async () => {
  const pending = state.pendingDepartmentUpload;
  if (!pending) {
    showNotice("Select a file before saving.", "warning");
    return;
  }
  const button = $("#saveDepartmentFile");
  button.disabled = true;
  button.textContent = "Saving...";
  try {
    const result = await window.erp.saveDepartmentFile({ pendingId: pending.pendingId, mappings: uploadHeaderMappings() });
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok) {
      if (pending.scope === "department") state.departmentStatuses[pending.department] = result.manifest;
      $("#uploadResult").textContent = `${result.manifest.file} · Saved successfully · Upload completed`;
      $("#uploadResult").classList.add("completed");
      $("#uploadNewFileSlot").value = "";
      state.pendingDepartmentUpload = null;
      await refreshAccess();
      $("#uploadLastSaved").textContent = `Last saved upload: ${result.manifest.fileSlot || result.manifest.file} · ${formatDate(result.manifest.uploadedAt)}`;
      renderUploadHeaderCheck(null);
    }
  } catch (error) {
    showNotice(`Save failed: ${error.message}`, "warning");
  } finally {
    button.disabled = !state.pendingDepartmentUpload;
    button.textContent = "Save";
  }
});
$("#saveSubAdmin")?.addEventListener("click", async (event) => {
  event.preventDefault();
  const saveButton = $("#saveSubAdmin");
  const record = {
    fullName: $("#subAdminName").value.trim(),
    employeeId: $("#subAdminEmployeeId").value.trim(),
    mobile: $("#subAdminMobile").value.trim(),
    telephone: $("#subAdminTelephone").value.trim(),
    designation: $("#subAdminDesignation").value.trim(),
    company: $("#subAdminCompany").value,
    role: $("#subAdminRole").value,
    windowsUsername: $("#subAdminEmployeeId").value.trim(),
    department: $("#subAdminDepartment").value,
    password: $("#subAdminPassword").value,
    filePermissions: $$("#subAdminFilePermissions input:checked").map((input) => input.value),
  };
  if (!record.fullName || !record.employeeId || !record.department) {
    showNotice("Full Name, Employee ID, and Department are required.", "warning");
    return;
  }
  saveButton.disabled = true;
  saveButton.textContent = "Saving Account...";
  try {
    const result = await window.erp.saveSubAdmin(record);
    showNotice(result.message, result.ok ? "success" : "warning");
    if (!result.ok) return;
    state.access.subAdmins = result.subAdmins;
    $$("#subAdminFilePermissions input").forEach((input) => { input.checked = true; });
    ["subAdminName", "subAdminEmployeeId", "subAdminMobile", "subAdminTelephone", "subAdminDesignation", "subAdminPassword"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    renderAccess();
  } catch (error) {
    showNotice(`Sub Admin could not be saved: ${error.message}`, "warning");
  } finally {
    saveButton.disabled = false;
    saveButton.textContent = "Add / Update Account";
  }
});
$("#subAdminRole")?.addEventListener("change", () => {
  const isAdminRole = ["admin", "master_user"].includes($("#subAdminRole").value);
  $("#subAdminDepartment").disabled = isAdminRole;
  $("#subAdminDepartment").title = isAdminRole ? "Admin has access to all departments." : "";
});
async function submitDepartmentChange(mode) {
  const oldName = $("#existingDepartment").value;
  const newName = $("#departmentName").value.trim();
  try {
    const result = await window.erp.manageDepartment({ mode, oldName, newName });
    showNotice(result.message, result.ok ? "success" : "warning");
    if (!result.ok) return;
    if (state.department === oldName) state.department = newName;
    state.departments = result.departments;
    state.access.allDepartments = result.departments;
    state.access.subAdmins = result.subAdmins;
    if (state.access.role === "main_admin") state.access.departments = result.departments;
    $("#departmentName").value = "";
    renderAccess();
  } catch (error) {
    showNotice(`Department could not be updated: ${error.message}`, "warning");
  }
}
function triggerReloadApp() {
  if (window.erp && typeof window.erp.reloadApp === "function") {
    window.erp.reloadApp();
  } else {
    window.location.reload();
  }
}

$("#reloadApp")?.addEventListener("click", triggerReloadApp);
$("#reloadAppSettings")?.addEventListener("click", triggerReloadApp);

window.addEventListener("keydown", (event) => {
  if (event.key === "F5" || (event.ctrlKey && (event.key === "r" || event.key === "R"))) {
    event.preventDefault();
    triggerReloadApp();
  }
});

$("#addDepartment")?.addEventListener("click", () => submitDepartmentChange("add"));
$("#renameDepartment")?.addEventListener("click", () => submitDepartmentChange("rename"));
$("#confirmAdminLogin")?.addEventListener("click", confirmAdminLogin);
$("#googleSignIn")?.addEventListener("click", () => { window.location.href = "/api/auth/google"; });
$("#openRegistration")?.addEventListener("click", openRegistration);
$("#cancelRegistration")?.addEventListener("click", closeRegistration);
$("#submitRegistration")?.addEventListener("click", submitRegistration);
$("#cancelAdminLogin")?.addEventListener("click", () => $("#adminLoginOverlay").classList.add("hidden"));
$("#cancelAccessRestricted")?.addEventListener("click", () => $("#accessRestrictedOverlay").classList.add("hidden"));
$("#goToLoginFromRestricted")?.addEventListener("click", () => {
  $("#accessRestrictedOverlay").classList.add("hidden");
  openAdminLogin();
});
$("#adminLoginPassword")?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") $("#adminLoginPassword").focus();
});
$("#savedAccountSelect")?.addEventListener("change", (event) => {
  const select = event.target;
  const opt = select.options[select.selectedIndex];
  if (opt && opt.value) {
    $("#adminLoginUsername").value = opt.value;
  } else {
    $("#adminLoginUsername").value = "";
  }
});

$("#forgetSavedAccount")?.addEventListener("click", async () => {
  const username = $("#adminLoginUsername").value.trim();
  if (!username) return;
  if (window.confirm(`Remove saved login for ${username} on this computer?`)) {
    if (window.erp && typeof window.erp.removeSavedAccount === "function") {
      await window.erp.removeSavedAccount(username);
    }
    $("#adminLoginUsername").value = "";
    await renderSavedAccountOptions();
    showNotice(`Saved account ${username} removed from this computer.`, "info");
  }
});

$("#adminLoginUsername")?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") confirmAdminLogin();
});
$$(".toggle-password").forEach((button) => button.addEventListener("click", () => {
  const input = document.getElementById(button.dataset.passwordTarget);
  const showing = input.type === "text";
  input.type = showing ? "password" : "text";
  button.classList.toggle("showing", !showing);
  button.title = showing ? "Show password" : "Hide password";
  button.setAttribute("aria-label", button.title);
}));
$("#setAdminPassword")?.addEventListener("click", async () => {
  const password = $("#newAdminPassword").value;
  const confirmation = $("#confirmAdminPassword").value;
  if (password !== confirmation) {
    showNotice("Password confirmation does not match.", "warning");
    return;
  }
  const result = await window.erp.setAdminPassword({
    targetUsername: $("#passwordAdminAccount").value,
    password,
  });
  showNotice(result.message, result.ok ? "success" : "warning");
  if (!result.ok) return;
  state.authenticated = true;
  $("#newAdminPassword").value = "";
  $("#confirmAdminPassword").value = "";
  await refreshAccess();
});
$("#changeOwnPassword")?.addEventListener("click", async () => {
  const currentPassword = $("#currentOwnPassword").value;
  const newPassword = $("#newOwnPassword").value;
  const confirmation = $("#confirmOwnPassword").value;
  if (newPassword !== confirmation) {
    showNotice("New password confirmation does not match.", "warning");
    return;
  }
  const result = await window.erp.changeOwnPassword({ currentPassword, newPassword });
  showNotice(result.message, result.ok ? "success" : "warning");
  if (!result.ok) return;
  ["currentOwnPassword", "newOwnPassword", "confirmOwnPassword"].forEach((id) => {
    $(`#${id}`).value = "";
  });
});
async function selectProfilePhoto() {
  const result = await window.erp.chooseProfileImage();
  if (result.canceled) return;
  showNotice(result.message, result.ok ? "success" : "warning");
  if (result.ok) {
    state.profileImageDataUrl = result.imageDataUrl;
    state.profileImagesByUser[state.access.username] = result.imageDataUrl;
    renderProfileImage();
  }
}
$("#changeProfilePhoto")?.addEventListener("click", (e) => {
  e.stopPropagation();
  if (!state.authenticated) {
    openAdminLogin();
  } else {
    selectProfilePhoto();
  }
});
$("#changeProfilePhotoButton")?.addEventListener("click", (e) => {
  e.stopPropagation();
  if (!state.authenticated) {
    openAdminLogin();
  } else {
    selectProfilePhoto();
  }
});
async function performLogoutAndSwitchUser() {
  // Logout must never transition into the sign-in flow. The user can choose
  // Sign In / Create Account afterward from Guest mode.
  state.logoutCompletedAt = Date.now();
  $("#adminLoginOverlay")?.classList.add("hidden");
  state.authenticated = false;
  state.access = {
    username: "",
    role: "guest",
    fullName: "Guest User",
    employeeId: "",
    designation: "",
    departments: [],
    permissions: ["view"],
    filePermissions: ["materials", "codeRequirement", "finishType", "routingUploadFormat", "zkngMaterialUpload", "lsmwBomUpload", "characteristicsUpload", "summary", "diaWiseSummary", "productionCalculation", "department"],
  };
  state.profileImageDataUrl = "";
  state.logoutCompletedAt = Date.now();
  renderAccess();
  changeView("dashboard");
  $("#adminLoginOverlay")?.classList.add("hidden");
  showNotice("Signed out successfully. You are now using Guest mode.", "success");
  if (window.erp && typeof window.erp.logout === "function") {
    window.erp.logout().catch((err) => console.warn("Background logout error:", err));
  }
}

$("#logoutUser")?.addEventListener("click", async (event) => {
  event?.stopPropagation();
  event?.stopImmediatePropagation();
  event?.preventDefault();
  if (state.authenticated) {
    const logoutButton = $("#logoutUser");
    logoutButton.disabled = true;
    logoutButton.textContent = "Logging out...";
    try {
      await performLogoutAndSwitchUser();
    } finally {
      logoutButton.disabled = false;
    }
  } else {
    openAdminLogin();
  }
});

$(".profile")?.addEventListener("click", async (event) => {
  if (event.target.closest("#changeProfilePhoto, #changeProfilePhotoButton, #logoutUser")) return;
  if (!state.authenticated) {
    openAdminLogin();
  }
});
$("#downloadCodeRequirement")?.addEventListener("click", () => downloadResource("codeRequirement"));
$("#uploadRoutingMaster")?.addEventListener("click", async () => {
  const button = $("#uploadRoutingMaster");
  button.disabled = true;
  button.textContent = "Uploading...";
  try {
    const result = await window.erp.uploadRoutingMaster();
    if (result.canceled) return;
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.state) {
      state.routing = result.state;
      renderRoutingState();
    }
  } finally {
    button.disabled = false;
    button.textContent = "Replace Routing Master";
  }
});
$("#uploadZkngMaster")?.addEventListener("click", async () => {
  const button = $("#uploadZkngMaster");
  button.disabled = true;
  button.textContent = "Uploading...";
  try {
    const result = await window.erp.uploadZkngMaster();
    if (result.canceled) return;
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.state) { state.zkng = result.state; renderZkngState(); }
  } catch (error) {
    showNotice(`ZKNG Master upload failed: ${error.message}`, "warning");
  } finally {
    button.disabled = false;
    button.textContent = "Replace ZKNG Master";
  }
});
function showProcessingOverlay(title = "Processing File…", text = "File is processing, please wait…") {
  const overlay = $("#fileProcessingOverlay");
  if (!overlay) return;
  if ($("#processingOverlayTitle")) $("#processingOverlayTitle").textContent = title;
  if ($("#processingOverlayText")) $("#processingOverlayText").textContent = text;
  overlay.classList.remove("hidden");
}

function hideProcessingOverlay() {
  const overlay = $("#fileProcessingOverlay");
  if (!overlay) return;
  overlay.classList.add("hidden");
}

$("#importRoutingMaterial")?.addEventListener("click", async () => {
  const button = $("#importRoutingMaterial");
  button.disabled = true;
  button.innerHTML = '<span class="spinner-icon"></span> Processing...';
  showProcessingOverlay("Processing Routing Material File", "File is processing, please wait…");
  try {
    const result = await window.erp.importRoutingMaterial();
    hideProcessingOverlay();
    if (result.canceled) return;
    showNotice(result.message, result.ok && !result.state?.unmatchedRows ? "success" : "warning");
    if (result.state) {
      state.routing = result.state;
      renderRoutingState();
    }
  } catch (error) {
    hideProcessingOverlay();
    showNotice(`Material import failed: ${error.message}`, "warning");
  } finally {
    hideProcessingOverlay();
    button.innerHTML = "Select & Process File";
    button.disabled = false;
  }
});
$("#importZkngMaterial")?.addEventListener("click", async () => {
  const button = $("#importZkngMaterial");
  button.disabled = true;
  button.innerHTML = '<span class="spinner-icon"></span> Processing...';
  showProcessingOverlay("Processing ZKNG Material File", "File is processing, please wait…");
  try {
    const result = await window.erp.importZkngMaterial();
    hideProcessingOverlay();
    if (result.canceled) return;
    showNotice(result.message, result.ok && !result.state?.unmatchedRows ? "success" : "warning");
    if (result.state) { state.zkng = result.state; renderZkngState(); }
  } catch (error) {
    hideProcessingOverlay();
    showNotice(`ZKNG Material import failed: ${error.message}`, "warning");
  } finally {
    hideProcessingOverlay();
    button.innerHTML = "Select & Process File";
    button.disabled = false;
  }
});
$("#downloadRoutingReady")?.addEventListener("click", async () => {
  const button = $("#downloadRoutingReady");
  button.disabled = true;
  try {
    const result = await window.erp.downloadRoutingReady();
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok) {
      state.lastDownloadedFile = result.path;
      window.erp.recordUsage({ module: "Routing Ready File Download", seconds: 0, file: result.path.split(/[\\/]/).pop() }).catch(() => {});
      $("#downloadedFileName").textContent = result.path.split(/[\\/]/).pop();
      $("#downloadComplete").classList.remove("hidden");
    }
  } finally {
    button.disabled = !state.routing?.ready;
  }
});
async function downloadZkngFiles(includeRouting) {
  const button = includeRouting ? $("#downloadBothReady") : $("#downloadZkngReady");
  button.disabled = true;
  button.classList.add("processing");
  try {
    const result = await window.erp.downloadZkngReady(includeRouting);
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok) {
      state.lastDownloadedFile = result.path;
      $("#downloadedFileName").textContent = (result.paths || [result.path]).map((file) => file.split(/[\\/]/).pop()).join(" + ");
      $("#downloadComplete").classList.remove("hidden");
      window.erp.recordUsage({ module: includeRouting ? "Material + Routing Ready Download" : "ZKNG Material Ready Download", seconds: 0, file: $("#downloadedFileName").textContent }).catch(() => {});
    }
  } catch (error) {
    showNotice(`Ready file download failed: ${error.message}`, "warning");
  } finally {
    button.classList.remove("processing");
    button.disabled = includeRouting ? !(state.zkng?.materialReady && state.zkng?.ready) : !state.zkng?.materialReady;
  }
}
$("#downloadZkngReady")?.addEventListener("click", () => downloadZkngFiles(false));
$("#downloadBothReady")?.addEventListener("click", () => downloadZkngFiles(true));
$("#clearRoutingMaterial")?.addEventListener("click", async () => {
  if (!window.confirm("Reset the previous Material Code input and Routing Ready file?")) return;
  const button = $("#clearRoutingMaterial");
  button.disabled = true;
  button.textContent = "Removing...";
  try {
    const result = await window.erp.clearRoutingSession("routing");
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok && result.state) {
      state.routing = result.state;
      renderRoutingState();
    }
  } catch (error) {
    showNotice(`Previous Routing file could not be removed: ${error.message}`, "warning");
  } finally {
    button.textContent = "Reset Input";
    button.disabled = !state.routing?.canImport || !state.routing?.materialFile;
  }
});
$("#clearZkngMaterial")?.addEventListener("click", async () => {
  const button = $("#clearZkngMaterial");
  button.disabled = true;
  button.textContent = "Resetting...";
  try {
    const result = await window.erp.clearRoutingSession("zkng");
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.state) { state.zkng = result.state; renderZkngState(); }
  } catch (error) {
    showNotice(`ZKNG input could not be reset: ${error.message}`, "warning");
  } finally {
    button.textContent = "Reset Input";
    button.disabled = !state.zkng?.canImport || !state.zkng?.materialFile;
  }
});
$("#uploadLsmwBomMaster")?.addEventListener("click", async () => {
  if (typeof window.erp.uploadLsmwBomMaster === "function") {
    const result = await window.erp.uploadLsmwBomMaster();
    if (result.state) { state.lsmwBom = result.state; renderLsmwBomState(); }
    showNotice(result.message, result.ok ? "success" : "warning");
  }
});
$("#importLsmwBomMaterial")?.addEventListener("click", async () => {
  const button = $("#importLsmwBomMaterial");
  button.disabled = true;
  button.innerHTML = '<span class="spinner-icon"></span> Processing...';
  showProcessingOverlay("Processing LSMW BOM Material File", "File is processing, please wait…");
  try {
    if (typeof window.erp.importLsmwBomMaterial === "function") {
      const result = await window.erp.importLsmwBomMaterial();
      hideProcessingOverlay();
      if (result.canceled) return;
      showNotice(result.message, result.ok ? "success" : "warning");
      if (result.state) { state.lsmwBom = result.state; renderLsmwBomState(); }
    } else {
      hideProcessingOverlay();
      showNotice("BOM file processing ready.", "info");
    }
  } catch (error) {
    hideProcessingOverlay();
    showNotice(`BOM import failed: ${error.message}`, "warning");
  } finally {
    hideProcessingOverlay();
    button.innerHTML = "Select & Process File";
    button.disabled = false;
  }
});
async function handleLsmwBomDownload(fileType) {
  if (typeof window.erp.downloadLsmwBomReady === "function") {
    const result = await window.erp.downloadLsmwBomReady(fileType);
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok) {
      state.lastDownloadedFile = result.path;
      $("#downloadedFileName").textContent = (result.paths || [result.path]).map((f) => f.split(/[\\/]/).pop()).join(" + ");
      $("#downloadComplete").classList.remove("hidden");
    }
  }
}
$("#downloadAllLsmwBom")?.addEventListener("click", () => handleLsmwBomDownload("all"));
$("#downloadHeaderBomXlsx")?.addEventListener("click", () => handleLsmwBomDownload("Header_BOM.xlsx"));
$("#downloadHeaderBomTxt")?.addEventListener("click", () => handleLsmwBomDownload("Header_BOM.txt"));
$("#downloadItemBomXlsx")?.addEventListener("click", () => handleLsmwBomDownload("Item_BOM.xlsx"));
$("#downloadItemBomTxt")?.addEventListener("click", () => handleLsmwBomDownload("Item_BOM.txt"));
$("#clearLsmwBomMaterial")?.addEventListener("click", async () => {
  if (typeof window.erp.clearLsmwBomSession === "function") {
    const result = await window.erp.clearLsmwBomSession();
    if (result.state) { state.lsmwBom = result.state; renderLsmwBomState(); }
    showNotice(result.message, result.ok ? "success" : "warning");
  }
});
$("#uploadCharacteristicsMaster")?.addEventListener("click", async () => {
  if (typeof window.erp.uploadCharacteristicsMaster === "function") {
    const result = await window.erp.uploadCharacteristicsMaster();
    if (result.state) { state.characteristics = result.state; renderCharacteristicsState(); }
    showNotice(result.message, result.ok ? "success" : "warning");
  }
});
$("#importCharacteristicsMaterial")?.addEventListener("click", async () => {
  const button = $("#importCharacteristicsMaterial");
  button.disabled = true;
  button.innerHTML = '<span class="spinner-icon"></span> Processing...';
  showProcessingOverlay("Processing Characteristics Material File", "File is processing, please wait…");
  try {
    if (typeof window.erp.importCharacteristicsMaterial === "function") {
      const result = await window.erp.importCharacteristicsMaterial();
      hideProcessingOverlay();
      if (result.canceled) return;
      showNotice(result.message, result.ok ? "success" : "warning");
      if (result.state) { state.characteristics = result.state; renderCharacteristicsState(); }
    }
  } catch (error) {
    hideProcessingOverlay();
    showNotice(`Characteristics import failed: ${error.message}`, "warning");
  } finally {
    hideProcessingOverlay();
    button.innerHTML = "Select & Process File";
    button.disabled = false;
  }
});
$("#downloadCharacteristicsReady")?.addEventListener("click", async () => {
  if (typeof window.erp.downloadCharacteristicsReady === "function") {
    const result = await window.erp.downloadCharacteristicsReady();
    showNotice(result.message, result.ok ? "success" : "warning");
    if (result.ok) {
      state.lastDownloadedFile = result.path;
      window.erp.recordUsage({ module: "Characteristics Ready File Download", seconds: 0, file: result.path.split(/[\\/]/).pop() }).catch(() => {});
      $("#downloadedFileName").textContent = result.path.split(/[\\/]/).pop();
      $("#downloadComplete").classList.remove("hidden");
    }
  }
});
$("#clearCharacteristicsMaterial")?.addEventListener("click", async () => {
  if (typeof window.erp.clearCharacteristicsSession === "function") {
    const result = await window.erp.clearCharacteristicsSession();
    if (result.state) { state.characteristics = result.state; renderCharacteristicsState(); }
    showNotice(result.message, result.ok ? "success" : "warning");
  }
});
$$(".export").forEach((button) => button.addEventListener("click", () => exportCurrent(button.dataset.format)));

initialize().catch((error) => showNotice(`Software could not start: ${error.message}`, "warning"));
window.setInterval(async () => {
  const seconds = Math.max(1, Math.round((Date.now() - (state.viewStartedAt || Date.now())) / 1000));
  const currentLabel = { dashboard: "Overview", materials: "Fabric Material Code", codeRequirement: "Code Requirement File", routingUploadFormat: "Routing Upload File", characteristicsUpload: "Characteristics Upload File", zkngMaterialUpload: "ZKNG Material Upload File", lsmwBomUpload: "LSMW BOM Upload File", finishType: "Finish Type Short Name", summary: "Circular Knitting", diaWiseSummary: "Dia Wise Summary", productionCalculation: "Production Calculation", department: state.department || "Department Workspace", administration: "Administration", settings: "Settings" }[state.view] || state.view;
  window.erp.recordUsage({ module: currentLabel, seconds, count: false, file: state.departmentStatuses?.[state.department]?.file || "" }).catch(() => {});
  state.viewStartedAt = Date.now();
  const result = await window.erp.accountHeartbeat();
  if (result.ok && state.access.role === "main_admin" && state.view === "administration") {
    await refreshAccess();
  }
}, 60 * 1000);
