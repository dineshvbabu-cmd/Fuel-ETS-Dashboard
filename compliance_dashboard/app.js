import {
  STORAGE_KEY,
  SHEETS,
  SHEET_COLUMNS,
  blankCalculatorRow,
  blankRowForSheet,
  createStateFromSeed,
  deepClone,
  normalizeText,
  numberOrZero,
  persistableState,
  recalculateWorkbook,
} from "./engine.js";

const LIBRARY_PAGE_SIZE = 100;
const CALCULATOR_HISTORY_PAGE_SIZE = 100;
const MAX_CALCULATOR_ROWS_PER_YEAR = 1000;
const REMOTE_SAVE_DELAY_MS = 700;
const APP_BUILD = "2026.06.11.4";
const REFERENCE_SHEETS = SHEETS.filter((sheet) => !["dashboard", "calculator", "vesselSummary"].includes(sheet.key));
const PROJECTION_BASELINE = {
  waspFactor: 0.95,
  bioBlend: 0,
  bioType: "Bio-diesel",
  rfnboBlend: 0,
  rfnboType: "e-diesel",
};
const EDITOR_SELECT_OPTIONS = {
  fuelReference: {
    fuelClass: [
      { value: "None", label: "None" },
      { value: "Fossil", label: "Fossil Fuel" },
      { value: "Biofuel", label: "Biofuel" },
      { value: "RFNBO", label: "RFNBO" },
      { value: "Electricity", label: "Electricity" },
      { value: "Other", label: "Other" },
    ],
  },
};

const elements = {
  viewTabs: document.getElementById("viewTabs"),
  vesselFilter: document.getElementById("vesselFilter"),
  syncStatus: document.getElementById("syncStatus"),
  generateReportButton: document.getElementById("generateReportButton"),
  libraryToggleButton: document.getElementById("libraryToggleButton"),
  exportFilteredButton: document.getElementById("exportFilteredButton"),
  resetWorkbookButton: document.getElementById("resetWorkbookButton"),
  kpiGrid: document.getElementById("kpiGrid"),
  contentView: document.getElementById("contentView"),
  libraryDrawer: document.getElementById("libraryDrawer"),
  libraryTabs: document.getElementById("libraryTabs"),
  libraryContent: document.getElementById("libraryContent"),
  libraryBackdrop: document.getElementById("libraryBackdrop"),
  closeLibraryButton: document.getElementById("closeLibraryButton"),
  rowEditorDialog: document.getElementById("rowEditorDialog"),
  editorDialogTitle: document.getElementById("editorDialogTitle"),
  editorDialogBody: document.getElementById("editorDialogBody"),
  closeEditorButton: document.getElementById("closeEditorButton"),
  saveEditorButton: document.getElementById("saveEditorButton"),
  portCodes: document.getElementById("portCodes"),
  fuelTypes: document.getElementById("fuelTypes"),
  imoNumbers: document.getElementById("imoNumbers"),
  reportDialog: document.getElementById("reportDialog"),
  closeReportButton: document.getElementById("closeReportButton"),
  reportVesselSelect: document.getElementById("reportVesselSelect"),
  reportSelectionSummary: document.getElementById("reportSelectionSummary"),
  reportRowList: document.getElementById("reportRowList"),
  selectAllReportRowsButton: document.getElementById("selectAllReportRowsButton"),
  clearReportRowsButton: document.getElementById("clearReportRowsButton"),
  generateReportConfirmButton: document.getElementById("generateReportConfirmButton"),
};

const stateStore = {
  seedState: null,
  state: null,
  derived: null,
  charts: {},
  market: {
    status: "idle",
    snapshot: null,
    error: null,
  },
  sync: {
    ready: false,
    status: "loading",
    storage: null,
    revision: "",
    updatedAt: "",
    timer: null,
    inFlight: false,
    queued: false,
    error: null,
  },
  ui: {
    activeView: "dashboard",
    vesselFilter: "all",
    calculatorSearch: "",
    calculatorColumnMenuOpen: false,
    calculatorVisibleColumns: [],
    calculatorActiveScrollLeft: 0,
    calculatorHistoryScrollLeft: 0,
    calculatorHistoryScrollTop: 0,
    detailSearch: "",
    detailScope: "all",
    detailScrollLeft: 0,
    calculatorSelectedId: null,
    calculatorHistoryOpen: false,
    calculatorHistoryPage: 1,
    libraryOpen: false,
    librarySheet: "parameters",
    librarySearch: "",
    libraryPage: 1,
    dialog: null,
    drilldown: null,
    kpisOpen: true,
    projectionOpen: true,
    chartsOpen: true,
    voyageTableOpen: true,
    calculatorOpen: true,
    projectionPreset: "baseline",
    projection: { ...PROJECTION_BASELINE },
  },
};

const numericColumns = {
  fuelReference: new Set(["lcvMjPerG", "wtWPerMj", "rwd", "etsCo2Cf", "cfCo2PerG", "cfCh4PerG", "cfN2oPerG", "cslipPercent"]),
  fleet: new Set(["imoNo", "gt", "nt", "summerDwt", "built", "wapsFwindFactor"]),
  derogations: new Set(["serialNo"]),
};

const CALCULATOR_COLUMNS = [
  { key: "recordId", label: "Voyage / Port-Stay ID", kind: "sticky-editable", input: "text", width: 130, placeholder: "V001 / P001" },
  { key: "type", label: "Type", kind: "sticky-editable", input: "select", width: 120, options: ["Voyage", "Port Stay"] },
  { key: "imoNo", label: "IMO No.", kind: "sticky-editable", input: "text", width: 110, list: "imoNumbers" },
  { key: "vesselName", label: "Vessel Name", kind: "sticky", width: 150 },
  { key: "shipType", label: "Ship Type", kind: "calculated", width: 150 },
  { key: "flagState", label: "Flag State", kind: "calculated", width: 120 },
  { key: "deadweightTonnes", label: "Deadweight (DWT, t)", kind: "calculated-number", width: 125, digits: 0 },
  { key: "netTonnage", label: "Net Tonnage (NT)", kind: "calculated-number", width: 120, digits: 0 },
  { key: "grossTonnage", label: "Gross Tonnage (GT)", kind: "calculated-number", width: 120, digits: 0 },
  { key: "departureDate", label: "Departure Date", kind: "editable", input: "date", width: 120 },
  { key: "fromPortCode", label: "From Port UN/LOCODE", kind: "editable", input: "text", width: 130, list: "portCodes" },
  { key: "fromPortName", label: "From Port Name", kind: "calculated", width: 150 },
  { key: "arrivalDate", label: "Arrival Date", kind: "editable", input: "date", width: 120 },
  { key: "toPortCode", label: "To Port UN/LOCODE", kind: "editable", input: "text", width: 130, list: "portCodes" },
  { key: "toPortName", label: "To Port Name", kind: "calculated", width: 150 },
  { key: "fuel1Type", label: "Fossil Fuel 1 Type", kind: "editable", input: "text", width: 130, list: "fuelTypes" },
  { key: "fuel1ConsumptionMt", label: "Fossil Fuel 1 Cons. (MT)", kind: "editable-number", input: "number", width: 130, digits: 2, step: "0.0001" },
  { key: "fuel2Type", label: "Fossil Fuel 2 Type", kind: "editable", input: "text", width: 130, list: "fuelTypes" },
  { key: "fuel2ConsumptionMt", label: "Fossil Fuel 2 Cons. (MT)", kind: "editable-number", input: "number", width: 130, digits: 2, step: "0.0001" },
  { key: "bioFuelType", label: "Biofuel / RFNBO Type", kind: "editable", input: "text", width: 140, list: "fuelTypes" },
  { key: "bioFuelConsumptionMt", label: "Biofuel / RFNBO Cons. (MT)", kind: "editable-number", input: "number", width: 140, digits: 2, step: "0.0001" },
  { key: "sustainabilityFactor", label: "Sustain. Factor WtW", kind: "editable-number", input: "number", width: 120, digits: 2, step: "0.01" },
  { key: "windFactor", label: "WASP Factor", kind: "editable-number", input: "number", width: 110, digits: 2, step: "0.01" },
  { key: "distanceNm", label: "Total Distance (nm)", kind: "editable-number", input: "number", width: 120, digits: 0, step: "0.01" },
  { key: "cargoTonnes", label: "Cargo Carried (t)", kind: "editable-number", input: "number", width: 120, digits: 0, step: "0.01" },
  { key: "timeAtSeaHours", label: "Time at Sea (h)", kind: "editable-number", input: "number", width: 120, digits: 1, step: "0.01" },
  { key: "berthHours", label: "Hours at Berth / Anchor", kind: "editable-number", input: "number", width: 130, digits: 1, step: "0.01" },
  { key: "opsElectricityMj", label: "OPS Electricity (MJ)", kind: "editable-number", input: "number", width: 130, digits: 0, step: "0.01" },
  { key: "fromEuEea", label: "From EU/EEA?", kind: "calculated", width: 100 },
  { key: "toEuEea", label: "To EU/EEA?", kind: "calculated", width: 100 },
  { key: "omrInvolved", label: "OMR Involved?", kind: "calculated", width: 110 },
  { key: "scopePercent", label: "Scope / Leg %", kind: "calculated-percent", width: 100 },
  { key: "scopeNote", label: "Derogation / Scope Note", kind: "calculated", width: 220 },
  { key: "totalEnergyMj", label: "Total Energy (MJ)", kind: "calculated-number", width: 130, digits: 0 },
  { key: "inScopeEnergyMj", label: "In-Scope Energy (MJ)", kind: "calculated-number", width: 130, digits: 0 },
  { key: "transportWork", label: "Transport Work", kind: "calculated-number", width: 130, digits: 0 },
  { key: "etsInScopeCo2eqT", label: "ETS In-Scope CO2eq (t)", kind: "calculated-number", width: 135, digits: 3 },
  { key: "euasRequiredT", label: "EUAs Required (t)", kind: "calculated-number", width: 120, digits: 3 },
  { key: "euasCostEur", label: "EUAs Cost (EUR)", kind: "calculated-currency", width: 125 },
  { key: "attainedGhgIntensity", label: "Attained GHG Intensity", kind: "calculated-number", width: 135, digits: 3 },
  { key: "targetGhgIntensity", label: "Target GHG Intensity", kind: "calculated-number", width: 130, digits: 3 },
  { key: "complianceBalanceT", label: "Compliance Balance (t)", kind: "calculated-number", width: 140, digits: 3 },
  { key: "fuelEuPenaltyEur", label: "FuelEU Penalty (EUR)", kind: "calculated-currency", width: 135 },
  { key: "rowActions", label: "Actions", kind: "actions", width: 100 },
];

const DETAIL_TABLE_COLUMNS = [
  { key: "recordId", label: "ID", width: 84 },
  { key: "type", label: "Type", width: 88, format: "type" },
  { key: "vesselName", label: "Vessel", width: 108 },
  { key: "fromPortName", label: "From", width: 138 },
  { key: "toPortName", label: "To", width: 138 },
  { key: "departureDate", label: "Departure", width: 104, format: "date" },
  { key: "arrivalDate", label: "Arrival", width: 104, format: "date" },
  { key: "fuel1Type", label: "Fuel 1", width: 92 },
  { key: "fuel1ConsumptionMt", label: "F1 MT", width: 82, format: "number", digits: 2 },
  { key: "fuel2Type", label: "Fuel 2", width: 92 },
  { key: "fuel2ConsumptionMt", label: "F2 MT", width: 82, format: "number", digits: 2 },
  { key: "bioFuelType", label: "Biofuel Type", width: 118 },
  { key: "bioFuelConsumptionMt", label: "Biofuel Qty", width: 98, format: "number", digits: 2 },
  { key: "opsElectricityMj", label: "OPS Electricity", width: 112, format: "number", digits: 0 },
  { key: "scopePercent", label: "Scope", width: 76, format: "scope" },
  { key: "euasRequiredT", label: "EUAs (t)", width: 90, format: "number", digits: 2 },
  { key: "euasCostEur", label: "EUA Cost", width: 106, format: "currency" },
  { key: "attainedGhgIntensity", label: "GHG", width: 84, format: "ghg", digits: 2 },
  { key: "complianceBalanceT", label: "Balance t", width: 96, format: "balance", digits: 2 },
  { key: "fuelEuPenaltyEur", label: "Penalty", width: 140, format: "currency" },
];

const LIBRARY_DISPLAY_COLUMNS = {
  fuelReference: [
    "fuelPathway",
    "fuelClass",
    "lcvMjPerG",
    "wtWPerMj",
    "ttwCo2eqPerG",
    "wtwIntensity",
    "rwd",
    "etsCo2Cf",
    "etsTtwAr5",
    "etsNonCo2Ar5",
    "cfCo2PerG",
    "cfCh4PerG",
    "cfN2oPerG",
    "cslipPercent",
    "consumerSource",
    "notes",
    "alias",
  ],
};

const LIBRARY_COLUMN_LABELS = {
  fuelReference: {
    fuelPathway: "Fuel Pathway",
    fuelClass: "Class",
    lcvMjPerG: "LCV (MJ/g)",
    wtWPerMj: "WtT (gCO2eq/MJ)",
    ttwCo2eqPerG: "TtW CO2eq (gCO2eq/gFuel)",
    wtwIntensity: "WtW intensity (gCO2eq/MJ)",
    rwd: "RWD (RFNBO reward)",
    etsCo2Cf: "ETS CO2 Cf (tCO2/t)",
    etsTtwAr5: "ETS TtW CO2eq AR5 (gCO2eq/gFuel)",
    alias: "Column 11",
    etsNonCo2Ar5: "ETS non-CO2 TtW AR5 (gCO2eq/gFuel)",
    cfCo2PerG: "Cf CO2 (g/gFuel)",
    cfCh4PerG: "Cf CH4 (g/gFuel)",
    cfN2oPerG: "Cf N2O (g/gFuel)",
    cslipPercent: "Cslip (% mass)",
    consumerSource: "Fuel consumer / source (Annex II)",
    notes: "Notes",
  },
  fleet: {
    imoNo: "IMO No.",
    vesselName: "Vessel Name",
    shipType: "Ship Type",
    flag: "Flag",
    className: "Class",
    gt: "GT (GRT)",
    nt: "NT (NRT)",
    summerDwt: "Summer DWT",
    built: "Built",
    wapsFwindFactor: "WAPS Fwind factor",
  },
};

const CALCULATOR_HISTORY_COLUMNS = [
  { key: "recordId", label: "ID" },
  { key: "type", label: "Type" },
  { key: "vesselName", label: "Vessel" },
  { key: "route", label: "Route" },
  { key: "departureDate", label: "Departure" },
  { key: "arrivalDate", label: "Arrival" },
  { key: "euasRequiredT", label: "EUAs" },
  { key: "attainedGhgIntensity", label: "GHG" },
];

const CALCULATOR_MANUAL_VALUE_KEYS = [
  "imoNo",
  "departureDate",
  "fromPortCode",
  "arrivalDate",
  "toPortCode",
  "fuel1Type",
  "fuel1ConsumptionMt",
  "fuel2Type",
  "fuel2ConsumptionMt",
  "bioFuelType",
  "bioFuelConsumptionMt",
  "sustainabilityFactor",
  "windFactor",
  "distanceNm",
  "cargoTonnes",
  "timeAtSeaHours",
  "berthHours",
  "opsElectricityMj",
];

const DEFAULT_CALCULATOR_VISIBLE_COLUMNS = CALCULATOR_COLUMNS.filter((column) => column.key !== "rowActions").map((column) => column.key);

function formatNumber(value, digits = 2) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: digits,
    minimumFractionDigits: 0,
  }).format(Number(value));
}

function formatInteger(value) {
  return formatNumber(value, 0);
}

function formatCurrency(value) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 0,
  }).format(Number(value));
}

function formatUsdCurrency(value) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(Number(value));
}

function formatPercent(value, digits = 0) {
  return `${formatNumber((Number(value) || 0) * 100, digits)}%`;
}

function formatDateValue(value) {
  if (!value) {
    return "-";
  }
  const text = String(value);
  return text.includes("T") ? text.slice(0, 10) : text;
}

function lower(value) {
  return normalizeText(value).toLowerCase();
}

function getSyncClientId() {
  const key = `${STORAGE_KEY}-client-id`;
  let clientId = localStorage.getItem(key);
  if (!clientId) {
    clientId = globalThis.crypto?.randomUUID?.() || `browser-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    localStorage.setItem(key, clientId);
  }
  return clientId;
}

function renderSyncStatus() {
  if (!elements.syncStatus) return;
  const { status, storage, updatedAt, error } = stateStore.sync;
  const durable = Boolean(storage?.durable);
  const labels = {
    loading: "Checking storage",
    pending: "Pending save",
    saving: "Saving...",
    saved: durable ? "Saved to R2" : "Server fallback",
    error: "Local only",
  };
  elements.syncStatus.textContent = labels[status] || "Local";
  elements.syncStatus.className = `sync-status sync-${status} ${durable ? "sync-durable" : ""}`;
  elements.syncStatus.title =
    status === "error"
      ? `Browser copy is safe. Server sync failed: ${error || "Unknown error"}`
      : durable
        ? `Durable R2 storage${updatedAt ? ` - last saved ${new Date(updatedAt).toLocaleString()}` : ""}`
        : "Browser autosave is active. Configure R2 variables in Railway for durable cross-device storage.";
}

function scheduleRemoteSave(delay = REMOTE_SAVE_DELAY_MS) {
  if (!stateStore.sync.ready) return;
  window.clearTimeout(stateStore.sync.timer);
  stateStore.sync.status = "pending";
  renderSyncStatus();
  stateStore.sync.timer = window.setTimeout(syncStateToServer, delay);
}

async function syncStateToServer() {
  if (!stateStore.sync.ready) return;
  if (stateStore.sync.inFlight) {
    stateStore.sync.queued = true;
    return;
  }

  stateStore.sync.inFlight = true;
  stateStore.sync.queued = false;
  stateStore.sync.status = "saving";
  renderSyncStatus();
  try {
    const response = await fetch("/api/state", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        clientId: getSyncClientId(),
        state: persistableState(stateStore.state),
      }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.message || `Storage responded with ${response.status}`);
    }
    stateStore.sync.storage = payload.storage;
    stateStore.sync.revision = payload.revision || "";
    stateStore.sync.updatedAt = payload.updatedAt || "";
    stateStore.sync.status = "saved";
    stateStore.sync.error = null;
  } catch (error) {
    stateStore.sync.status = "error";
    stateStore.sync.error = error.message;
  } finally {
    stateStore.sync.inFlight = false;
    renderSyncStatus();
    if (stateStore.sync.queued) {
      scheduleRemoteSave(100);
    }
  }
}

async function hydrateFromServer(localState) {
  try {
    const response = await fetch("/api/state", { headers: { Accept: "application/json" } });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.message || `Storage responded with ${response.status}`);
    }
    stateStore.sync.storage = payload.storage;
    stateStore.sync.ready = true;
    if (payload.exists && payload.document?.state && Array.isArray(payload.document.state.calculatorRows)) {
      stateStore.sync.revision = payload.document.revision || "";
      stateStore.sync.updatedAt = payload.document.updatedAt || "";
      stateStore.sync.status = "saved";
      return payload.document.state;
    }
    stateStore.sync.status = "pending";
    window.setTimeout(() => scheduleRemoteSave(100), 0);
    return localState;
  } catch (error) {
    stateStore.sync.ready = true;
    stateStore.sync.status = "error";
    stateStore.sync.error = error.message;
    return localState;
  } finally {
    renderSyncStatus();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(persistableState(stateStore.state)));
  scheduleRemoteSave();
}

function getCurrentReportYear() {
  return Number(stateStore.derived?.parameterValues?.reportYear) || new Date().getFullYear();
}

function extractYearFromDate(value) {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})-/);
  return match ? Number(match[1]) : null;
}

function resolveRowStorageYear(row) {
  return Number(row?.storageYear) || extractYearFromDate(row?.departureDate) || extractYearFromDate(row?.arrivalDate) || getCurrentReportYear();
}

function getCalculatorStateRow(rowId) {
  return stateStore.state.calculatorRows.find((item) => item.id === rowId) || null;
}

function rowHasMeaningfulInputs(stateRow) {
  if (!stateRow) return false;
  return CALCULATOR_MANUAL_VALUE_KEYS.some((key) => {
    const value = stateRow[key];
    if (value === null || value === undefined || value === "") {
      return false;
    }
    if (typeof value === "string" && (value === "(none)" || !value.trim())) {
      return false;
    }
    return true;
  });
}

function getMeaningfulDerivedRows() {
  return stateStore.derived.calculatorRows.filter((row) => rowHasMeaningfulInputs(getCalculatorStateRow(row.id)));
}

function countRowsForYear(year) {
  return stateStore.state.calculatorRows.filter((row) => rowHasMeaningfulInputs(row) && resolveRowStorageYear(row) === year).length;
}

function compactCalculatorRowsForRuntime() {
  const rows = Array.isArray(stateStore.state.calculatorRows) ? stateStore.state.calculatorRows : [];
  const meaningfulRows = rows.filter((row) => rowHasMeaningfulInputs(row));
  const existingDraft = rows.find((row) => !rowHasMeaningfulInputs(row));
  const draft = existingDraft ? { ...existingDraft } : blankCalculatorRow();
  draft.recordId = "";
  draft.type = draft.type === "Port Stay" ? "Port Stay" : "Voyage";
  stateStore.state.calculatorRows = [...meaningfulRows, draft];
}

function pruneCalculatorDraftRows(keepIds = []) {
  const keep = new Set(keepIds.filter(Boolean));
  stateStore.state.calculatorRows = stateStore.state.calculatorRows.filter((row) => rowHasMeaningfulInputs(row) || keep.has(row.id));
}

function getVisibleCalculatorColumnKeys() {
  const selected = Array.isArray(stateStore.ui.calculatorVisibleColumns) && stateStore.ui.calculatorVisibleColumns.length
    ? stateStore.ui.calculatorVisibleColumns
    : DEFAULT_CALCULATOR_VISIBLE_COLUMNS;
  const ordered = CALCULATOR_COLUMNS.map((column) => column.key).filter((key) => key === "rowActions" || selected.includes(key));
  return ordered.includes("rowActions") ? ordered : [...ordered, "rowActions"];
}

function getVisibleCalculatorColumns() {
  const keys = getVisibleCalculatorColumnKeys();
  return CALCULATOR_COLUMNS.filter((column) => keys.includes(column.key));
}

function estimateCalculatorColumnWidth(column, rows, inputRowsById) {
  if (column.key === "rowActions") {
    return 180;
  }
  const sampleRows = rows.slice(0, 16);
  const maxValueLength = sampleRows.reduce((max, row) => {
    const inputRow = inputRowsById.get(row.id) || blankCalculatorRow();
    const displayValue = column.kind.includes("editable")
      ? String(calculatorInputValue(inputRow, row, column) ?? "")
      : String(calculatorCellValue(row, column) ?? "");
    return Math.max(max, displayValue.length);
  }, 0);
  const sourceLength = Math.max(column.label.length, maxValueLength, column.placeholder?.length || 0);
  const charWidth = column.kind.includes("editable") ? 8.1 : 7.6;
  const minWidth = Math.min(column.width || 100, 120);
  const maxWidth = column.kind.includes("calculated") ? 260 : 240;
  return Math.max(minWidth, Math.min(maxWidth, Math.round(sourceLength * charWidth + 34)));
}

function syncCalculatorDraftRowsWithDerived() {
  const derivedById = new Map(stateStore.derived.calculatorRows.map((row) => [row.id, row]));
  let changed = false;

  stateStore.state.calculatorRows.forEach((row) => {
    if (!normalizeText(row.type)) {
      row.type = "Voyage";
      changed = true;
    }

    if (!normalizeText(row.recordId)) {
      row.recordId = nextRecordSerial(row.type || "Voyage");
      changed = true;
    }

    const derivedRow = derivedById.get(row.id);
    if (!derivedRow) {
      return;
    }

    if (!normalizeText(row.recordId) && normalizeText(derivedRow.recordId)) {
      row.recordId = derivedRow.recordId;
      changed = true;
    }

    if (!normalizeText(row.type) && normalizeText(derivedRow.type)) {
      row.type = derivedRow.type;
      changed = true;
    }

    if ((row.windFactor === null || row.windFactor === undefined || row.windFactor === "") && derivedRow.windFactor !== null && derivedRow.windFactor !== undefined && derivedRow.windFactor !== "") {
      row.windFactor = Number(derivedRow.windFactor);
      changed = true;
    }
  });

  return changed;
}

function hydrateFromStorage(seedState) {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return deepClone(seedState);
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.calculatorRows) || !Array.isArray(parsed.parameters)) {
      return deepClone(seedState);
    }
    return parsed;
  } catch {
    return deepClone(seedState);
  }
}

function getCollection(sheetKey) {
  return stateStore.state[sheetKey];
}

function recomputeAndRender() {
  stateStore.derived = recalculateWorkbook(stateStore.state);
  if (syncCalculatorDraftRowsWithDerived()) {
    stateStore.derived = recalculateWorkbook(stateStore.state);
  }
  if (!getCalculatorStateRow(stateStore.ui.calculatorSelectedId)) {
    const draftRow = buildCalculatorRowForCurrentFilter();
    stateStore.state.calculatorRows.unshift(draftRow);
    stateStore.ui.calculatorSelectedId = draftRow.id;
    stateStore.derived = recalculateWorkbook(stateStore.state);
  }
  stateStore.state.parameters = deepClone(stateStore.derived.parameters);
  saveState();
  render();
}

function destroyCharts() {
  Object.values(stateStore.charts).forEach((chart) => chart.destroy());
  stateStore.charts = {};
}

function buildDataLists() {
  // Heavy global datalists make spreadsheet-style editing feel sluggish.
  // Active-row inputs now use lightweight per-cell suggestion lists instead.
  elements.portCodes.innerHTML = "";
  elements.fuelTypes.innerHTML = "";
  elements.imoNumbers.innerHTML = "";
}

function renderViewTabs() {
  elements.viewTabs.innerHTML = [
    ["dashboard", "Dashboard", "tab-blue"],
    ["calculator", "Voyage Inputs", "tab-amber"],
  ]
    .map(
      ([key, label, tone]) => `
        <button
          class="view-tab ${tone} ${stateStore.ui.activeView === key ? "active" : ""}"
          type="button"
          data-action="select-view"
          data-view="${key}"
        >
          ${label}
        </button>
      `
    )
    .join("");
}

function getActiveRows() {
  const allRows = getMeaningfulDerivedRows();
  if (stateStore.ui.vesselFilter === "all") {
    return allRows;
  }
  return allRows.filter((row) => row.vesselName === stateStore.ui.vesselFilter);
}

function getVisibleVessels() {
  return [...new Set(getMeaningfulDerivedRows().map((row) => row.vesselName).filter(Boolean))].sort();
}

function renderVesselFilter() {
  const vessels = getVisibleVessels();
  elements.vesselFilter.innerHTML = [
    `<option value="all">All vessels</option>`,
    ...vessels.map((vessel) => `<option value="${vessel}">${vessel}</option>`),
  ].join("");
  elements.vesselFilter.value = vessels.includes(stateStore.ui.vesselFilter) ? stateStore.ui.vesselFilter : "all";
  if (!vessels.includes(stateStore.ui.vesselFilter) && stateStore.ui.vesselFilter !== "all") {
    stateStore.ui.vesselFilter = "all";
  }
}

function computeFilteredDashboard(activeRows) {
  return computeDashboardSummary(activeRows);
}

function computeDashboardSummary(activeRows, metrics = {}) {
  const euasKey = metrics.euasKey || "euasRequiredT";
  const costKey = metrics.costKey || "euasCostEur";
  const penaltyKey = metrics.penaltyKey || "fuelEuPenaltyEur";
  const numeratorKey = metrics.numeratorKey || "fuelEuWtwEmissionsG";
  const denominatorKey = metrics.denominatorKey || "fuelEuDenomStep1Mj";
  const energyKey = metrics.energyKey || "fuelEuEnergyStep2Mj";
  const balanceKey = metrics.balanceKey || "complianceBalanceT";
  const ghgKey = metrics.ghgKey || "attainedGhgIntensity";

  const totalEuasRequired = activeRows.reduce((sum, row) => sum + numberOrZero(row[euasKey]), 0);
  const totalEuasCost = activeRows.reduce((sum, row) => sum + numberOrZero(row[costKey]), 0);
  const totalPenalty = activeRows.reduce((sum, row) => sum + numberOrZero(row[penaltyKey]), 0);
  const totalFuelConsumed = activeRows.reduce(
    (sum, row) =>
      sum +
      numberOrZero(row.fuel1ConsumptionMt) +
      numberOrZero(row.fuel2ConsumptionMt) +
      numberOrZero(row.bioFuelConsumptionMt),
    0
  );
  const totalNumerator = activeRows.reduce((sum, row) => sum + numberOrZero(row[numeratorKey]), 0);
  const totalDenominator = activeRows.reduce((sum, row) => sum + numberOrZero(row[denominatorKey]), 0);
  const totalEnergy = activeRows.reduce((sum, row) => sum + numberOrZero(row[energyKey]), 0);
  const averageIntensity = totalDenominator > 0 ? totalNumerator / totalDenominator : 0;
  const target = stateStore.derived.parameterValues.fueleuTarget;
  const complianceBalance = totalDenominator > 0 ? (target - averageIntensity) * totalEnergy / 1_000_000 : activeRows.reduce((sum, row) => sum + numberOrZero(row[balanceKey]), 0);
  const voyageRows = activeRows.filter((row) => row.type === "Voyage");
  const portStayRows = activeRows.filter((row) => row.type === "Port Stay");

  const byVessel = [...new Set(activeRows.map((row) => row.vesselName).filter(Boolean))]
    .map((vesselName) => {
      const rows = activeRows.filter((row) => row.vesselName === vesselName);
      return {
        vesselName,
        euasRequired: rows.reduce((sum, row) => sum + numberOrZero(row[euasKey]), 0),
        euasCost: rows.reduce((sum, row) => sum + numberOrZero(row[costKey]), 0),
        averageIntensity: (() => {
          const numerator = rows.reduce((sum, row) => sum + numberOrZero(row[numeratorKey]), 0);
          const denominator = rows.reduce((sum, row) => sum + numberOrZero(row[denominatorKey]), 0);
          return denominator > 0 ? numerator / denominator : 0;
        })(),
      };
    })
    .sort((a, b) => b.euasRequired - a.euasRequired);

  return {
    totalEuasRequired,
    totalEuasCost,
    complianceBalance,
    totalPenalty,
    averageIntensity,
    totalFuelConsumed,
    voyageRows,
    portStayRows,
    byVessel,
    ghgKey,
  };
}

function projectionIsActive() {
  const { waspFactor, bioBlend, rfnboBlend } = stateStore.ui.projection;
  return Math.abs(waspFactor - PROJECTION_BASELINE.waspFactor) > 0.0001 || bioBlend > 0 || rfnboBlend > 0;
}

function findFuelReference(name) {
  const query = lower(name);
  return stateStore.derived.fuelReference.find(
    (row) => lower(row.alias) === query || lower(row.fuelPathway) === query || lower(`${row.fuelPathway} ${row.fuelClass}`) === query
  ) || null;
}

function deltaTone(delta, lowerIsBetter = true) {
  if (Math.abs(delta) < 0.0001) return "same";
  const improved = lowerIsBetter ? delta < 0 : delta > 0;
  return improved ? "better" : "worse";
}

function deltaArrow(delta, lowerIsBetter = true) {
  if (Math.abs(delta) < 0.0001) return "•";
  const improved = lowerIsBetter ? delta < 0 : delta > 0;
  return improved ? "▼" : "▲";
}

function formatDelta(value, digits = 1, unit = "", lowerIsBetter = true) {
  const tone = deltaTone(value, lowerIsBetter);
  const arrow = deltaArrow(value, lowerIsBetter);
  const absValue = formatNumber(Math.abs(value), digits);
  return {
    tone,
    arrow,
    text: `${arrow} ${absValue}${unit ? ` ${unit}` : ""}`,
  };
}

function renderSectionBadge(label, tone = "dark") {
  return `<span class="section-badge ${tone}">${label}</span>`;
}

function renderCollapsibleSection({ action, title, badges = "", note = "", open = true, body = "" }) {
  return `
    <section class="compact-section">
      <button class="section-bar" type="button" data-action="${action}" aria-expanded="${open}">
        <span class="section-bar-main">
          <span class="section-chevron ${open ? "open" : ""}">⌄</span>
          <span class="section-title-group">
            <strong>${title}</strong>
            <span class="section-badges">${badges}</span>
          </span>
        </span>
        ${note ? `<span class="section-bar-note">${note}</span>` : ""}
      </button>
      <div class="section-body ${open ? "" : "collapsed"}">
        ${body}
      </div>
    </section>
  `;
}

function getProjectionRows(activeRows) {
  const scenario = stateStore.ui.projection;
  const params = stateStore.derived.parameterValues;
  const bioRef = findFuelReference(scenario.bioType);
  const rfnboRef = findFuelReference(scenario.rfnboType);

  return activeRows.map((row) => {
    const scope = numberOrZero(row.scopePercent);
    const actualWind = row.windFactor === null || row.windFactor === undefined ? PROJECTION_BASELINE.waspFactor : numberOrZero(row.windFactor) || PROJECTION_BASELINE.waspFactor;
    const projectedWind = projectionIsActive() ? numberOrZero(scenario.waspFactor) || PROJECTION_BASELINE.waspFactor : actualWind;
    const totalFossilMt = numberOrZero(row.fuel1ConsumptionMt) + numberOrZero(row.fuel2ConsumptionMt);
    const fossilEnergyMj = numberOrZero(row.fuel1EnergyMj) + numberOrZero(row.fuel2EnergyMj);
    const fossilWtwG = numberOrZero(row.fuel1WtwG) + numberOrZero(row.fuel2WtwG);
    const fossilCfAvg = totalFossilMt > 0 && scope > 0 ? numberOrZero(row.fossilCo2InScopeT) / (totalFossilMt * scope) : 0;
    const fossilWtwAvg = fossilEnergyMj > 0 ? fossilWtwG / fossilEnergyMj : 0;
    const actualBioEnergyMj = numberOrZero(row.bioEnergyMj);
    const actualBioMt = numberOrZero(row.bioFuelConsumptionMt);
    const actualBioWtwAvg = actualBioEnergyMj > 0 ? numberOrZero(row.bioWtwG) / actualBioEnergyMj : numberOrZero(bioRef?.wtWPerMj);
    const actualBioCf = actualBioMt > 0 ? (numberOrZero(row.etsInScopeCo2eqT) - numberOrZero(row.fossilCo2InScopeT)) / Math.max(actualBioMt * Math.max(scope, 1), 1) : numberOrZero(findFuelReference(row.bioFuelType)?.etsCo2Cf);
    const actualBioRwd = actualBioEnergyMj > 0 && scope > 0 ? Math.max(1, (numberOrZero(row.fuelEuDenomStep1Mj) - numberOrZero(row.opsElectricityInScopeMj) - fossilEnergyMj * scope) / (actualBioEnergyMj * scope) || 1) : Math.max(1, numberOrZero(findFuelReference(row.bioFuelType)?.rwd) || 1);
    const bioBlend = Math.min(0.5, numberOrZero(scenario.bioBlend) / 100);
    const rfnboBlend = Math.min(0.3, numberOrZero(scenario.rfnboBlend) / 100);
    const combinedBlend = Math.min(0.8, bioBlend + rfnboBlend);
    const fossilRetainedShare = Math.max(0, 1 - combinedBlend);
    const newBioMt = totalFossilMt * bioBlend;
    const newRfnboMt = totalFossilMt * rfnboBlend;
    const newBioEnergyMj = fossilEnergyMj * bioBlend;
    const newRfnboEnergyMj = fossilEnergyMj * rfnboBlend;
    const retainedFossilEnergyMj = fossilEnergyMj * fossilRetainedShare;
    const retainedFossilMt = totalFossilMt * fossilRetainedShare;
    const bioWtW = numberOrZero(bioRef?.wtWPerMj) || actualBioWtwAvg;
    const bioCf = numberOrZero(bioRef?.etsCo2Cf);
    const bioRwd = Math.max(1, numberOrZero(bioRef?.rwd) || 1);
    const rfnboWtW = numberOrZero(rfnboRef?.wtWPerMj);
    const rfnboCf = numberOrZero(rfnboRef?.etsCo2Cf);
    const rfnboRwd = Math.max(1, numberOrZero(rfnboRef?.rwd) || 2);
    const projectedPreWindWtw =
      retainedFossilEnergyMj * fossilWtwAvg +
      actualBioEnergyMj * actualBioWtwAvg +
      newBioEnergyMj * bioWtW +
      newRfnboEnergyMj * rfnboWtW;
    const projectedFuelEuWtwEmissionsG = projectedPreWindWtw * projectedWind * scope + numberOrZero(row.elecWtwG);
    const projectedFuelEuDenomStep1Mj =
      (retainedFossilEnergyMj + actualBioEnergyMj * actualBioRwd + newBioEnergyMj * bioRwd + newRfnboEnergyMj * rfnboRwd) * scope +
      numberOrZero(row.opsElectricityInScopeMj);
    const projectedFuelEuEnergyStep2Mj = numberOrZero(row.fuelEuEnergyStep2Mj);
    const projectedAttainedGhgIntensity =
      projectedFuelEuDenomStep1Mj > 0 ? projectedFuelEuWtwEmissionsG / projectedFuelEuDenomStep1Mj : null;
    const projectedComplianceBalanceT =
      projectedFuelEuDenomStep1Mj > 0 && projectedAttainedGhgIntensity !== null
        ? (params.fueleuTarget - projectedAttainedGhgIntensity) * projectedFuelEuEnergyStep2Mj / 1_000_000
        : null;
    const projectedFuelEuPenaltyEur =
      projectedAttainedGhgIntensity && projectedComplianceBalanceT !== null && projectedComplianceBalanceT < 0
        ? Math.abs(projectedComplianceBalanceT) * 1_000_000 / (projectedAttainedGhgIntensity * params.vlsfoMj) * params.penRate * params.penMultiplier
        : 0;
    const projectedEtsInScopeCo2eqT =
      (retainedFossilMt * fossilCfAvg + actualBioMt * Math.max(0, actualBioCf) + newBioMt * bioCf + newRfnboMt * rfnboCf) * scope;
    const projectedEuasRequiredT = projectedEtsInScopeCo2eqT * params.etsPhaseIn;
    const projectedEuasCostEur = projectedEuasRequiredT * params.euaPrice;

    return {
      ...row,
      projectedFuelEuWtwEmissionsG,
      projectedFuelEuDenomStep1Mj,
      projectedFuelEuEnergyStep2Mj,
      projectedAttainedGhgIntensity,
      projectedComplianceBalanceT,
      projectedFuelEuPenaltyEur,
      projectedEtsInScopeCo2eqT,
      projectedEuasRequiredT,
      projectedEuasCostEur,
      deltaEuasRequiredT: projectedEuasRequiredT - numberOrZero(row.euasRequiredT),
      deltaAttainedGhgIntensity: numberOrZero(projectedAttainedGhgIntensity) - numberOrZero(row.attainedGhgIntensity),
      deltaComplianceBalanceT: numberOrZero(projectedComplianceBalanceT) - numberOrZero(row.complianceBalanceT),
    };
  });
}

function renderKpis() {
  const dashboard = computeFilteredDashboard(getActiveRows());
  const projectedRows = getProjectionRows(getActiveRows());
  const projectedDashboard = computeDashboardSummary(projectedRows, {
    euasKey: "projectedEuasRequiredT",
    costKey: "projectedEuasCostEur",
    penaltyKey: "projectedFuelEuPenaltyEur",
    numeratorKey: "projectedFuelEuWtwEmissionsG",
    denominatorKey: "projectedFuelEuDenomStep1Mj",
    energyKey: "projectedFuelEuEnergyStep2Mj",
    balanceKey: "projectedComplianceBalanceT",
    ghgKey: "projectedAttainedGhgIntensity",
  });
  const showProjection = projectionIsActive();
  const cards = [
    {
      key: "total-euas",
      label: "Total EUAs required",
      value: `${formatNumber(dashboard.totalEuasRequired, 1)}`,
      detail: "t CO2eq",
      note: "Click charts for record-level rows",
      tone: "risk",
      projectionTag: formatDelta(projectedDashboard.totalEuasRequired - dashboard.totalEuasRequired, 1, "t", true),
      projectionValue: `${formatNumber(projectedDashboard.totalEuasRequired, 1)} projected`,
    },
    {
      key: "total-cost",
      label: "Total EUA cost",
      value: `${formatCurrency(dashboard.totalEuasCost)}`,
      detail: `@ EUR ${formatInteger(stateStore.derived.parameterValues.euaPrice)} / EUA`,
      note: "Filtered by current vessel",
      tone: "warn",
      projectionTag: formatDelta(projectedDashboard.totalEuasCost - dashboard.totalEuasCost, 0, "", true),
      projectionValue: `${formatCurrency(projectedDashboard.totalEuasCost)} projected`,
    },
    {
      key: "compliance-balance",
      label: "Compliance balance",
      value: `${dashboard.complianceBalance >= 0 ? "+" : ""}${formatNumber(dashboard.complianceBalance, 1)}`,
      detail: "t CO2eq surplus / deficit",
      note: "Based on FuelEU target",
      tone: dashboard.complianceBalance >= 0 ? "good" : "risk",
      projectionTag: formatDelta(projectedDashboard.complianceBalance - dashboard.complianceBalance, 1, "t", false),
      projectionValue: `${projectedDashboard.complianceBalance >= 0 ? "+" : ""}${formatNumber(projectedDashboard.complianceBalance, 1)} projected`,
    },
    {
      key: "penalty",
      label: "FuelEU penalty",
      value: `${formatCurrency(dashboard.totalPenalty)}`,
      detail: dashboard.totalPenalty > 0 ? "Penalty triggered by deficits" : "No penalty due",
      note: "Calculator-derived total",
      tone: dashboard.totalPenalty > 0 ? "risk" : "good",
      projectionTag: formatDelta(projectedDashboard.totalPenalty - dashboard.totalPenalty, 0, "", true),
      projectionValue: `${formatCurrency(projectedDashboard.totalPenalty)} projected`,
    },
    {
      key: "avg-ghg",
      label: "Avg GHG intensity",
      value: `${formatNumber(dashboard.averageIntensity, 2)}`,
      detail: `g/MJ vs ${formatNumber(stateStore.derived.parameterValues.fueleuTarget, 2)} target`,
      note: "Voyage and port-stay weighted",
      tone: dashboard.averageIntensity <= stateStore.derived.parameterValues.fueleuTarget ? "good" : "warn",
      projectionTag: formatDelta(projectedDashboard.averageIntensity - dashboard.averageIntensity, 2, "g/MJ", true),
      projectionValue: `${formatNumber(projectedDashboard.averageIntensity, 2)} projected`,
    },
    {
      key: "fuel-consumed",
      label: "Total fuel consumed",
      value: `${formatNumber(dashboard.totalFuelConsumed, 1)}`,
      detail: "MT all fuel types",
      note: "Fossil plus biofuel",
      tone: "neutral",
      projectionTag: formatDelta(0, 1, "MT", true),
      projectionValue: "Energy held constant",
    },
    {
      key: "voyages",
      label: "Voyage records",
      value: `${formatInteger(dashboard.voyageRows.length)}`,
      detail: "Click charts to view voyages",
      note: "Current filter only",
      tone: "neutral",
      projectionTag: formatDelta(0, 0, "", true),
      projectionValue: "Count unchanged",
    },
    {
      key: "port-stays",
      label: "Port stay records",
      value: `${formatInteger(dashboard.portStayRows.length)}`,
      detail: "Click charts to view port stays",
      note: "Current filter only",
      tone: "neutral",
      projectionTag: formatDelta(0, 0, "", true),
      projectionValue: "Count unchanged",
    },
  ];

  const cardsHtml = cards
    .map(
      (card) => `
        <button class="kpi-card tone-${card.tone}" type="button" data-action="open-kpi-drilldown" data-kpi="${card.key}">
          <div class="kpi-label">${card.label}</div>
          <div class="kpi-value">${card.value}</div>
          <div class="kpi-detail">${card.detail}</div>
          <div class="kpi-note">${card.note}</div>
          ${
            showProjection
              ? `<div class="kpi-projection-tag ${card.projectionTag.tone}">${card.projectionTag.text}</div><div class="kpi-projection-value">${card.projectionValue}</div>`
              : ""
          }
        </button>
      `
    )
    .join("");

  elements.kpiGrid.innerHTML = renderCollapsibleSection({
    action: "toggle-kpis",
    title: "Key Performance Indicators",
    badges: `${renderSectionBadge("8 metrics")}${showProjection ? renderSectionBadge("Projection active", "purple") : ""}`,
    note: "Click any card to drill into records",
    open: stateStore.ui.kpisOpen,
    body: `<div class="metrics-grid">${cardsHtml}</div>`,
  });
}

function toneClass(value) {
  if (value > 0) return "tag-good";
  if (value < 0) return "tag-risk";
  return "muted";
}

function totalOperationalCostEur(row) {
  return numberOrZero(row.euasCostEur) + numberOrZero(row.fuelEuPenaltyEur);
}

function surplusValueUsd(row) {
  return Math.max(0, numberOrZero(row.complianceBalanceT)) * 215;
}

function monthKeyForRow(row) {
  const raw = row.departureDate || row.arrivalDate || "";
  if (!raw) {
    return "Undated";
  }
  return String(raw).slice(0, 7);
}

function formatMonthLabel(monthKey) {
  if (!monthKey || monthKey === "Undated") {
    return "Undated";
  }
  const [year, month] = monthKey.split("-");
  return new Date(Number(year), Number(month) - 1, 1).toLocaleDateString("en-US", {
    month: "short",
    year: "numeric",
  });
}

function getVoyageRowsForCharts() {
  return [...getActiveRows()]
    .filter((row) => row.type === "Voyage" && row.recordId)
    .sort((left, right) => {
      const leftDate = `${left.departureDate || ""}-${left.recordId}`;
      const rightDate = `${right.departureDate || ""}-${right.recordId}`;
      return leftDate.localeCompare(rightDate);
    });
}

function buildMonthlyAnalytics(rows) {
  const monthMap = new Map();

  rows.forEach((row) => {
    const key = monthKeyForRow(row);
    if (!monthMap.has(key)) {
      monthMap.set(key, {
        monthKey: key,
        monthLabel: formatMonthLabel(key),
        rows: [],
        eua: 0,
        euaCost: 0,
        penalty: 0,
        totalCost: 0,
      });
    }

    const bucket = monthMap.get(key);
    bucket.rows.push(row);
    bucket.eua += numberOrZero(row.euasRequiredT);
    bucket.euaCost += numberOrZero(row.euasCostEur);
    bucket.penalty += numberOrZero(row.fuelEuPenaltyEur);
    bucket.totalCost += totalOperationalCostEur(row);
  });

  return [...monthMap.values()].sort((left, right) => left.monthKey.localeCompare(right.monthKey));
}

function renderEuaMarketStrip() {
  const { status, snapshot, error } = stateStore.market;
  const loading = status === "loading" || status === "idle";
  const failed = status === "error";

  return `
    <section class="market-strip">
      <div class="market-strip-copy">
        <p class="eyebrow">Live EUA Market</p>
        <h2>EU Carbon Permits</h2>
        <p class="helper-text">Live benchmark reference from Trading Economics for dashboard context.</p>
      </div>
      <div class="market-strip-metrics">
        <article class="market-card primary">
          <span class="market-label">Current price</span>
          <strong class="market-value">${loading ? "Loading..." : failed ? "-" : `${formatNumber(snapshot.price, 2)} EUR`}</strong>
          <span class="market-note">${loading ? "Fetching market feed" : failed ? error || "Market feed unavailable" : snapshot.asOfDate || "Latest session"}</span>
        </article>
        <article class="market-card">
          <span class="market-label">Previous day</span>
          <strong class="market-value">${loading || failed ? "-" : `${formatNumber(snapshot.previous, 2)} EUR`}</strong>
          <span class="market-note">${loading || failed ? "" : `${snapshot.dayChangePercent >= 0 ? "+" : ""}${formatNumber(snapshot.dayChangePercent, 2)}% day move`}</span>
        </article>
        <article class="market-card">
          <span class="market-label">1 month</span>
          <strong class="market-value ${loading || failed ? "" : snapshot.monthChangePercent <= 0 ? "tag-risk" : "tag-good"}">${loading || failed ? "-" : `${snapshot.monthChangePercent >= 0 ? "+" : ""}${formatNumber(snapshot.monthChangePercent, 2)}%`}</strong>
          <span class="market-note">Monthly performance</span>
        </article>
        <article class="market-card">
          <span class="market-label">1 year</span>
          <strong class="market-value ${loading || failed ? "" : snapshot.yearChangePercent >= 0 ? "tag-good" : "tag-risk"}">${loading || failed ? "-" : `${snapshot.yearChangePercent >= 0 ? "+" : ""}${formatNumber(snapshot.yearChangePercent, 2)}%`}</strong>
          <span class="market-note"><a href="${snapshot?.sourceUrl || "https://tradingeconomics.com/commodity/carbon"}" target="_blank" rel="noreferrer">Open source</a></span>
        </article>
      </div>
    </section>
  `;
}

function calculatorCellValue(row, column) {
  const value = row[column.key];
  if (column.kind === "calculated-currency") {
    return formatCurrency(value);
  }
  if (column.kind === "calculated-percent") {
    return formatPercent(value);
  }
  if (column.kind === "calculated-number" || column.kind === "sticky-number") {
    return formatNumber(value, column.digits ?? 2);
  }
  if (column.kind === "editable-number") {
    return value ?? "";
  }
  return value ?? "";
}

function calculatorInputValue(inputRow, derivedRow, column) {
  const inputValue = inputRow[column.key];
  if (inputValue !== null && inputValue !== undefined && inputValue !== "") {
    return inputValue;
  }
  const derivedValue = derivedRow[column.key];
  return derivedValue ?? "";
}

function getCalculatorDatalistId(rowId, columnKey) {
  return `calc-list-${rowId}-${columnKey}`;
}

function getSuggestionChoices(sourceKey, rawValue) {
  const query = lower(rawValue);
  if (sourceKey === "portCodes") {
    return stateStore.state.ports
      .filter((row) => !query || lower(`${row.unlocode} ${row.portName} ${row.country}`).includes(query))
      .slice(0, 40)
      .map((row) => ({ value: row.unlocode, label: row.portName }));
  }
  if (sourceKey === "fuelTypes") {
    return stateStore.derived.fuelReference
      .filter((row) => !query || lower(`${row.fuelPathway} ${row.fuelClass}`).includes(query))
      .slice(0, 25)
      .map((row) => ({ value: row.fuelPathway, label: row.fuelClass }));
  }
  if (sourceKey === "imoNumbers") {
    return stateStore.state.fleet
      .filter((row) => !query || lower(`${row.imoNo} ${row.vesselName}`).includes(query))
      .slice(0, 25)
      .map((row) => ({ value: row.imoNo, label: row.vesselName }));
  }
  return [];
}

function renderSuggestionDatalist(rowId, column, rawValue) {
  if (!column.list) {
    return "";
  }
  const datalistId = getCalculatorDatalistId(rowId, column.key);
  const options = getSuggestionChoices(column.list, rawValue);
  return `
    <datalist id="${datalistId}">
      ${options.map((option) => `<option value="${option.value}">${option.label || option.value}</option>`).join("")}
    </datalist>
  `;
}

function updateInlineSuggestions(input, columnKey, rawValue) {
  const column = CALCULATOR_COLUMNS.find((item) => item.key === columnKey);
  if (!column?.list) {
    return;
  }
  const datalist = document.getElementById(getCalculatorDatalistId(input.dataset.rowId, columnKey));
  if (!datalist) {
    return;
  }
  const options = getSuggestionChoices(column.list, rawValue);
  datalist.innerHTML = options.map((option) => `<option value="${option.value}">${option.label || option.value}</option>`).join("");
}

function renderCalculatorCell(row, inputRow, column, stickyLeft, context = "active") {
  const classes = ["calculator-cell"];
  const styleParts = [`min-width:${column.width}px`, `width:${column.width}px`];

  if (column.kind.includes("sticky")) {
    classes.push("sticky-cell");
    styleParts.push(`left:${stickyLeft}px`);
  }

  if (column.kind.includes("editable")) {
    classes.push("editable-cell");
  }

  if (column.kind.includes("calculated")) {
    classes.push("calculated-cell");
  }

  if (column.kind.includes("number") || column.kind.includes("currency") || column.key === "rowActions") {
    classes.push("number-cell");
  }

  if (column.key === "rowActions") {
    if (context === "history") {
      return `
        <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
          <div class="history-row-actions">
            <button class="inline-button compact-button" type="button" data-action="insert-calculator-after" data-row-id="${row.id}">Insert below</button>
            <button class="inline-button compact-button button-red" type="button" data-action="delete-calculator-row" data-row-id="${row.id}">Delete</button>
          </div>
        </td>
      `;
    }
    return `
      <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
        <button class="inline-button compact-button" type="button" data-action="delete-calculator-row" data-row-id="${row.id}">Delete</button>
      </td>
    `;
  }

  if (column.kind.includes("editable")) {
    if (column.input === "select") {
      return `
        <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
          <select class="calculator-grid-input input-orange" data-calc-cell="${column.key}" data-row-id="${row.id}" data-row-context="${context}">
            ${(column.options || [])
              .map((option) => `<option value="${option}" ${calculatorInputValue(inputRow, row, column) === option ? "selected" : ""}>${option}</option>`)
              .join("")}
          </select>
        </td>
      `;
    }

    return `
      <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
        <input
          class="calculator-grid-input input-orange"
          data-calc-cell="${column.key}"
          data-row-id="${row.id}"
          data-row-context="${context}"
          type="${column.input}"
          value="${calculatorInputValue(inputRow, row, column)}"
          ${column.placeholder ? `placeholder="${column.placeholder}"` : ""}
          ${column.list ? `list="${getCalculatorDatalistId(row.id, column.key)}"` : ""}
          ${column.step ? `step="${column.step}"` : ""}
          autocomplete="off"
        >
        ${renderSuggestionDatalist(row.id, column, calculatorInputValue(inputRow, row, column))}
      </td>
    `;
  }

  return `
    <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
      ${calculatorCellValue(row, column)}
    </td>
  `;
}

function openDrilldown(title, subtitle, columns, rows, sourceRows = []) {
  stateStore.ui.drilldown = {
    title,
    subtitle,
    columns,
    rows,
    sourceRecordIds: sourceRows.map((row) => row.recordId).filter(Boolean),
  };
  render();
}

function closeDrilldown() {
  stateStore.ui.drilldown = null;
  render();
}

function openKpiDrilldown(kpiKey) {
  const activeRows = getActiveRows();
  const dashboard = computeFilteredDashboard(activeRows);

  if (kpiKey === "total-euas") {
    openDrilldown(
      "Total EUAs required",
      "Rows contributing to the current EUA total.",
      ["Record", "Vessel", "Route", "Type", "EUAs Required", "ETS CO2eq"],
      activeRows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        row.type,
        formatNumber(row.euasRequiredT, 3),
        formatNumber(row.etsInScopeCo2eqT, 3),
      ]),
      activeRows
    );
    return;
  }

  if (kpiKey === "total-cost") {
    openDrilldown(
      "Total EUA cost",
      "Rows contributing to the current ETS cost.",
      ["Record", "Vessel", "Route", "EUAs Required", "ETS Cost"],
      activeRows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        formatNumber(row.euasRequiredT, 3),
        formatCurrency(row.euasCostEur),
      ]),
      activeRows
    );
    return;
  }

  if (kpiKey === "compliance-balance") {
    openDrilldown(
      "Compliance balance",
      "FuelEU balance by record for the current filter.",
      ["Record", "Vessel", "Route", "GHG Intensity", "Target", "Compliance Balance"],
      activeRows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        formatNumber(row.attainedGhgIntensity, 3),
        formatNumber(row.targetGhgIntensity, 3),
        formatNumber(row.complianceBalanceT, 3),
      ]),
      activeRows
    );
    return;
  }

  if (kpiKey === "penalty") {
    const penaltyRows = activeRows.filter((row) => numberOrZero(row.fuelEuPenaltyEur) > 0);
    openDrilldown(
      "FuelEU penalty",
      "Penalty-bearing rows for the current filter.",
      ["Record", "Vessel", "Route", "Compliance Balance", "Penalty"],
      penaltyRows.map((row) => [
          row.recordId,
          row.vesselName,
          row.route,
          formatNumber(row.complianceBalanceT, 3),
          formatCurrency(row.fuelEuPenaltyEur),
        ]),
      penaltyRows
    );
    return;
  }

  if (kpiKey === "avg-ghg") {
    openDrilldown(
      "Average GHG intensity",
      `Fleet average ${formatNumber(dashboard.averageIntensity, 3)} g/MJ against target ${formatNumber(stateStore.derived.parameterValues.fueleuTarget, 3)} g/MJ.`,
      ["Record", "Vessel", "Route", "Type", "GHG Intensity", "Target"],
      activeRows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        row.type,
        formatNumber(row.attainedGhgIntensity, 3),
        formatNumber(row.targetGhgIntensity, 3),
      ]),
      activeRows
    );
    return;
  }

  if (kpiKey === "fuel-consumed") {
    openDrilldown(
      "Total fuel consumed",
      "Fuel consumption split by record.",
      ["Record", "Vessel", "Fuel 1 MT", "Fuel 2 MT", "Bio MT", "Total MT"],
      activeRows.map((row) => {
        const total = numberOrZero(row.fuel1ConsumptionMt) + numberOrZero(row.fuel2ConsumptionMt) + numberOrZero(row.bioFuelConsumptionMt);
        return [
          row.recordId,
          row.vesselName,
          formatNumber(row.fuel1ConsumptionMt, 2),
          formatNumber(row.fuel2ConsumptionMt, 2),
          formatNumber(row.bioFuelConsumptionMt, 2),
          formatNumber(total, 2),
        ];
      }),
      activeRows
    );
    return;
  }

  if (kpiKey === "voyages") {
    openDrilldown(
      "Voyage records",
      "Current filtered voyage rows.",
      ["Record", "Vessel", "Route", "EUAs", "GHG Intensity", "ETS Cost"],
      dashboard.voyageRows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        formatNumber(row.euasRequiredT, 3),
        formatNumber(row.attainedGhgIntensity, 3),
        formatCurrency(row.euasCostEur),
      ]),
      dashboard.voyageRows
    );
    return;
  }

  if (kpiKey === "port-stays") {
    openDrilldown(
      "Port stay records",
      "Current filtered port stay rows.",
      ["Record", "Vessel", "Route", "EUAs", "GHG Intensity", "ETS Cost"],
      dashboard.portStayRows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        formatNumber(row.euasRequiredT, 3),
        formatNumber(row.attainedGhgIntensity, 3),
        formatCurrency(row.euasCostEur),
      ]),
      dashboard.portStayRows
    );
  }
}

function renderDrilldownPane() {
  const drilldown = stateStore.ui.drilldown;
  if (!drilldown) {
    return `
      <aside class="drilldown-pane empty">
        <div class="drilldown-empty">
          <p class="eyebrow">Chart Drilldown</p>
          <h3>Click any chart</h3>
          <p class="helper-text">The selected bar, point, or segment will open its related records here as a readable table.</p>
        </div>
      </aside>
    `;
  }

  return `
    <aside class="drilldown-pane">
      <div class="table-head">
        <div>
          <p class="eyebrow">Chart Drilldown</p>
          <h3>${drilldown.title}</h3>
          <p class="helper-text">${drilldown.subtitle || ""}</p>
        </div>
        <button class="icon-button compact" type="button" data-action="close-drilldown" title="Close drilldown">✕</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>${drilldown.columns.map((column) => `<th>${column}</th>`).join("")}</tr>
          </thead>
          <tbody>
            ${drilldown.rows
              .map(
                (row) => `
                  <tr>${row.map((value) => `<td>${value}</td>`).join("")}</tr>
                `
              )
              .join("") || `<tr><td colspan="${drilldown.columns.length}">No records</td></tr>`}
          </tbody>
        </table>
      </div>
    </aside>
  `;
}

function getDetailRows() {
  const projectedRows = getProjectionRows(getActiveRows());
  const drilldownRecordIds = stateStore.ui.drilldown?.sourceRecordIds || [];
  const searchTerm = lower(stateStore.ui.detailSearch);

  return projectedRows.filter((row) => {
    if (drilldownRecordIds.length && !drilldownRecordIds.includes(row.recordId)) {
      return false;
    }

    if (stateStore.ui.detailScope === "100" && numberOrZero(row.scopePercent) !== 1) {
      return false;
    }
    if (stateStore.ui.detailScope === "50" && numberOrZero(row.scopePercent) !== 0.5) {
      return false;
    }
    if (stateStore.ui.detailScope === "0" && numberOrZero(row.scopePercent) !== 0) {
      return false;
    }

    if (!searchTerm) {
      return true;
    }

    return DETAIL_TABLE_COLUMNS.filter((column) => !column.proj).some((column) => lower(row[column.key]).includes(searchTerm));
  });
}

function renderDetailCell(row, column) {
  const value = row[column.key];

  if (column.format === "type") {
    const tone = row.type === "Voyage" ? "voyage" : row.type === "Port Stay" ? "port" : "other";
    const label = row.type === "Port Stay" ? "Port" : row.type || "-";
    return `<span class="detail-pill ${tone}">${label}</span>`;
  }

  if (column.format === "scope") {
    const scope = numberOrZero(value);
    if (scope === 1) return `<span class="detail-pill scope-full">100%</span>`;
    if (scope === 0.5) return `<span class="detail-pill scope-half">50%</span>`;
    if (scope === 0) return `<span class="detail-pill scope-out">Out</span>`;
    return "-";
  }

  if (column.format === "date") {
    return formatDateValue(value);
  }

  if (column.format === "currency") {
    return `<span class="detail-value number-cell">${formatCurrency(value)}</span>`;
  }

  if (column.format === "number") {
    return `<span class="detail-value number-cell">${formatNumber(value, column.digits || 0)}</span>`;
  }

  if (column.format === "ghg") {
    const good = numberOrZero(row.attainedGhgIntensity) <= numberOrZero(row.targetGhgIntensity);
    return `<span class="detail-value number-cell ${good ? "tag-good" : "tag-risk"}">${formatNumber(value, column.digits || 0)}</span>`;
  }

  if (column.format === "balance") {
    return `<span class="detail-value number-cell ${numberOrZero(value) >= 0 ? "tag-good" : "tag-risk"}">${numberOrZero(value) >= 0 ? "+" : ""}${formatNumber(value, column.digits || 0)}</span>`;
  }

  if (column.format?.startsWith("projection")) {
    const delta = numberOrZero(row[column.projectionKey]);
    const lowerIsBetter = column.key !== "projectedComplianceBalanceT";
    const deltaMeta = formatDelta(delta, 1, "", lowerIsBetter);
    const valueTone =
      column.format === "projection-balance"
        ? numberOrZero(value) >= 0
          ? "tag-good"
          : "tag-risk"
        : column.format === "projection-ghg"
          ? numberOrZero(value) <= numberOrZero(row.targetGhgIntensity)
            ? "tag-good"
            : "tag-risk"
          : "";
    const prefix = column.format === "projection-balance" && numberOrZero(value) >= 0 ? "+" : "";
    return `
      <span class="detail-value number-cell ${valueTone}">
        ${prefix}${formatNumber(value, column.digits || 0)}
        <span class="projection-delta ${deltaMeta.tone}">${deltaMeta.text}</span>
      </span>
    `;
  }

  if (column.format === "note") {
    return `<span class="detail-note" title="${value || ""}">${value || "-"}</span>`;
  }

  return value || "-";
}

function renderUnifiedDetailSection() {
  const filteredRows = getDetailRows();
  const showProjection = projectionIsActive();
  const visibleColumns = showProjection ? DETAIL_TABLE_COLUMNS : DETAIL_TABLE_COLUMNS.filter((column) => !column.proj);
  const detailTableWidth = visibleColumns.reduce((sum, column) => sum + (column.width || 120), 0);
  const drilldown = stateStore.ui.drilldown;
  const totalEuas = filteredRows.reduce((sum, row) => sum + numberOrZero(row.euasRequiredT), 0);
  const totalCost = filteredRows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0);
  const totalBalance = filteredRows.reduce((sum, row) => sum + numberOrZero(row.complianceBalanceT), 0);
  const projectedEuas = filteredRows.reduce((sum, row) => sum + numberOrZero(row.projectedEuasRequiredT), 0);
  const projectedBalance = filteredRows.reduce((sum, row) => sum + numberOrZero(row.projectedComplianceBalanceT), 0);

  return renderCollapsibleSection({
    action: "toggle-voyage-table",
    title: "Voyage / Port-Stay Detail",
    badges: `${renderSectionBadge(`${filteredRows.length} rows`)}${showProjection ? renderSectionBadge("Projected columns", "purple") : ""}`,
    note: "Charts, KPI clicks, and search all feed this one section",
    open: stateStore.ui.voyageTableOpen,
    body: `
      <article class="table-card compact-table-card detail-table-card">
        <div class="table-head detail-table-head">
          <div>
            <p class="eyebrow">Unified Detail View</p>
            <h3>${drilldown ? drilldown.title : "All filtered records"}</h3>
            <p class="helper-text">${drilldown?.subtitle || "Click a KPI or chart to focus this same table, or use the search and scope filters below."}</p>
          </div>
          ${drilldown ? `<button class="inline-button compact-button" type="button" data-action="close-drilldown">Clear focus</button>` : ""}
        </div>

        <div class="detail-toolbar">
          <div class="toolbar-field">
            <span>Search</span>
            <input
              class="search-input"
              type="search"
              data-action="detail-search"
              value="${stateStore.ui.detailSearch}"
              placeholder="Search ID, vessel, port, fuel, note, or type"
            >
          </div>
          <div class="toolbar-field detail-scope-field">
            <span>Scope</span>
            <select class="toolbar-select" data-action="detail-scope">
              <option value="all" ${stateStore.ui.detailScope === "all" ? "selected" : ""}>All scopes</option>
              <option value="100" ${stateStore.ui.detailScope === "100" ? "selected" : ""}>100% in-scope</option>
              <option value="50" ${stateStore.ui.detailScope === "50" ? "selected" : ""}>50% in-scope</option>
              <option value="0" ${stateStore.ui.detailScope === "0" ? "selected" : ""}>Out of scope</option>
            </select>
          </div>
        </div>

        <div class="detail-summary">
          <span class="chip">${filteredRows.length} visible rows</span>
          <span class="chip">EUAs ${formatNumber(totalEuas, 1)} t</span>
          <span class="chip">Cost ${formatCurrency(totalCost)}</span>
          <span class="chip ${totalBalance >= 0 ? "tag-good" : "tag-risk"}">Balance ${totalBalance >= 0 ? "+" : ""}${formatNumber(totalBalance, 1)} t</span>
          ${showProjection ? `<span class="chip projection-chip">Projected EUAs ${formatNumber(projectedEuas, 1)} t</span><span class="chip projection-chip ${projectedBalance >= 0 ? "tag-good" : "tag-risk"}">Projected balance ${projectedBalance >= 0 ? "+" : ""}${formatNumber(projectedBalance, 1)} t</span>` : ""}
        </div>

        <div class="table-wrap dense-table-wrap detail-table-wrap" data-scroll-group="detail">
          <table class="detail-table" style="width:${detailTableWidth}px;min-width:${detailTableWidth}px">
            <thead>
              <tr>
                ${visibleColumns
                  .map(
                    (column) => `
                      <th style="min-width:${column.width}px;width:${column.width}px" class="${column.proj ? "projection-col-head" : ""}">
                        ${column.proj ? "⚡ " : ""}${column.label}
                      </th>
                    `
                  )
                  .join("")}
              </tr>
            </thead>
            <tbody>
              ${
                filteredRows.length
                  ? filteredRows
                      .map(
                        (row) => `
                          <tr>
                            ${visibleColumns
                              .map(
                                (column) => `
                                  <td class="${["number", "currency", "ghg", "balance"].includes(column.format) || column.format?.startsWith("projection") ? "detail-number-cell" : ""}">
                                    ${renderDetailCell(row, column)}
                                  </td>
                                `
                              )
                              .join("")}
                          </tr>
                        `
                      )
                      .join("")
                  : `<tr><td colspan="${visibleColumns.length}" class="no-data">No matching records for the current chart focus and filters.</td></tr>`
              }
            </tbody>
          </table>
        </div>
      </article>
    `,
  });
}

function renderProjectionPanel() {
  const activeRows = getActiveRows();
  const projectedRows = getProjectionRows(activeRows);
  const baseline = computeFilteredDashboard(activeRows);
  const projected = computeDashboardSummary(projectedRows, {
    euasKey: "projectedEuasRequiredT",
    costKey: "projectedEuasCostEur",
    penaltyKey: "projectedFuelEuPenaltyEur",
    numeratorKey: "projectedFuelEuWtwEmissionsG",
    denominatorKey: "projectedFuelEuDenomStep1Mj",
    energyKey: "projectedFuelEuEnergyStep2Mj",
    balanceKey: "projectedComplianceBalanceT",
    ghgKey: "projectedAttainedGhgIntensity",
  });
  const showProjection = projectionIsActive();
  const { waspFactor, bioBlend, bioType, rfnboBlend, rfnboType } = stateStore.ui.projection;
  const summaryItems = [
    ["Projected EUAs", formatNumber(projected.totalEuasRequired, 1), formatDelta(projected.totalEuasRequired - baseline.totalEuasRequired, 1, "t", true)],
    ["Projected cost", formatCurrency(projected.totalEuasCost), formatDelta(projected.totalEuasCost - baseline.totalEuasCost, 0, "", true)],
    ["Compliance balance", `${projected.complianceBalance >= 0 ? "+" : ""}${formatNumber(projected.complianceBalance, 1)}`, formatDelta(projected.complianceBalance - baseline.complianceBalance, 1, "t", false)],
    ["GHG vs baseline", `${formatNumber(projected.averageIntensity, 2)} g/MJ`, formatDelta(projected.averageIntensity - baseline.averageIntensity, 2, "g/MJ", true)],
  ];

  return `
    <section class="projection-section">
      <button class="section-toggle" type="button" data-action="toggle-projection-panel" aria-expanded="${stateStore.ui.projectionOpen}">
        <span class="section-toggle-copy">
          <span class="section-chevron ${stateStore.ui.projectionOpen ? "open" : ""}">⌄</span>
          <span><strong>Compliance Projection Scenarios</strong><small>WASP · Biofuel · RFNBO</small></span>
        </span>
        <span class="section-toggle-note">Adjust sliders and charts update live</span>
      </button>
      <div class="projection-shell ${stateStore.ui.projectionOpen ? "" : "collapsed"}">
        <div class="projection-card">
          <div class="projection-header">
            <h3>Scenario Builder</h3>
            <div class="scenario-btns">
              ${[
                ["baseline", "Baseline"],
                ["wasp10", "WASP 10%"],
                ["wasp20", "WASP 20%"],
                ["bio10", "Bio 10%"],
                ["bio20", "Bio 20%"],
                ["rfnbo10", "RFNBO 10%"],
                ["rfnbo20", "RFNBO 20%"],
                ["combined", "Combined"],
              ]
                .map(
                  ([preset, label]) => `
                    <button class="scenario-chip ${stateStore.ui.projectionPreset === preset ? "active" : ""}" type="button" data-action="apply-projection-preset" data-preset="${preset}">
                      ${label}
                    </button>
                  `
                )
                .join("")}
            </div>
          </div>
          <div class="projection-grid">
            <article class="projection-block">
              <h4>Wind-Assisted Propulsion (WASP)</h4>
              <div class="projection-row"><span>Wind factor (f_wind)</span><strong>${formatNumber(waspFactor, 2)}</strong></div>
              <input class="projection-range" type="range" min="0.70" max="0.95" step="0.01" value="${waspFactor}" data-projection-key="waspFactor">
              <div class="projection-scale"><span>0.70 max wind</span><span>0.95 baseline</span></div>
              <p class="helper-text">Lower f_wind reduces the FuelEU WtW numerator and improves voyage GHG intensity.</p>
            </article>
            <article class="projection-block">
              <h4>Biofuel Blend</h4>
              <div class="projection-row"><span>Blend</span><strong>${bioBlend}%</strong></div>
              <input class="projection-range" type="range" min="0" max="50" step="1" value="${bioBlend}" data-projection-key="bioBlend">
              <div class="projection-scale"><span>0%</span><span>50% blend</span></div>
              <label class="projection-select-wrap">
                <span>Fuel type</span>
                <select class="toolbar-select projection-select" data-projection-key="bioType">
                  ${["Bio-diesel", "HVO"].map((option) => `<option value="${option}" ${bioType === option ? "selected" : ""}>${option}</option>`).join("")}
                </select>
              </label>
              <p class="helper-text">Certified biofuel reduces ETS CO2 directly and also helps the FuelEU intensity outcome.</p>
            </article>
            <article class="projection-block">
              <h4>RFNBO Blend</h4>
              <div class="projection-row"><span>Blend</span><strong>${rfnboBlend}%</strong></div>
              <input class="projection-range" type="range" min="0" max="30" step="1" value="${rfnboBlend}" data-projection-key="rfnboBlend">
              <div class="projection-scale"><span>0%</span><span>30% RFNBO</span></div>
              <label class="projection-select-wrap">
                <span>Fuel type</span>
                <select class="toolbar-select projection-select" data-projection-key="rfnboType">
                  ${["e-diesel", "e-methanol"].map((option) => `<option value="${option}" ${rfnboType === option ? "selected" : ""}>${option}</option>`).join("")}
                </select>
              </label>
              <p class="helper-text">RFNBO uses the reward factor window, making denominator uplift strongest during 2025–2033.</p>
            </article>
          </div>
          <div class="projection-summary-bar ${showProjection ? "active" : ""}">
            ${summaryItems
              .map(
                ([label, value, delta]) => `
                  <div class="projection-summary-item">
                    <div class="projection-summary-label">${label}</div>
                    <div class="projection-summary-value">${value}</div>
                    <div class="projection-summary-delta ${delta.tone}">${delta.text}</div>
                  </div>
                `
              )
              .join("")}
          </div>
        </div>
      </div>
    </section>
  `;
}

function renderDashboard() {
  return `
    ${renderEuaMarketStrip()}
    <section class="dashboard-layout">
      ${renderCollapsibleSection({
        action: "toggle-charts",
        title: "Visual Analytics",
        badges: renderSectionBadge("10 charts"),
        note: "Collapse when you need more space",
        open: stateStore.ui.chartsOpen,
        body: `
          <div class="chart-grid">
          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>EUAs by Voyage (t CO2eq)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="vesselEuaChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Month</p>
                <h3>EUAs by Month (t CO2eq)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="monthlyEuaChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>EUA Cost by Voyage (EUR)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="voyageCostChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Month</p>
                <h3>EUA Cost by Month (EUR)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="monthlyCostChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>Attained GHG Intensity by Voyage (g/MJ)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="voyageGhgChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>Compliance Balance by Voyage (t CO2eq)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="balanceChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>FuelEU Penalty by Voyage (EUR)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="voyagePenaltyChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Month</p>
                <h3>FuelEU Penalty by Month (EUR)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="monthlyPenaltyChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>Surplus Value by Voyage (USD @ 215/t)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="surplusChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>Total Cost by Voyage (EUR)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="totalCostChart"></canvas></div>
          </article>
          </div>
        `,
      })}
      ${renderUnifiedDetailSection()}
    </section>
  `;
}

function renderDashboardCharts() {
  destroyCharts();
  const voyageRows = getVoyageRowsForCharts();
  const monthlyBuckets = buildMonthlyAnalytics(voyageRows);

  const openMonthDrilldown = (title, subtitle, monthBucket) => {
    if (!monthBucket) return;
    openDrilldown(
      title,
      subtitle,
      ["Record", "Vessel", "Route", "Date", "EUAs", "EUA Cost", "Penalty", "Total Cost"],
      monthBucket.rows.map((row) => [
        row.recordId,
        row.vesselName,
        row.route,
        formatDateValue(row.departureDate || row.arrivalDate),
        formatNumber(row.euasRequiredT, 2),
        formatCurrency(row.euasCostEur),
        formatCurrency(row.fuelEuPenaltyEur),
        formatCurrency(totalOperationalCostEur(row)),
      ]),
      monthBucket.rows
    );
  };

  const baseOptions = (onClick, showLegend = false) => ({
    maintainAspectRatio: false,
    responsive: true,
    onClick,
    plugins: {
      legend: {
        display: showLegend,
        position: "bottom",
      },
    },
    scales: {
      y: {
        beginAtZero: true,
      },
    },
  });

  const createChart = (key, canvasId, config) => {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    stateStore.charts[key] = new Chart(canvas, config);
  };

  createChart("voyageEua", "vesselEuaChart", {
    type: "bar",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "EUAs required",
          data: voyageRows.map((row) => row.euasRequiredT),
          backgroundColor: voyageRows.map((row) => (numberOrZero(row.scopePercent) === 1 ? "#d78a1f" : "#e7b25b")),
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const row = voyageRows[elementsClicked[0].index];
      if (!row) return;
      openDrilldown(
        `EUAs for ${row.recordId}`,
        row.route,
        ["Record", "Vessel", "Route", "Scope", "EUAs", "EUA Cost", "GHG"],
        [[
          row.recordId,
          row.vesselName,
          row.route,
          formatPercent(row.scopePercent),
          formatNumber(row.euasRequiredT, 2),
          formatCurrency(row.euasCostEur),
          formatNumber(row.attainedGhgIntensity, 2),
        ]],
        [row]
      );
    }),
  });

  createChart("monthlyEua", "monthlyEuaChart", {
    type: "bar",
    data: {
      labels: monthlyBuckets.map((bucket) => bucket.monthLabel),
      datasets: [
        {
          label: "Monthly EUAs",
          data: monthlyBuckets.map((bucket) => bucket.eua),
          backgroundColor: "#4f8edc",
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const bucket = monthlyBuckets[elementsClicked[0].index];
      openMonthDrilldown(`EUAs in ${bucket.monthLabel}`, "Voyage rows contributing to the selected monthly EUA total.", bucket);
    }),
  });

  createChart("voyageCost", "voyageCostChart", {
    type: "bar",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "EUA cost",
          data: voyageRows.map((row) => row.euasCostEur),
          backgroundColor: "#2f9b72",
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const row = voyageRows[elementsClicked[0].index];
      openDrilldown(
        `EUA Cost for ${row.recordId}`,
        row.route,
        ["Record", "Vessel", "EUAs", "EUA Cost", "Penalty", "Total Cost"],
        [[row.recordId, row.vesselName, formatNumber(row.euasRequiredT, 2), formatCurrency(row.euasCostEur), formatCurrency(row.fuelEuPenaltyEur), formatCurrency(totalOperationalCostEur(row))]],
        [row]
      );
    }),
  });

  createChart("monthlyCost", "monthlyCostChart", {
    type: "line",
    data: {
      labels: monthlyBuckets.map((bucket) => bucket.monthLabel),
      datasets: [
        {
          label: "Monthly EUA cost",
          data: monthlyBuckets.map((bucket) => bucket.euaCost),
          borderColor: "#167f5b",
          backgroundColor: "rgba(22, 127, 91, 0.16)",
          pointRadius: 4,
          borderWidth: 3,
          tension: 0.25,
          fill: true,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const bucket = monthlyBuckets[elementsClicked[0].index];
      openMonthDrilldown(`EUA Cost in ${bucket.monthLabel}`, "Voyage rows contributing to the selected monthly EUA cost.", bucket);
    }),
  });

  createChart("voyageGhg", "voyageGhgChart", {
    type: "line",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "Attained",
          data: voyageRows.map((row) => row.attainedGhgIntensity),
          borderColor: "#4288d6",
          backgroundColor: "#4288d6",
          pointRadius: 4,
          tension: 0.2,
        },
        {
          label: `Target ${formatNumber(stateStore.derived.parameterValues.fueleuTarget, 2)}`,
          data: voyageRows.map(() => stateStore.derived.parameterValues.fueleuTarget),
          borderColor: "#cf4e3a",
          borderDash: [6, 6],
          pointRadius: 0,
          tension: 0,
        },
      ],
    },
    options: {
      ...baseOptions((_, elementsClicked) => {
        if (!elementsClicked.length) return;
        const row = voyageRows[elementsClicked[0].index];
        openDrilldown(
          `GHG Intensity for ${row.recordId}`,
          row.route,
          ["Record", "Vessel", "Attained", "Target", "Compliance Balance", "FuelEU Penalty"],
          [[row.recordId, row.vesselName, formatNumber(row.attainedGhgIntensity, 2), formatNumber(row.targetGhgIntensity, 2), formatNumber(row.complianceBalanceT, 2), formatCurrency(row.fuelEuPenaltyEur)]],
          [row]
        );
      }, true),
      scales: {
        y: (() => {
          const values = voyageRows
            .flatMap((row) => [numberOrZero(row.attainedGhgIntensity), numberOrZero(row.targetGhgIntensity)])
            .filter((value) => Number.isFinite(value) && value > 0);
          if (!values.length) {
            return { beginAtZero: true };
          }
          const minValue = Math.min(...values);
          const maxValue = Math.max(...values);
          const padding = Math.max(0.2, (maxValue - minValue) * 0.25 || 0.4);
          return {
            beginAtZero: false,
            min: Math.max(0, minValue - padding),
            max: maxValue + padding,
          };
        })(),
      },
    },
  });

  createChart("balance", "balanceChart", {
    type: "bar",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "Compliance balance",
          data: voyageRows.map((row) => row.complianceBalanceT),
          backgroundColor: voyageRows.map((row) => (numberOrZero(row.complianceBalanceT) >= 0 ? "#2b8a3e" : "#cf4e3a")),
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const row = voyageRows[elementsClicked[0].index];
      openDrilldown(
        `Compliance Balance for ${row.recordId}`,
        row.route,
        ["Record", "Vessel", "Balance", "Surplus Value", "Penalty", "GHG"],
        [[row.recordId, row.vesselName, formatNumber(row.complianceBalanceT, 2), formatUsdCurrency(surplusValueUsd(row)), formatCurrency(row.fuelEuPenaltyEur), formatNumber(row.attainedGhgIntensity, 2)]],
        [row]
      );
    }),
  });

  createChart("voyagePenalty", "voyagePenaltyChart", {
    type: "bar",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "FuelEU penalty",
          data: voyageRows.map((row) => row.fuelEuPenaltyEur),
          backgroundColor: "#c35a4b",
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const row = voyageRows[elementsClicked[0].index];
      openDrilldown(
        `FuelEU Penalty for ${row.recordId}`,
        row.route,
        ["Record", "Vessel", "Compliance Balance", "Penalty", "Total Cost"],
        [[row.recordId, row.vesselName, formatNumber(row.complianceBalanceT, 2), formatCurrency(row.fuelEuPenaltyEur), formatCurrency(totalOperationalCostEur(row))]],
        [row]
      );
    }),
  });

  createChart("monthlyPenalty", "monthlyPenaltyChart", {
    type: "line",
    data: {
      labels: monthlyBuckets.map((bucket) => bucket.monthLabel),
      datasets: [
        {
          label: "Monthly FuelEU penalty",
          data: monthlyBuckets.map((bucket) => bucket.penalty),
          borderColor: "#af3d33",
          backgroundColor: "rgba(175, 61, 51, 0.12)",
          pointRadius: 4,
          borderWidth: 3,
          tension: 0.25,
          fill: true,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const bucket = monthlyBuckets[elementsClicked[0].index];
      openMonthDrilldown(`FuelEU Penalty in ${bucket.monthLabel}`, "Voyage rows contributing to the selected monthly FuelEU penalty.", bucket);
    }),
  });

  createChart("surplusValue", "surplusChart", {
    type: "bar",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "Surplus value",
          data: voyageRows.map((row) => surplusValueUsd(row)),
          backgroundColor: "#7f77dd",
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const row = voyageRows[elementsClicked[0].index];
      openDrilldown(
        `Surplus Value for ${row.recordId}`,
        row.route,
        ["Record", "Vessel", "Compliance Balance", "Surplus Value (USD)", "Penalty"],
        [[row.recordId, row.vesselName, formatNumber(row.complianceBalanceT, 2), formatUsdCurrency(surplusValueUsd(row)), formatCurrency(row.fuelEuPenaltyEur)]],
        [row]
      );
    }),
  });

  createChart("totalCost", "totalCostChart", {
    type: "bar",
    data: {
      labels: voyageRows.map((row) => row.recordId),
      datasets: [
        {
          label: "Total cost",
          data: voyageRows.map((row) => totalOperationalCostEur(row)),
          backgroundColor: "#145c9e",
          borderRadius: 6,
        },
      ],
    },
    options: baseOptions((_, elementsClicked) => {
      if (!elementsClicked.length) return;
      const row = voyageRows[elementsClicked[0].index];
      openDrilldown(
        `Total Cost for ${row.recordId}`,
        row.route,
        ["Record", "Vessel", "EUA Cost", "FuelEU Penalty", "Total Cost"],
        [[row.recordId, row.vesselName, formatCurrency(row.euasCostEur), formatCurrency(row.fuelEuPenaltyEur), formatCurrency(totalOperationalCostEur(row))]],
        [row]
      );
    }),
  });
}

function ensureCalculatorSelection() {
  const selectedStateRow = getCalculatorStateRow(stateStore.ui.calculatorSelectedId);
  if (selectedStateRow) {
    return stateStore.derived.calculatorRows.find((row) => row.id === selectedStateRow.id) || null;
  }

  const draftStateRow = stateStore.state.calculatorRows.find((row) => !rowHasMeaningfulInputs(row));
  if (draftStateRow) {
    stateStore.ui.calculatorSelectedId = draftStateRow.id;
    return stateStore.derived.calculatorRows.find((row) => row.id === draftStateRow.id) || null;
  }
  const fallback = getMeaningfulDerivedRows()[0] || stateStore.derived.calculatorRows[0] || null;
  stateStore.ui.calculatorSelectedId = fallback?.id || null;
  return fallback;
}

function nextRecordSerial(type) {
  const prefix = type === "Port Stay" ? "P" : "V";
  const values = stateStore.state.calculatorRows
    .filter((row) => rowHasMeaningfulInputs(row))
    .map((row) => String(row.recordId || "").trim().toUpperCase())
    .filter((value) => value.startsWith(prefix))
    .map((value) => Number(value.slice(1)))
    .filter((value) => Number.isFinite(value));
  const next = (values.length ? Math.max(...values) : 0) + 1;
  return `${prefix}${String(next).padStart(3, "0")}`;
}

function buildCalculatorRowForCurrentFilter(sourceRow = null) {
  const row = blankCalculatorRow();
  let targetYear = getCurrentReportYear();
  while (countRowsForYear(targetYear) >= MAX_CALCULATOR_ROWS_PER_YEAR) {
    targetYear += 1;
  }
  row.storageYear = targetYear;
  row.type = "Voyage";
  row.recordId = nextRecordSerial(row.type);
  if (sourceRow?.imoNo) {
    row.imoNo = sourceRow.imoNo;
  }

  if (sourceRow?.windFactor !== undefined && sourceRow?.windFactor !== null) {
    row.windFactor = sourceRow.windFactor;
  }

  if (sourceRow?.sourceSystem) {
    row.sourceSystem = sourceRow.sourceSystem;
    row.entrySource = sourceRow.entrySource || "manual";
  }

  if (sourceRow?.type === "Voyage" || sourceRow?.type === "Port Stay") {
    row.type = sourceRow.type;
    row.recordId = nextRecordSerial(row.type);
  }

  if (sourceRow?.storageYear) {
    row.storageYear = sourceRow.storageYear;
  }

  if (sourceRow) {
    return row;
  }

  if (stateStore.ui.vesselFilter === "all") {
    return row;
  }

  const vesselRows = stateStore.derived.calculatorRows.filter((item) => item.vesselName === stateStore.ui.vesselFilter);
  const lastVesselRow = vesselRows[vesselRows.length - 1];
  const filterSourceRow = stateStore.state.calculatorRows.find((item) => item.id === lastVesselRow?.id);

  if (filterSourceRow?.imoNo) {
    row.imoNo = filterSourceRow.imoNo;
  }

  if (filterSourceRow?.windFactor !== undefined && filterSourceRow?.windFactor !== null) {
    row.windFactor = filterSourceRow.windFactor;
  }

  if (filterSourceRow?.type === "Voyage" || filterSourceRow?.type === "Port Stay") {
    row.type = filterSourceRow.type;
    row.recordId = nextRecordSerial(row.type);
  }

  if (filterSourceRow?.storageYear) {
    row.storageYear = filterSourceRow.storageYear;
  }
  return row;
}

function insertCalculatorRow(row, insertAfterRowId = null) {
  if (insertAfterRowId) {
    const sourceIndex = stateStore.state.calculatorRows.findIndex((item) => item.id === insertAfterRowId);
    if (sourceIndex >= 0) {
      stateStore.state.calculatorRows.splice(sourceIndex + 1, 0, row);
      return;
    }
  }
  if (stateStore.ui.vesselFilter === "all") {
    stateStore.state.calculatorRows.unshift(row);
    return;
  }

  const vesselRows = stateStore.derived.calculatorRows.filter((item) => item.vesselName === stateStore.ui.vesselFilter);
  const lastVesselRow = vesselRows[vesselRows.length - 1];
  const insertIndex = lastVesselRow ? stateStore.state.calculatorRows.findIndex((item) => item.id === lastVesselRow.id) : -1;

  if (insertIndex === -1) {
    stateStore.state.calculatorRows.unshift(row);
    return;
  }

  stateStore.state.calculatorRows.splice(insertIndex + 1, 0, row);
}

function ensureDraftRow(sourceRow = null, insertAfterRowId = null) {
  let draftRow = stateStore.state.calculatorRows.find((row) => !rowHasMeaningfulInputs(row));
  if (!draftRow) {
    draftRow = buildCalculatorRowForCurrentFilter(sourceRow);
    insertCalculatorRow(draftRow, insertAfterRowId);
  }
  stateStore.ui.calculatorSelectedId = draftRow.id;
  pruneCalculatorDraftRows([draftRow.id]);
  return draftRow;
}

function getFilteredCalculatorRows() {
  const searchTerm = lower(stateStore.ui.calculatorSearch);
  return stateStore.derived.calculatorRows.filter((row) => {
    const stateRow = getCalculatorStateRow(row.id);
    const isMeaningful = rowHasMeaningfulInputs(stateRow);
    if (!isMeaningful && row.id !== stateStore.ui.calculatorSelectedId) {
      return false;
    }
    if (stateStore.ui.vesselFilter !== "all" && row.vesselName !== stateStore.ui.vesselFilter) {
      return false;
    }
    if (!searchTerm) return true;
    return [
      row.recordId,
      row.vesselName,
      row.route,
      row.shipType,
      row.flagState,
      row.fromPortName,
      row.toPortName,
      row.fromPortCode,
      row.toPortCode,
      row.fuel1Type,
      row.fuel2Type,
      row.bioFuelType,
      row.type,
      row.imoNo,
    ]
      .filter(Boolean)
      .some((value) => lower(value).includes(searchTerm));
  });
}

function renderCalculator() {
  ensureCalculatorSelection();
  const filteredRows = getFilteredCalculatorRows();
  const focusedRow = filteredRows.find((row) => row.id === stateStore.ui.calculatorSelectedId) || filteredRows[0] || null;
  const historyRows = focusedRow ? filteredRows.filter((row) => row.id !== focusedRow.id) : [];
  if (focusedRow && stateStore.ui.calculatorSelectedId !== focusedRow.id) {
    stateStore.ui.calculatorSelectedId = focusedRow.id;
  }

  const visibleColumns = getVisibleCalculatorColumns();
  const focusedInputRow = focusedRow ? stateStore.state.calculatorRows.find((item) => item.id === focusedRow.id) || blankCalculatorRow() : blankCalculatorRow();
  const inputRowsById = new Map(filteredRows.map((row) => [row.id, getCalculatorStateRow(row.id) || blankCalculatorRow()]));
  const computedWidths = new Map(
    visibleColumns.map((column) => [
      column.key,
      filteredRows.length ? estimateCalculatorColumnWidth(column, filteredRows, inputRowsById) : (column.width || 120),
    ])
  );
  const stickyColumns = visibleColumns.filter((column) => column.kind.startsWith("sticky"));
  let stickyOffset = 0;
  const stickyOffsets = new Map();
  stickyColumns.forEach((column) => {
    stickyOffsets.set(column.key, stickyOffset);
    stickyOffset += computedWidths.get(column.key) || column.width || 120;
  });
  const totalTableWidth = visibleColumns.reduce((sum, column) => sum + (computedWidths.get(column.key) || column.width || 120), 0);

  const totalRows = filteredRows.length;
  const activeRows = filteredRows.filter((row) => rowHasMeaningfulInputs(getCalculatorStateRow(row.id)));

  return `
    <section class="calculator-shell">
      <div class="calculator-toolbar">
        <div class="calculator-toolbar-copy">
          <h2>Voyage Inputs</h2>
          <p class="helper-text">Orange cells are editable inputs, including voyage ID and type. Blue cells are workbook-driven outputs. Only the current row stays open for editing so the grid feels much faster.</p>
        </div>
        <div class="calculator-actions">
          <input
            class="search-input"
            type="search"
            data-action="calculator-search"
            value="${stateStore.ui.calculatorSearch}"
            placeholder="Search vessel, route, port, IMO, or fuel"
          >
          <div class="calculator-column-picker">
            <button class="inline-button" type="button" data-action="toggle-calculator-columns">Columns</button>
            ${
              stateStore.ui.calculatorColumnMenuOpen
                ? `
                  <div class="calculator-column-menu">
                    <div class="calculator-column-menu-head">
                      <strong>Choose visible columns</strong>
                      <button class="inline-button compact-button" type="button" data-action="reset-calculator-columns">Show all</button>
                    </div>
                    <div class="calculator-column-list">
                      ${CALCULATOR_COLUMNS.filter((column) => column.key !== "rowActions")
                        .map((column) => {
                          const checked = getVisibleCalculatorColumnKeys().includes(column.key) ? "checked" : "";
                          return `
                            <label class="calculator-column-option">
                              <input type="checkbox" data-action="toggle-calculator-column" data-column-key="${column.key}" ${checked}>
                              <span>${column.label}</span>
                            </label>
                          `;
                        })
                        .join("")}
                    </div>
                  </div>
                `
                : ""
            }
          </div>
          <button class="inline-button button-amber" type="button" data-action="add-calculator-row">Add row</button>
        </div>
      </div>

      ${renderCollapsibleSection({
        action: "toggle-calculator",
        title: `Voyage Input Sheet${stateStore.ui.vesselFilter !== "all" ? ` for ${stateStore.ui.vesselFilter}` : ""}`,
        badges: `${renderSectionBadge(`${totalRows} rows`)}${renderSectionBadge(`${activeRows.length} active`, "purple")}`,
        note: "Collapse when you need a cleaner view",
        open: stateStore.ui.calculatorOpen,
        body: `
          <article class="table-card compact-table-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Workbook Grid</p>
            <h3>Active editable row${stateStore.ui.vesselFilter !== "all" ? ` for ${stateStore.ui.vesselFilter}` : ""}</h3>
            <p class="helper-text">Only one row stays open for editing at a time. Earlier rows remain available in the history panel below.</p>
          </div>
          <span class="chip">${totalRows} rows / ${activeRows.length} active records</span>
        </div>
        <div class="calculator-table-wrap" data-scroll-group="calculator-active">
          <table class="calculator-table" style="width:${totalTableWidth}px;min-width:${totalTableWidth}px">
            <thead>
              <tr>
                ${visibleColumns.map((column) => {
                  const classes = ["calculator-header"];
                  const columnWidth = computedWidths.get(column.key) || column.width || 120;
                  const styleParts = [`min-width:${columnWidth}px`, `width:${columnWidth}px`];
                  if (column.kind.includes("sticky")) {
                    classes.push("sticky-header");
                    styleParts.push(`left:${stickyOffsets.get(column.key) || 0}px`);
                  } else if (column.kind.includes("editable")) {
                    classes.push("editable-header");
                  } else if (column.kind.includes("calculated")) {
                    classes.push("calculated-header");
                  } else if (column.kind === "actions") {
                    classes.push("actions-header");
                  }
                  return `<th class="${classes.join(" ")}" style="${styleParts.join(";")}">${column.label}</th>`;
                }).join("")}
              </tr>
            </thead>
            <tbody>
              ${
                    focusedRow
                      ? (() => {
                          return `
                            <tr class="selected-row" data-action="select-calculator-row" data-row-id="${focusedRow.id}">
                          ${visibleColumns.map((column) => renderCalculatorCell(focusedRow, focusedInputRow, { ...column, width: computedWidths.get(column.key) || column.width || 120 }, stickyOffsets.get(column.key) || 0)).join("")}
                            </tr>
                          `;
                        })()
                  : `<tr><td colspan="${visibleColumns.length}" class="no-data">No voyage rows match the current vessel filter and search.</td></tr>`
              }
            </tbody>
          </table>
        </div>
        ${renderCalculatorHistory(historyRows, visibleColumns, computedWidths, stickyOffsets, totalTableWidth)}
          </article>
        `,
      })}
    </section>
  `;
}

function groupCalculatorRowsByYear(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const year = resolveRowStorageYear(getCalculatorStateRow(row.id) || row);
    if (!grouped.has(year)) {
      grouped.set(year, []);
    }
    grouped.get(year).push(row);
  });
  return [...grouped.entries()].sort((a, b) => Number(b[0]) - Number(a[0]));
}

function renderCalculatorHistory(historyRows, visibleColumns, computedWidths, stickyOffsets, totalTableWidth) {
  if (!historyRows.length) {
    return "";
  }

  const pageCount = Math.max(1, Math.ceil(historyRows.length / CALCULATOR_HISTORY_PAGE_SIZE));
  stateStore.ui.calculatorHistoryPage = Math.min(Math.max(1, stateStore.ui.calculatorHistoryPage), pageCount);
  const pageStart = (stateStore.ui.calculatorHistoryPage - 1) * CALCULATOR_HISTORY_PAGE_SIZE;
  const pageRows = historyRows.slice(pageStart, pageStart + CALCULATOR_HISTORY_PAGE_SIZE);
  const yearGroups = stateStore.ui.calculatorHistoryOpen ? groupCalculatorRowsByYear(pageRows) : [];
  return `
    <div class="calculator-history">
      <button class="section-bar calculator-history-toggle" type="button" data-action="toggle-calculator-history">
        <span class="section-bar-main">
          <span class="section-chevron ${stateStore.ui.calculatorHistoryOpen ? "open" : ""}">&#9662;</span>
          <span class="section-title-group">
            <strong>Previous rows</strong>
            ${renderSectionBadge(`${historyRows.length} stored`)}
          </span>
        </span>
        <span class="section-bar-note">Up to ${CALCULATOR_HISTORY_PAGE_SIZE} rows per page. Editing keeps each row in its current position.</span>
      </button>
      ${
        stateStore.ui.calculatorHistoryOpen
          ? `
            <div class="section-body">
              <div class="calculator-history-groups">
                ${yearGroups
                  .map(
                    ([year, rows]) => `
                      <section class="calculator-year-group">
                        <div class="calculator-year-head">
                          <span class="eyebrow">Reporting Year</span>
                          <strong>${year}</strong>
                          <span class="chip">${rows.length} on this page / ${MAX_CALCULATOR_ROWS_PER_YEAR} max per year</span>
                        </div>
                        <div class="table-wrap calculator-history-wrap" data-scroll-group="calculator-history">
                          <table class="calculator-history-table" style="width:${totalTableWidth}px;min-width:${totalTableWidth}px">
                            <thead>
                              <tr>
                                ${visibleColumns.map((column) => {
                                  const columnWidth = computedWidths.get(column.key) || column.width || 120;
                                  const classes = ["calculator-header"];
                                  const styleParts = [`min-width:${columnWidth}px`, `width:${columnWidth}px`];
                                  if (column.kind.includes("sticky")) {
                                    classes.push("sticky-header");
                                    styleParts.push(`left:${stickyOffsets.get(column.key) || 0}px`);
                                  } else if (column.kind.includes("editable")) {
                                    classes.push("editable-header");
                                  } else if (column.kind.includes("calculated")) {
                                    classes.push("calculated-header");
                                  } else if (column.kind === "actions") {
                                    classes.push("actions-header");
                                  }
                                  return `<th class="${classes.join(" ")}" style="${styleParts.join(";")}">${column.label}</th>`;
                                }).join("")}
                              </tr>
                            </thead>
                            <tbody>
                              ${rows
                                .map((row) => {
                                  const sourceRow = getCalculatorStateRow(row.id);
                                  return `
                                    <tr>
                                      ${visibleColumns
                                        .map((column) =>
                                          renderCalculatorCell(
                                            row,
                                            sourceRow || blankCalculatorRow(),
                                            { ...column, width: computedWidths.get(column.key) || column.width || 120 },
                                            stickyOffsets.get(column.key) || 0,
                                            "history"
                                          )
                                        )
                                        .join("")}
                                    </tr>
                                  `;
                                })
                                .join("")}
                            </tbody>
                          </table>
                        </div>
                      </section>
                    `
                  )
                  .join("")}
              </div>
              <div class="calculator-history-pagination">
                <button class="inline-button compact-button" type="button" data-action="calculator-history-prev" ${stateStore.ui.calculatorHistoryPage === 1 ? "disabled" : ""}>Previous</button>
                <span>Page ${stateStore.ui.calculatorHistoryPage} of ${pageCount}</span>
                <button class="inline-button compact-button" type="button" data-action="calculator-history-next" ${stateStore.ui.calculatorHistoryPage === pageCount ? "disabled" : ""}>Next</button>
              </div>
            </div>
          `
          : ""
      }
    </div>
  `;
}

function getSheetRowsForDisplay(sheetKey) {
  if (sheetKey === "fuelReference") {
    return stateStore.derived.fuelReference;
  }
  return getCollection(sheetKey) || [];
}

function renderLibraryTabs() {
  elements.libraryTabs.innerHTML = REFERENCE_SHEETS.map(
    (sheet) => `
      <button
        class="sheet-tab ${stateStore.ui.librarySheet === sheet.key ? "active" : ""}"
        type="button"
        data-action="select-library-sheet"
        data-sheet="${sheet.key}"
      >
        ${sheet.label}
      </button>
    `
  ).join("");
}

function getLibraryDisplayColumns(sheetKey) {
  return LIBRARY_DISPLAY_COLUMNS[sheetKey] || SHEET_COLUMNS[sheetKey] || [];
}

function getLibraryColumnLabel(sheetKey, column) {
  return LIBRARY_COLUMN_LABELS[sheetKey]?.[column] || column;
}

function formatLibraryValue(value) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  if (typeof value === "number") {
    if (Number.isInteger(value)) {
      return formatInteger(value);
    }
    return formatNumber(value, Math.abs(value) < 10 ? 4 : 2);
  }
  return value;
}

function renderLibraryContent() {
  const sheetKey = stateStore.ui.librarySheet;
  const allRows = getSheetRowsForDisplay(sheetKey);
  const search = lower(stateStore.ui.librarySearch);
  const filteredRows = allRows.filter((row) => {
    if (!search) return true;
    return Object.values(row).some((value) => lower(value).includes(search));
  });
  const pageCount = Math.max(1, Math.ceil(filteredRows.length / LIBRARY_PAGE_SIZE));
  if (stateStore.ui.libraryPage > pageCount) {
    stateStore.ui.libraryPage = pageCount;
  }
  const pageRows = filteredRows.slice(
    (stateStore.ui.libraryPage - 1) * LIBRARY_PAGE_SIZE,
    stateStore.ui.libraryPage * LIBRARY_PAGE_SIZE
  );
  const visibleColumns = getLibraryDisplayColumns(sheetKey);

  elements.libraryContent.innerHTML = `
    <div class="library-toolbar">
      <div>
        <h3>${REFERENCE_SHEETS.find((sheet) => sheet.key === sheetKey)?.label || "Library"}</h3>
        <p class="helper-text">Hidden from the main navigation, but still editable here whenever you need to adjust the calculation library.</p>
      </div>
      <div class="library-toolbar-actions">
        <input
          class="search-input"
          type="search"
          value="${stateStore.ui.librarySearch}"
          data-action="library-search"
          placeholder="Search this library sheet"
        >
        <button class="inline-button" type="button" data-action="open-row-editor" data-sheet="${sheetKey}">Add row</button>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            ${visibleColumns.map((column) => `<th>${getLibraryColumnLabel(sheetKey, column)}</th>`).join("")}
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${pageRows
            .map(
              (row) => `
                <tr>
                  ${visibleColumns.map((column) => `<td>${formatLibraryValue(row[column])}</td>`).join("")}
                  <td>
                    <div class="toolbar-actions">
                      <button class="inline-button" type="button" data-action="open-row-editor" data-sheet="${sheetKey}" data-row-id="${row.id}">Edit</button>
                      <button class="danger-button" type="button" data-action="delete-sheet-row" data-sheet="${sheetKey}" data-row-id="${row.id}">Delete</button>
                    </div>
                  </td>
                </tr>
              `
            )
            .join("") || `<tr><td colspan="${visibleColumns.length + 1}">No rows match the current search.</td></tr>`}
        </tbody>
      </table>
    </div>

    <div class="pagination-row">
      <button class="inline-button" type="button" data-action="library-page-prev">Previous</button>
      <span class="page-label">Page ${stateStore.ui.libraryPage} of ${pageCount}</span>
      <button class="inline-button" type="button" data-action="library-page-next">Next</button>
    </div>
  `;
}

function renderContent() {
  destroyCharts();
  if (stateStore.ui.activeView === "dashboard") {
    elements.contentView.innerHTML = renderDashboard();
    renderDashboardCharts();
    return;
  }

  elements.contentView.innerHTML = renderCalculator();
}

function renderLibraryDrawer() {
  elements.libraryDrawer.classList.toggle("open", stateStore.ui.libraryOpen);
  elements.libraryBackdrop.classList.toggle("open", stateStore.ui.libraryOpen);
  document.body.classList.toggle("library-open", stateStore.ui.libraryOpen);
  renderLibraryTabs();
  renderLibraryContent();
}

function render() {
  renderViewTabs();
  renderVesselFilter();
  renderSyncStatus();
  if (stateStore.ui.activeView === "dashboard") {
    elements.kpiGrid.classList.remove("hidden");
    renderKpis();
  } else {
    elements.kpiGrid.classList.add("hidden");
    elements.kpiGrid.innerHTML = "";
  }
  buildDataLists();
  renderContent();
  wireContentScrollRegions();
  renderLibraryDrawer();
}

async function loadEuaMarketSnapshot() {
  stateStore.market.status = "loading";

  try {
    const response = await fetch("/api/market/eua");
    if (!response.ok) {
      throw new Error(`Market API returned ${response.status}`);
    }

    const payload = await response.json();
    stateStore.market.status = "ready";
    stateStore.market.snapshot = payload;
    stateStore.market.error = null;
  } catch (error) {
    stateStore.market.status = "error";
    stateStore.market.error = error.message;
  }

  render();
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"', '""')}"`;
}

function showToast(message, tone = "success") {
  document.querySelector(".app-toast")?.remove();
  const toast = document.createElement("div");
  toast.className = `app-toast ${tone}`;
  toast.setAttribute("role", "status");
  toast.textContent = message;
  document.body.appendChild(toast);
  window.requestAnimationFrame(() => toast.classList.add("visible"));
  window.setTimeout(() => {
    toast.classList.remove("visible");
    window.setTimeout(() => toast.remove(), 180);
  }, 3200);
}

function downloadText(filename, content, type = "text/plain") {
  const blob = new Blob([content], { type });
  downloadBlob(filename, blob);
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.rel = "noopener";
  link.style.display = "none";
  document.body.appendChild(link);
  window.requestAnimationFrame(() => {
    link.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
    showToast(`Export started: ${filename}`, "success");
    window.setTimeout(() => {
      URL.revokeObjectURL(url);
      link.remove();
    }, 4000);
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getReportRowsForImo(imoNo) {
  return stateStore.derived.calculatorRows.filter((row) => {
    const sourceRow = getCalculatorStateRow(row.id);
    return String(row.imoNo || "") === String(imoNo || "") && rowHasMeaningfulInputs(sourceRow);
  });
}

function getReportBasis() {
  return elements.reportDialog.querySelector('input[name="reportBasis"]:checked')?.value || "vessel";
}

function updateReportSelectionSummary() {
  const basis = getReportBasis();
  const checkboxes = [...elements.reportRowList.querySelectorAll('input[type="checkbox"]')];
  const selected = basis === "vessel" ? checkboxes.length : checkboxes.filter((item) => item.checked).length;
  checkboxes.forEach((item) => {
    item.disabled = basis === "vessel";
    if (basis === "vessel") item.checked = true;
  });
  elements.reportSelectionSummary.textContent = `${selected} of ${checkboxes.length} records will be included.`;
}

function renderReportRows() {
  const imoNo = elements.reportVesselSelect.value;
  const rows = getReportRowsForImo(imoNo);
  elements.reportRowList.innerHTML =
    rows
      .map(
        (row) => `
          <label class="report-row-option">
            <input type="checkbox" value="${escapeHtml(row.id)}" checked>
            <span class="report-row-id">${escapeHtml(row.recordId)}</span>
            <span>${escapeHtml(row.type)}</span>
            <span>${escapeHtml(formatDateValue(row.departureDate))}</span>
            <span class="report-row-route">${escapeHtml(row.route)}</span>
            <span>${formatNumber(row.euasRequiredT, 2)} EUAs</span>
          </label>
        `
      )
      .join("") || `<div class="empty-state compact-empty-state">No voyage or port-stay records are available for this vessel.</div>`;
  updateReportSelectionSummary();
}

function openReportDialog() {
  const vesselRows = stateStore.state.fleet
    .map((vessel) => ({ vessel, count: getReportRowsForImo(vessel.imoNo).length }))
    .filter((item) => item.count > 0);
  if (!vesselRows.length) {
    showToast("Add at least one voyage input before generating a report.", "info");
    return;
  }

  elements.reportVesselSelect.innerHTML = vesselRows
    .map(
      ({ vessel, count }) =>
        `<option value="${escapeHtml(vessel.imoNo)}">${escapeHtml(vessel.vesselName)} - IMO ${escapeHtml(vessel.imoNo)} (${count})</option>`
    )
    .join("");
  const filteredVessel = vesselRows.find((item) => item.vessel.vesselName === stateStore.ui.vesselFilter);
  if (filteredVessel) {
    elements.reportVesselSelect.value = String(filteredVessel.vessel.imoNo);
  }
  elements.reportDialog.querySelector('input[name="reportBasis"][value="vessel"]').checked = true;
  renderReportRows();
  elements.reportDialog.showModal();
}

async function generateComplianceReport() {
  const imoNo = elements.reportVesselSelect.value;
  const vessel = stateStore.state.fleet.find((item) => String(item.imoNo) === String(imoNo));
  if (!vessel) {
    showToast("Select a vessel before generating the report.", "info");
    return;
  }

  const allRows = getReportRowsForImo(imoNo);
  const basis = getReportBasis();
  const selectedIds = new Set(
    [...elements.reportRowList.querySelectorAll('input[type="checkbox"]:checked')].map((item) => item.value)
  );
  const rows = basis === "vessel" ? allRows : allRows.filter((row) => selectedIds.has(row.id));
  if (!rows.length) {
    showToast("Select at least one voyage or port stay.", "info");
    return;
  }

  elements.generateReportConfirmButton.disabled = true;
  elements.generateReportConfirmButton.textContent = "Generating...";
  try {
    const response = await fetch("/api/reports/compliance-statement", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        vessel,
        rows,
        reportYear: stateStore.derived.parameterValues.reportYear,
        euaPrice: stateStore.derived.parameterValues.euaPrice,
        issuedAt: new Date().toISOString(),
        selectionLabel:
          basis === "vessel"
            ? `All ${rows.length} records for ${vessel.vesselName}`
            : `${rows.length} selected records: ${rows.map((row) => row.recordId).join(", ")}`,
      }),
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.message || `Report service responded with ${response.status}`);
    }
    const disposition = response.headers.get("Content-Disposition") || "";
    const filename = disposition.match(/filename="([^"]+)"/i)?.[1] || `Statement_${vessel.vesselName}_${stateStore.derived.parameterValues.reportYear}.pdf`;
    downloadBlob(filename, await response.blob());
    elements.reportDialog.close();
    showToast(`Report generated: ${filename}`, "success");
  } catch (error) {
    showToast(`Report could not be generated: ${error.message}`, "info");
  } finally {
    elements.generateReportConfirmButton.disabled = false;
    elements.generateReportConfirmButton.textContent = "Generate PDF";
  }
}

function wireScrollGroup(selector, stateKey, verticalStateKey = null) {
  const regions = [...elements.contentView.querySelectorAll(selector)];
  if (!regions.length) {
    return;
  }

  const savedScrollLeft = Number(stateStore.ui[stateKey] || 0);
  const savedScrollTop = verticalStateKey ? Number(stateStore.ui[verticalStateKey] || 0) : 0;
  regions.forEach((region) => {
    region.scrollLeft = savedScrollLeft;
    if (verticalStateKey) {
      region.scrollTop = savedScrollTop;
    }
  });

  let syncing = false;
  regions.forEach((region) => {
    region.addEventListener(
      "scroll",
      () => {
        if (syncing) {
          return;
        }
        syncing = true;
        const nextScrollLeft = region.scrollLeft;
        stateStore.ui[stateKey] = nextScrollLeft;
        if (verticalStateKey) {
          stateStore.ui[verticalStateKey] = region.scrollTop;
        }
        regions.forEach((otherRegion) => {
          if (otherRegion !== region && Math.abs(otherRegion.scrollLeft - nextScrollLeft) > 1) {
            otherRegion.scrollLeft = nextScrollLeft;
          }
        });
        syncing = false;
      },
      { passive: true }
    );
  });
}

function wireContentScrollRegions() {
  wireScrollGroup('[data-scroll-group="detail"]', "detailScrollLeft");
  wireScrollGroup('[data-scroll-group="calculator-active"]', "calculatorActiveScrollLeft");
  wireScrollGroup(
    '[data-scroll-group="calculator-history"]',
    "calculatorHistoryScrollLeft",
    "calculatorHistoryScrollTop"
  );
}

function exportValueForColumn(row, column) {
  const value = row[column.key];
  if (column.format === "date") return formatDateValue(value);
  if (column.format === "currency") return value === null || value === undefined || value === "" ? "" : Number(value);
  if (column.format === "number" || column.format === "ghg" || column.format === "balance") {
    return value === null || value === undefined || value === "" ? "" : Number(value);
  }
  if (column.format === "scope") return numberOrZero(value);
  return value ?? "";
}

function exportFilteredData() {
  const slug = stateStore.ui.vesselFilter === "all" ? "all-vessels" : stateStore.ui.vesselFilter.toLowerCase().replaceAll(/\s+/g, "-");
  if (stateStore.ui.activeView === "dashboard") {
    const visibleColumns = projectionIsActive() ? DETAIL_TABLE_COLUMNS : DETAIL_TABLE_COLUMNS.filter((column) => !column.proj);
    const rows = getDetailRows();
    const csv = [
      visibleColumns.map((column) => csvEscape(column.label)).join(","),
      ...rows.map((row) => visibleColumns.map((column) => csvEscape(exportValueForColumn(row, column))).join(",")),
    ].join("\n");
    downloadText(`fuel-ets-${slug}-dashboard-detail.csv`, csv, "text/csv");
    return;
  }

  const visibleColumns = getVisibleCalculatorColumns().filter((column) => column.key !== "rowActions");
  const rows = getFilteredCalculatorRows().filter((row) => rowHasMeaningfulInputs(getCalculatorStateRow(row.id)));
  const csv = [
    visibleColumns.map((column) => csvEscape(column.label)).join(","),
    ...rows.map((row) =>
      visibleColumns
        .map((column) => {
          const inputRow = getCalculatorStateRow(row.id) || blankCalculatorRow();
          const value = column.kind.includes("editable") ? calculatorInputValue(inputRow, row, column) : calculatorCellValue(row, column);
          return csvEscape(value);
        })
        .join(",")
    ),
  ].join("\n");
  downloadText(`fuel-ets-${slug}-voyage-inputs.csv`, csv, "text/csv");
}

function mapExternalRowToCalculatorRow(sourceSystem, payload = {}) {
  const row = blankCalculatorRow();
  row.entrySource = "api";
  row.sourceSystem = sourceSystem || "External";
  row.sourceRecordId = String(payload.sourceRecordId || payload.externalId || payload.id || "");
  row.sourceUpdatedAt = String(payload.sourceUpdatedAt || payload.updatedAt || new Date().toISOString());
  row.storageYear = Number(payload.storageYear) || extractYearFromDate(payload.departureDate) || extractYearFromDate(payload.arrivalDate) || getCurrentReportYear();

  const directFields = [
    "recordId",
    "type",
    "imoNo",
    "departureDate",
    "fromPortCode",
    "arrivalDate",
    "toPortCode",
    "fuel1Type",
    "fuel1ConsumptionMt",
    "fuel2Type",
    "fuel2ConsumptionMt",
    "bioFuelType",
    "bioFuelConsumptionMt",
    "sustainabilityFactor",
    "windFactor",
    "distanceNm",
    "cargoTonnes",
    "timeAtSeaHours",
    "berthHours",
    "opsElectricityMj",
  ];

  directFields.forEach((field) => {
    if (payload[field] !== undefined) {
      setCalculatorRowValue(row, field, payload[field]);
    }
  });

  if (!row.recordId) {
    row.recordId = nextRecordSerial(row.type || "Voyage");
  }
  return row;
}

function importVoyageRowsFromExternal(sourceSystem, payloadRows = [], options = {}) {
  const rows = Array.isArray(payloadRows) ? payloadRows : [];
  const replaceMatchingSource = Boolean(options.replaceMatchingSource);
  const incoming = rows.map((payload) => mapExternalRowToCalculatorRow(sourceSystem, payload));

  incoming.forEach((row) => {
    const existingIndex = stateStore.state.calculatorRows.findIndex((item) => {
      if (replaceMatchingSource && row.sourceRecordId) {
        return item.sourceSystem === row.sourceSystem && item.sourceRecordId === row.sourceRecordId;
      }
      return false;
    });
    if (existingIndex >= 0) {
      stateStore.state.calculatorRows[existingIndex] = row;
    } else {
      insertCalculatorRow(row);
    }
  });

  stateStore.ui.calculatorSelectedId = incoming[0]?.id || stateStore.ui.calculatorSelectedId;
  recomputeAndRender();
}

function updateCalculatorField(field, rawValue) {
  const row = stateStore.state.calculatorRows.find((item) => item.id === stateStore.ui.calculatorSelectedId);
  if (!row) return;
  setCalculatorRowValue(row, field, rawValue);
  recomputeAndRender();
}

function setCalculatorRowValue(row, field, rawValue) {
  const numericKeys = new Set([
    "imoNo",
    "fuel1ConsumptionMt",
    "fuel2ConsumptionMt",
    "bioFuelConsumptionMt",
    "sustainabilityFactor",
    "windFactor",
    "distanceNm",
    "cargoTonnes",
    "timeAtSeaHours",
    "berthHours",
    "opsElectricityMj",
  ]);
  row[field] = numericKeys.has(field) ? (rawValue === "" ? null : Number(rawValue)) : rawValue;
  row.entrySource = row.entrySource || "manual";

  if (field === "recordId") {
    row.recordId = String(row.recordId || "").toUpperCase();
  }

  if (field === "departureDate" || field === "arrivalDate") {
    row.storageYear = extractYearFromDate(field === "departureDate" ? row.departureDate : row.arrivalDate) || row.storageYear || getCurrentReportYear();
  }

  if (field === "imoNo" && (row.windFactor === null || row.windFactor === undefined || row.windFactor === "")) {
    const vessel = stateStore.state.fleet.find((item) => String(item.imoNo) === String(Math.trunc(Number(row.imoNo) || 0)));
    if (vessel?.wapsFwindFactor !== null && vessel?.wapsFwindFactor !== undefined && vessel?.wapsFwindFactor !== "") {
      row.windFactor = Number(vessel.wapsFwindFactor);
    }
  }

  if (field === "type" && (row.type === "Voyage" || row.type === "Port Stay")) {
    const prefix = row.type === "Port Stay" ? "P" : "V";
    const currentRecordId = String(row.recordId || "").toUpperCase();
    if (!currentRecordId || !currentRecordId.startsWith(prefix)) {
      row.recordId = nextRecordSerial(row.type);
    }
  }
}

function updateCalculatorCell(rowId, field, rawValue, commit = true, context = "active") {
  const row = stateStore.state.calculatorRows.find((item) => item.id === rowId);
  if (!row) return;
  setCalculatorRowValue(row, field, rawValue);
  if (context === "active") {
    stateStore.ui.calculatorSelectedId = rowId;
  }
  if (commit) {
    recomputeAndRender();
  }
}

function openEditorDialog(sheetKey, rowId) {
  const rows = getCollection(sheetKey);
  const row = rowId ? rows.find((item) => item.id === rowId) : blankRowForSheet(sheetKey);
  stateStore.ui.dialog = {
    sheetKey,
    rowId: rowId || null,
    draft: deepClone(row),
  };
  renderEditorDialog();
  elements.rowEditorDialog.showModal();
}

function renderEditorDialog() {
  const dialogState = stateStore.ui.dialog;
  if (!dialogState) return;
  const { sheetKey, draft } = dialogState;
  const columns = SHEET_COLUMNS[sheetKey];
  elements.editorDialogTitle.textContent = `${dialogState.rowId ? "Edit" : "Add"} ${sheetKey} row`;
  elements.editorDialogBody.innerHTML = columns
    .map((column) => {
      const value = draft[column] ?? "";
      const multiline = column.toLowerCase().includes("note") || column.toLowerCase().includes("detail") || column.toLowerCase().includes("formula");
      const selectOptions = EDITOR_SELECT_OPTIONS[sheetKey]?.[column];
      const inputType =
        numericColumns[sheetKey]?.has(column) || (sheetKey === "parameters" && column === "value" && draft.type === "number")
          ? "number"
          : "text";
      return `
        <div class="editor-field">
          <label for="editor-${column}">${column}</label>
          ${
            selectOptions
              ? `
                <select class="editor-input" id="editor-${column}" data-editor-field="${column}">
                  ${selectOptions
                    .map(
                      (option) =>
                        `<option value="${escapeHtml(option.value)}" ${String(value) === option.value ? "selected" : ""}>${escapeHtml(option.label)}</option>`
                    )
                    .join("")}
                </select>
              `
              : multiline
              ? `<textarea class="editor-textarea" id="editor-${column}" data-editor-field="${column}">${value}</textarea>`
              : `<input class="editor-input" id="editor-${column}" type="${inputType}" data-editor-field="${column}" value="${value}">`
          }
        </div>
      `;
    })
    .join("");
}

function saveEditorDialog() {
  const dialogState = stateStore.ui.dialog;
  if (!dialogState) return;
  const { sheetKey, rowId } = dialogState;
  const rows = getCollection(sheetKey);
  const columns = SHEET_COLUMNS[sheetKey];
  const draft = { ...(rowId ? rows.find((item) => item.id === rowId) : blankRowForSheet(sheetKey)) };

  columns.forEach((column) => {
    const field = elements.editorDialogBody.querySelector(`[data-editor-field="${column}"]`);
    if (!field) return;
    const rawValue = field.value;
    if (numericColumns[sheetKey]?.has(column) || (sheetKey === "parameters" && column === "value" && draft.type === "number")) {
      draft[column] = rawValue === "" ? "" : Number(rawValue);
    } else if (sheetKey === "parameters" && column === "editable") {
      draft[column] = /^true$/i.test(rawValue);
    } else {
      draft[column] = rawValue;
    }
  });

  if (rowId) {
    const index = rows.findIndex((item) => item.id === rowId);
    rows[index] = draft;
  } else {
    rows.unshift(draft);
  }

  stateStore.ui.dialog = null;
  elements.rowEditorDialog.close();
  recomputeAndRender();
}

function deleteSheetRow(sheetKey, rowId) {
  const rows = getCollection(sheetKey);
  const index = rows.findIndex((item) => item.id === rowId);
  if (index === -1) return;
  rows.splice(index, 1);
  recomputeAndRender();
}

function applyProjectionPreset(preset) {
  const next = { ...PROJECTION_BASELINE };
  if (preset === "wasp10") next.waspFactor = 0.85;
  if (preset === "wasp20") next.waspFactor = 0.75;
  if (preset === "bio10") next.bioBlend = 10;
  if (preset === "bio20") next.bioBlend = 20;
  if (preset === "rfnbo10") next.rfnboBlend = 10;
  if (preset === "rfnbo20") next.rfnboBlend = 20;
  if (preset === "combined") {
    next.waspFactor = 0.80;
    next.bioBlend = 20;
    next.bioType = "HVO";
    next.rfnboBlend = 10;
    next.rfnboType = "e-diesel";
  }
  stateStore.ui.projection = next;
  stateStore.ui.projectionPreset = preset;
  render();
}

function updateProjectionField(key, rawValue) {
  const numericKeys = new Set(["waspFactor", "bioBlend", "rfnboBlend"]);
  stateStore.ui.projection[key] = numericKeys.has(key) ? Number(rawValue) : rawValue;
  stateStore.ui.projectionPreset = "custom";
  render();
}

function handleLibraryClick(event) {
  const actionTarget = event.target.closest("[data-action]");
  if (!actionTarget) return;
  const { action, sheet, rowId } = actionTarget.dataset;

  if (action === "select-library-sheet") {
    stateStore.ui.librarySheet = sheet;
    stateStore.ui.librarySearch = "";
    stateStore.ui.libraryPage = 1;
    renderLibraryDrawer();
    return;
  }

  if (action === "library-page-prev") {
    stateStore.ui.libraryPage = Math.max(1, stateStore.ui.libraryPage - 1);
    renderLibraryDrawer();
    return;
  }

  if (action === "library-page-next") {
    stateStore.ui.libraryPage += 1;
    renderLibraryDrawer();
    return;
  }

  if (action === "open-row-editor") {
    openEditorDialog(sheet, rowId);
    return;
  }

  if (action === "delete-sheet-row") {
    deleteSheetRow(sheet, rowId);
  }
}

function handleLibraryInput(event) {
  if (event.target.dataset.action === "library-search") {
    stateStore.ui.librarySearch = event.target.value;
    stateStore.ui.libraryPage = 1;
    renderLibraryDrawer();
  }
}

function handleMainClick(event) {
  const actionTarget = event.target.closest("[data-action]");
  if (!actionTarget) return;
  const { action, view, rowId } = actionTarget.dataset;

  if (action === "select-view") {
    stateStore.ui.activeView = view;
    render();
    return;
  }

  if (action === "select-calculator-row") {
    if (event.target.closest("input, select, textarea, option")) {
      stateStore.ui.calculatorSelectedId = rowId;
      return;
    }
    stateStore.ui.calculatorSelectedId = rowId;
    render();
    return;
  }

  if (action === "add-calculator-row") {
    const currentRow = getCalculatorStateRow(stateStore.ui.calculatorSelectedId);
    if (!currentRow || !rowHasMeaningfulInputs(currentRow)) {
      stateStore.ui.calculatorActiveScrollLeft = 0;
      render();
      window.requestAnimationFrame(() => {
        elements.contentView.querySelector('[data-calc-cell="imoNo"]')?.focus();
      });
      showToast("Complete the current row before adding another voyage.", "info");
      return;
    }
    const row = buildCalculatorRowForCurrentFilter(currentRow);
    insertCalculatorRow(row, currentRow.id);
    stateStore.ui.calculatorSelectedId = row.id;
    stateStore.ui.calculatorActiveScrollLeft = 0;
    stateStore.ui.calculatorHistoryPage = 1;
    recomputeAndRender();
    window.requestAnimationFrame(() => {
      elements.contentView.querySelector('[data-calc-cell="imoNo"]')?.focus();
    });
    showToast(`New voyage input ${row.recordId} is ready.`, "success");
    return;
  }

  if (action === "toggle-calculator-columns") {
    stateStore.ui.calculatorColumnMenuOpen = !stateStore.ui.calculatorColumnMenuOpen;
    render();
    return;
  }

  if (action === "reset-calculator-columns") {
    stateStore.ui.calculatorVisibleColumns = CALCULATOR_COLUMNS
      .filter((column) => column.key !== "rowActions")
      .map((column) => column.key);
    stateStore.ui.calculatorColumnMenuOpen = true;
    render();
    return;
  }

  if (action === "toggle-calculator-column") {
    const columnKey = actionTarget.dataset.columnKey;
    const current = new Set(getVisibleCalculatorColumnKeys().filter((key) => key !== "rowActions"));
    if (current.has(columnKey)) {
      if (current.size > 1) {
        current.delete(columnKey);
      }
    } else {
      current.add(columnKey);
    }
    stateStore.ui.calculatorVisibleColumns = CALCULATOR_COLUMNS
      .filter((column) => column.key !== "rowActions" && current.has(column.key))
      .map((column) => column.key);
    stateStore.ui.calculatorColumnMenuOpen = true;
    render();
    return;
  }

  if (action === "edit-calculator-row") {
    stateStore.ui.calculatorSelectedId = rowId;
    render();
    return;
  }

  if (action === "duplicate-calculator-row") {
    const selected = stateStore.state.calculatorRows.find((row) => row.id === stateStore.ui.calculatorSelectedId);
    if (!selected) return;
    const copy = { ...deepClone(selected), id: `calc-${Date.now()}` };
    stateStore.state.calculatorRows.unshift(copy);
    stateStore.ui.calculatorSelectedId = copy.id;
    recomputeAndRender();
    return;
  }

  if (action === "delete-calculator-row") {
    const targetRowId = rowId || stateStore.ui.calculatorSelectedId;
    const index = stateStore.state.calculatorRows.findIndex((row) => row.id === targetRowId);
    if (index === -1) return;
    stateStore.state.calculatorRows.splice(index, 1);
    stateStore.ui.calculatorSelectedId = null;
    stateStore.ui.calculatorHistoryPage = 1;
    recomputeAndRender();
    showToast("Voyage row deleted.", "success");
    return;
  }

  if (action === "insert-calculator-after") {
    const sourceStateRow = getCalculatorStateRow(rowId);
    const row = buildCalculatorRowForCurrentFilter(sourceStateRow);
    insertCalculatorRow(row, rowId);
    stateStore.ui.calculatorSelectedId = row.id;
    stateStore.ui.calculatorActiveScrollLeft = 0;
    stateStore.ui.calculatorHistoryPage = 1;
    recomputeAndRender();
    showToast(`Inserted ${row.recordId} below the selected voyage.`, "success");
    return;
  }

  if (action === "close-drilldown") {
    closeDrilldown();
    return;
  }

  if (action === "toggle-projection-panel") {
    stateStore.ui.projectionOpen = !stateStore.ui.projectionOpen;
    render();
    return;
  }

  if (action === "toggle-kpis") {
    stateStore.ui.kpisOpen = !stateStore.ui.kpisOpen;
    render();
    return;
  }

  if (action === "toggle-charts") {
    stateStore.ui.chartsOpen = !stateStore.ui.chartsOpen;
    render();
    return;
  }

  if (action === "toggle-voyage-table") {
    stateStore.ui.voyageTableOpen = !stateStore.ui.voyageTableOpen;
    render();
    return;
  }

  if (action === "toggle-calculator") {
    stateStore.ui.calculatorOpen = !stateStore.ui.calculatorOpen;
    render();
    return;
  }

  if (action === "toggle-calculator-history") {
    stateStore.ui.calculatorHistoryOpen = !stateStore.ui.calculatorHistoryOpen;
    render();
    return;
  }

  if (action === "calculator-history-prev") {
    stateStore.ui.calculatorHistoryPage = Math.max(1, stateStore.ui.calculatorHistoryPage - 1);
    stateStore.ui.calculatorHistoryScrollTop = 0;
    render();
    return;
  }

  if (action === "calculator-history-next") {
    stateStore.ui.calculatorHistoryPage += 1;
    stateStore.ui.calculatorHistoryScrollTop = 0;
    render();
    return;
  }

  if (action === "apply-projection-preset") {
    applyProjectionPreset(actionTarget.dataset.preset);
    return;
  }

  if (action === "open-kpi-drilldown") {
    openKpiDrilldown(actionTarget.dataset.kpi);
  }
}

function handleMainInput(event) {
  const projectionKey = event.target.dataset.projectionKey;
  if (projectionKey) {
    updateProjectionField(projectionKey, event.target.value);
    return;
  }

  const calcCell = event.target.dataset.calcCell;
  if (calcCell) {
    updateInlineSuggestions(event.target, calcCell, event.target.value);
    const rowContext = event.target.dataset.rowContext || "active";
    if (event.type === "input") {
      updateCalculatorCell(event.target.dataset.rowId, calcCell, event.target.value, false, rowContext);
      return;
    }
    updateCalculatorCell(event.target.dataset.rowId, calcCell, event.target.value, true, rowContext);
    return;
  }

  const calcField = event.target.dataset.calcField;
  if (calcField) {
    if (event.type === "input") {
      const row = stateStore.state.calculatorRows.find((item) => item.id === stateStore.ui.calculatorSelectedId);
      if (row) {
        setCalculatorRowValue(row, calcField, event.target.value);
      }
      return;
    }
    updateCalculatorField(calcField, event.target.value);
    return;
  }

  if (event.target.dataset.action === "calculator-search") {
    stateStore.ui.calculatorSearch = event.target.value;
    stateStore.ui.calculatorHistoryPage = 1;
    stateStore.ui.calculatorHistoryScrollTop = 0;
    render();
    return;
  }

  if (event.target.dataset.action === "detail-search") {
    stateStore.ui.detailSearch = event.target.value;
    render();
    return;
  }

  if (event.target.dataset.action === "detail-scope") {
    stateStore.ui.detailScope = event.target.value;
    render();
  }
}

async function bootstrap() {
  const response = await fetch("/data/workbook-seed.json");
  const seed = await response.json();
  stateStore.seedState = createStateFromSeed(seed);
  const localState = hydrateFromStorage(stateStore.seedState);
  stateStore.state = await hydrateFromServer(localState);
  compactCalculatorRowsForRuntime();
  stateStore.derived = recalculateWorkbook(stateStore.state);
  if (syncCalculatorDraftRowsWithDerived()) {
    stateStore.derived = recalculateWorkbook(stateStore.state);
  }
  stateStore.state.parameters = deepClone(stateStore.derived.parameters);
  ensureCalculatorSelection();
  saveState();

  elements.viewTabs.addEventListener("click", handleMainClick);
  elements.contentView.addEventListener("click", handleMainClick);
  elements.contentView.addEventListener("input", handleMainInput);
  elements.contentView.addEventListener("change", handleMainInput);
  elements.vesselFilter.addEventListener("change", (event) => {
    stateStore.ui.vesselFilter = event.target.value;
    stateStore.ui.calculatorSelectedId = null;
    stateStore.ui.calculatorHistoryPage = 1;
    stateStore.ui.calculatorActiveScrollLeft = 0;
    stateStore.ui.calculatorHistoryScrollLeft = 0;
    stateStore.ui.calculatorHistoryScrollTop = 0;
    stateStore.ui.drilldown = null;
    stateStore.ui.detailSearch = "";
    stateStore.ui.detailScope = "all";
    render();
  });
  elements.libraryToggleButton.addEventListener("click", () => {
    stateStore.ui.libraryOpen = !stateStore.ui.libraryOpen;
    renderLibraryDrawer();
  });
  elements.closeLibraryButton.addEventListener("click", () => {
    stateStore.ui.libraryOpen = false;
    renderLibraryDrawer();
  });
  elements.libraryBackdrop.addEventListener("click", () => {
    stateStore.ui.libraryOpen = false;
    renderLibraryDrawer();
  });
  elements.libraryTabs.addEventListener("click", handleLibraryClick);
  elements.libraryContent.addEventListener("click", handleLibraryClick);
  elements.libraryContent.addEventListener("input", handleLibraryInput);
  elements.exportFilteredButton.addEventListener("click", exportFilteredData);
  elements.generateReportButton.addEventListener("click", openReportDialog);
  elements.closeReportButton.addEventListener("click", () => elements.reportDialog.close());
  elements.reportVesselSelect.addEventListener("change", renderReportRows);
  elements.reportDialog.addEventListener("change", (event) => {
    if (event.target.name === "reportBasis" || event.target.matches('.report-row-option input[type="checkbox"]')) {
      updateReportSelectionSummary();
    }
  });
  elements.selectAllReportRowsButton.addEventListener("click", () => {
    elements.reportRowList.querySelectorAll('input[type="checkbox"]').forEach((item) => {
      item.checked = true;
    });
    updateReportSelectionSummary();
  });
  elements.clearReportRowsButton.addEventListener("click", () => {
    elements.reportDialog.querySelector('input[name="reportBasis"][value="selected"]').checked = true;
    elements.reportRowList.querySelectorAll('input[type="checkbox"]').forEach((item) => {
      item.checked = false;
    });
    updateReportSelectionSummary();
  });
  elements.generateReportConfirmButton.addEventListener("click", generateComplianceReport);
  if (elements.resetWorkbookButton) {
    elements.resetWorkbookButton.addEventListener("click", () => {
      stateStore.state = deepClone(stateStore.seedState);
      compactCalculatorRowsForRuntime();
      stateStore.ui.calculatorSelectedId = null;
      stateStore.ui.drilldown = null;
      stateStore.ui.detailSearch = "";
      stateStore.ui.detailScope = "all";
      stateStore.ui.projection = { ...PROJECTION_BASELINE };
      stateStore.ui.projectionPreset = "baseline";
      recomputeAndRender();
    });
  }
  elements.closeEditorButton.addEventListener("click", () => {
    stateStore.ui.dialog = null;
    elements.rowEditorDialog.close();
  });
  elements.saveEditorButton.addEventListener("click", saveEditorDialog);

  window.fuelEtsDashboard = {
    build: APP_BUILD,
    importVoyageRows: importVoyageRowsFromExternal,
    exportState: () => deepClone(stateStore.state),
    getBlankVoyageTemplate: () => deepClone(blankCalculatorRow()),
    syncNow: syncStateToServer,
  };

  render();
  loadEuaMarketSnapshot();
}

bootstrap().catch((error) => {
  elements.contentView.innerHTML = `<div class="empty-state">Failed to load dashboard: ${error.message}</div>`;
});
