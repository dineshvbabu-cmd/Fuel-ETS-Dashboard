const fs = require("fs/promises");
const path = require("path");
const crypto = require("crypto");
const { parseComplianceWorkbook } = require("./excel-import");
const { loadStateDocument, saveStateDocument } = require("./storage");

const SEED_PATH = path.join(__dirname, "compliance_dashboard", "data", "workbook-seed.json");
const POLL_INTERVAL_MS = Number(process.env.WORKBOOK_SYNC_INTERVAL_MS || 5 * 60 * 1000);

const syncState = {
  configured: false,
  enabled: false,
  sourceUrl: "",
  status: "disabled",
  running: false,
  startedAt: "",
  finishedAt: "",
  lastSuccessAt: "",
  lastAttemptAt: "",
  lastError: "",
  lastWarnings: [],
  workbookHash: "",
  workbookFileName: "",
  appliedSections: [],
  documentRevision: "",
  pollIntervalMs: POLL_INTERVAL_MS,
};

let syncTimer = null;
let syncPromise = null;
let seedCache = null;

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeUpper(value) {
  return normalizeText(value).toUpperCase();
}

function numberOrZero(value) {
  if (value === null || value === undefined || value === "") return 0;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function maybeNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function dateToInputValue(value) {
  if (!value) return "";
  if (typeof value === "string") return value.slice(0, 10);
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeParameter(row, index) {
  return {
    id: `parameter-${index + 1}`,
    section: normalizeText(row.section),
    key: normalizeText(row.key),
    label: normalizeText(row.label),
    value: row.value,
    note: normalizeText(row.note),
    editable: Boolean(row.editable),
    type: normalizeText(row.type) || "text",
  };
}

function normalizeFuelRow(row, index) {
  return {
    id: `fuel-${index + 1}`,
    fuelPathway: normalizeText(row["Fuel Pathway"]) || "(none)",
    fuelClass: normalizeText(row.Class),
    lcvMjPerG: numberOrZero(row["LCV (MJ/g)"]),
    wtWPerMj: numberOrZero(row["WtT (gCOâ‚‚eq/MJ)"]),
    rwd: numberOrZero(row["RWD (RFNBO reward)"]) || 1,
    etsCo2Cf: numberOrZero(row["ETS COâ‚‚ Cf (tCOâ‚‚/t)"]),
    notes: normalizeText(row.Notes),
    alias: normalizeText(row["Column 11"]) || "(none)",
    cfCo2PerG: numberOrZero(row["Cf COâ‚‚ (g/gFuel)"]),
    cfCh4PerG: numberOrZero(row["Cf CHâ‚„ (g/gFuel)"]),
    cfN2oPerG: numberOrZero(row["Cf Nâ‚‚O (g/gFuel)"]),
    cslipPercent: numberOrZero(row["Cslip (% mass)"]),
    consumerSource: normalizeText(row["Fuel consumer / source (Annex II)"]),
  };
}

function normalizeFleetRow(row, index) {
  return {
    id: `fleet-${index + 1}`,
    imoNo: numberOrZero(row["IMO No."]),
    vesselName: normalizeText(row["Vessel Name"]),
    shipType: normalizeText(row["Ship Type"]),
    flag: normalizeText(row.Flag),
    className: normalizeText(row.Class),
    gt: numberOrZero(row["GT (GRT)"]),
    nt: numberOrZero(row["NT (NRT)"]),
    summerDwt: numberOrZero(row["Summer DWT"]),
    built: row.Built,
    wapsFwindFactor: maybeNumber(row["WAPS Fwind factor"]) ?? maybeNumber(row.wapsFwindFactor),
  };
}

function normalizePortRow(row, index) {
  return {
    id: `port-${index + 1}`,
    unlocode: normalizeUpper(row["UN/LOCODE"]),
    portName: normalizeText(row["Port Name"]),
    country: normalizeText(row.Country),
    countryCode: normalizeText(row["Country Code"]),
    euEeaInScope: normalizeText(row["EU/EEA In-Scope"]),
    outermostRegion: normalizeText(row["Outermost Region"]),
    specialCategory: normalizeText(row["Special Category"]),
  };
}

function normalizeFlagRow(row, index) {
  return {
    id: `flag-${index + 1}`,
    flagState: normalizeText(row["Flag State"]),
    iso: normalizeText(row.ISO),
    registryType: normalizeText(row["Registry Type"]),
    euEeaFlag: normalizeText(row["EU/EEA Flag"]),
    notes: normalizeText(row.Notes),
  };
}

function normalizeDerogationRow(row, index) {
  return {
    id: `derogation-${index + 1}`,
    serialNo: row["S.No"] ?? index + 1,
    euMemberState: normalizeText(row["EU Member State"]),
    outermostRegion: normalizeText(row["Outermost Region"]),
    omrPortName: normalizeText(row["OMR Port Name"]),
    unlocode: normalizeUpper(row["UN/LOCODE"]),
  };
}

function normalizeMethodologyRow(row, index) {
  return {
    id: `methodology-${index + 1}`,
    detail: normalizeText(row.detail),
  };
}

function normalizeFormulaRow(row, index) {
  return {
    id: `formula-${index + 1}`,
    stepField: normalizeText(row["Step / Field"]),
    resultColumn: normalizeText(row["Result column"]),
    formulaPlainEnglish: normalizeText(row["Formula (plain English)"]),
  };
}

function normalizeCalculatorRow(row, index) {
  return {
    id: `calc-${index + 1}`,
    recordId: normalizeUpper(row.recordId),
    type: normalizeText(row.type),
    storageYear: maybeNumber(row.storageYear),
    entrySource: normalizeText(row.entrySource) || "manual",
    sourceSystem: normalizeText(row.sourceSystem),
    sourceRecordId: normalizeText(row.sourceRecordId),
    sourceUpdatedAt: normalizeText(row.sourceUpdatedAt),
    imoNo: maybeNumber(row.imoNo),
    departureDate: dateToInputValue(row.departureDate),
    fromPortCode: normalizeUpper(row.fromPortCode),
    arrivalDate: dateToInputValue(row.arrivalDate),
    toPortCode: normalizeUpper(row.toPortCode),
    fuel1Type: normalizeText(row.fuel1Type) || "(none)",
    fuel1ConsumptionMt: maybeNumber(row.fuel1ConsumptionMt),
    fuel2Type: normalizeText(row.fuel2Type) || "(none)",
    fuel2ConsumptionMt: maybeNumber(row.fuel2ConsumptionMt),
    bioFuelType: normalizeText(row.bioFuelType) || "(none)",
    bioFuelConsumptionMt: maybeNumber(row.bioFuelConsumptionMt),
    sustainabilityFactor: maybeNumber(row.sustainabilityFactor),
    windFactor: maybeNumber(row.windFactor),
    distanceNm: maybeNumber(row.distanceNm),
    cargoTonnes: maybeNumber(row.cargoTonnes),
    timeAtSeaHours: maybeNumber(row.timeAtSeaHours),
    berthHours: maybeNumber(row.berthHours),
    opsElectricityMj: maybeNumber(row.opsElectricityMj),
  };
}

function createStateFromSeed(seed) {
  return {
    meta: {
      generatedAt: seed.generatedAt,
      sourceWorkbook: seed.sourceWorkbook,
    },
    parameters: seed.parameters.map(normalizeParameter),
    fuelReference: seed.fuelReference.map(normalizeFuelRow),
    fleet: seed.fleet.map(normalizeFleetRow),
    ports: seed.ports.map(normalizePortRow),
    flags: seed.flags.map(normalizeFlagRow),
    derogations: seed.derogations.map(normalizeDerogationRow),
    methodology: seed.methodology.map(normalizeMethodologyRow),
    formulaGuide: seed.formulaGuide.map(normalizeFormulaRow),
    calculatorRows: seed.calculator.rows.map(normalizeCalculatorRow),
  };
}

async function loadSeedState() {
  if (seedCache) return deepClone(seedCache);
  const raw = JSON.parse(await fs.readFile(SEED_PATH, "utf8"));
  seedCache = createStateFromSeed(raw);
  return deepClone(seedCache);
}

function getConfiguredSourceUrl() {
  return normalizeText(process.env.WORKBOOK_SYNC_SOURCE_URL);
}

function parseHeaders() {
  const headerJson = normalizeText(process.env.WORKBOOK_SYNC_HEADERS_JSON);
  if (!headerJson) return {};
  try {
    const parsed = JSON.parse(headerJson);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    throw new Error("WORKBOOK_SYNC_HEADERS_JSON must be valid JSON.");
  }
}

function ensureWorkbookPayload(response, buffer) {
  const finalUrl = response.url || "";
  const contentType = normalizeText(response.headers.get("content-type")).toLowerCase();
  const looksLikeZip = buffer.length >= 4 && buffer[0] === 0x50 && buffer[1] === 0x4b;
  if (looksLikeZip) {
    return;
  }
  if (contentType.includes("text/html") || /login\.microsoftonline\.com/i.test(finalUrl)) {
    throw new Error("Workbook source requires SharePoint or Microsoft 365 authentication from the server.");
  }
  throw new Error(`Workbook source returned an unexpected response (${contentType || "unknown content type"}).`);
}

async function downloadWorkbookBuffer(sourceUrl) {
  const response = await fetch(sourceUrl, {
    headers: {
      Accept: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
      ...parseHeaders(),
    },
    redirect: "follow",
  });
  if (!response.ok) {
    throw new Error(`Workbook source responded with ${response.status}.`);
  }
  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  ensureWorkbookPayload(response, buffer);
  return { buffer, finalUrl: response.url || sourceUrl };
}

function applyWorkbookSections(baseState, preview, sourceUrl) {
  const nextState = deepClone(baseState);
  const sectionKeys = Object.keys(preview.sections);
  sectionKeys.forEach((key) => {
    nextState[key] = deepClone(preview.sections[key].rows);
  });
  nextState.meta = {
    ...(nextState.meta || {}),
    sourceWorkbook: preview.fileName,
    importedAt: preview.importedAt,
    importedSheets: sectionKeys.map((key) => preview.sections[key].sourceSheet),
    workbookSync: {
      sourceUrl,
      lastWorkbookFileName: preview.fileName,
      lastImportedAt: preview.importedAt,
      lastAppliedSections: sectionKeys,
    },
  };
  return nextState;
}

async function loadBaseState() {
  const current = await loadStateDocument();
  if (current?.state) return deepClone(current.state);
  return loadSeedState();
}

function updateSyncState(patch) {
  Object.assign(syncState, patch);
}

async function runWorkbookSync({ force = false, reason = "manual" } = {}) {
  const sourceUrl = getConfiguredSourceUrl();
  updateSyncState({
    configured: Boolean(sourceUrl),
    enabled: Boolean(sourceUrl),
    sourceUrl,
    pollIntervalMs: POLL_INTERVAL_MS,
  });

  if (!sourceUrl) {
    updateSyncState({
      status: "disabled",
      lastError: "Set WORKBOOK_SYNC_SOURCE_URL to enable workbook sync.",
    });
    return getWorkbookSyncStatus();
  }

  if (syncPromise) return syncPromise;

  syncPromise = (async () => {
    const startedAt = new Date().toISOString();
    updateSyncState({
      running: true,
      status: "syncing",
      startedAt,
      lastAttemptAt: startedAt,
      lastError: "",
    });
    try {
      const { buffer } = await downloadWorkbookBuffer(sourceUrl);
      const workbookHash = crypto.createHash("sha256").update(buffer).digest("hex");
      if (!force && syncState.workbookHash && syncState.workbookHash === workbookHash) {
        updateSyncState({
          status: "up_to_date",
          finishedAt: new Date().toISOString(),
          running: false,
        });
        return getWorkbookSyncStatus();
      }

      const preview = await parseComplianceWorkbook(buffer, path.basename(new URL(sourceUrl).pathname) || "workbook.xlsx");
      const baseState = await loadBaseState();
      const nextState = applyWorkbookSections(baseState, preview, sourceUrl);
      const document = await saveStateDocument(nextState, `workbook-sync:${reason}`);

      updateSyncState({
        status: "ok",
        finishedAt: new Date().toISOString(),
        lastSuccessAt: new Date().toISOString(),
        lastWarnings: preview.warnings || [],
        lastError: "",
        workbookHash,
        workbookFileName: preview.fileName,
        appliedSections: Object.keys(preview.sections),
        documentRevision: document.revision,
        running: false,
      });
      return getWorkbookSyncStatus();
    } catch (error) {
      updateSyncState({
        status: "error",
        finishedAt: new Date().toISOString(),
        lastError: error.message,
        running: false,
      });
      return getWorkbookSyncStatus();
    } finally {
      syncPromise = null;
    }
  })();

  return syncPromise;
}

function startWorkbookSyncScheduler() {
  const sourceUrl = getConfiguredSourceUrl();
  updateSyncState({
    configured: Boolean(sourceUrl),
    enabled: Boolean(sourceUrl),
    sourceUrl,
    pollIntervalMs: POLL_INTERVAL_MS,
    status: sourceUrl ? "idle" : "disabled",
    lastError: sourceUrl ? "" : "Set WORKBOOK_SYNC_SOURCE_URL to enable workbook sync.",
  });
  if (!sourceUrl) return;
  if (syncTimer) clearInterval(syncTimer);
  syncTimer = setInterval(() => {
    runWorkbookSync({ force: false, reason: "interval" }).catch(() => {});
  }, POLL_INTERVAL_MS);
  setTimeout(() => {
    runWorkbookSync({ force: false, reason: "startup" }).catch(() => {});
  }, 500);
}

function getWorkbookSyncStatus() {
  return deepClone(syncState);
}

module.exports = {
  getWorkbookSyncStatus,
  runWorkbookSync,
  startWorkbookSyncScheduler,
};
