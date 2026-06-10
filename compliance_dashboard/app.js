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

const PAGE_SIZE = 18;
const REFERENCE_SHEETS = SHEETS.filter((sheet) => !["dashboard", "calculator", "vesselSummary"].includes(sheet.key));
const PROJECTION_BASELINE = {
  waspFactor: 0.95,
  bioBlend: 0,
  bioType: "Bio-diesel",
  rfnboBlend: 0,
  rfnboType: "e-diesel",
};

const elements = {
  viewTabs: document.getElementById("viewTabs"),
  vesselFilter: document.getElementById("vesselFilter"),
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
};

const stateStore = {
  seedState: null,
  state: null,
  derived: null,
  charts: {},
  ui: {
    activeView: "dashboard",
    vesselFilter: "all",
    calculatorSearch: "",
    calculatorSelectedId: null,
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
  fleet: new Set(["imoNo", "gt", "nt", "summerDwt", "built"]),
  derogations: new Set(["serialNo"]),
};

const CALCULATOR_COLUMNS = [
  { key: "recordId", label: "Voyage / Port-Stay ID", kind: "sticky", width: 110 },
  { key: "type", label: "Type", kind: "sticky", width: 100 },
  { key: "imoNo", label: "IMO No.", kind: "editable", input: "number", width: 110, list: "imoNumbers" },
  { key: "vesselName", label: "Vessel Name", kind: "sticky", width: 150 },
  { key: "shipType", label: "Ship Type", kind: "sticky", width: 150 },
  { key: "flagState", label: "Flag State", kind: "sticky", width: 120 },
  { key: "deadweightTonnes", label: "Deadweight (DWT, t)", kind: "sticky-number", width: 125, digits: 0 },
  { key: "netTonnage", label: "Net Tonnage (NT)", kind: "sticky-number", width: 120, digits: 0 },
  { key: "grossTonnage", label: "Gross Tonnage (GT)", kind: "sticky-number", width: 120, digits: 0 },
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
  { key: "eeoi", label: "EEOI", kind: "calculated-number", width: 100, digits: 3 },
  { key: "rowActions", label: "Actions", kind: "actions", width: 100 },
];

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

function formatPercent(value, digits = 0) {
  return `${formatNumber((Number(value) || 0) * 100, digits)}%`;
}

function lower(value) {
  return normalizeText(value).toLowerCase();
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(persistableState(stateStore.state)));
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
  stateStore.state.parameters = deepClone(stateStore.derived.parameters);
  saveState();
  render();
}

function destroyCharts() {
  Object.values(stateStore.charts).forEach((chart) => chart.destroy());
  stateStore.charts = {};
}

function buildDataLists() {
  elements.portCodes.innerHTML = stateStore.state.ports
    .slice(0, 12000)
    .map((row) => `<option value="${row.unlocode}">${row.portName}</option>`)
    .join("");

  elements.fuelTypes.innerHTML = stateStore.derived.fuelReference
    .map((row) => `<option value="${row.fuelPathway}">${row.fuelClass}</option>`)
    .join("");

  elements.imoNumbers.innerHTML = stateStore.state.fleet
    .map((row) => `<option value="${row.imoNo}">${row.vesselName}</option>`)
    .join("");
}

function renderViewTabs() {
  elements.viewTabs.innerHTML = [
    ["dashboard", "Dashboard"],
    ["calculator", "Calculator"],
  ]
    .map(
      ([key, label]) => `
        <button
          class="view-tab ${stateStore.ui.activeView === key ? "active" : ""}"
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
  const allRows = stateStore.derived.calculatorRows.filter((row) => row.recordId);
  if (stateStore.ui.vesselFilter === "all") {
    return allRows;
  }
  return allRows.filter((row) => row.vesselName === stateStore.ui.vesselFilter);
}

function getVisibleVessels() {
  return [...new Set(stateStore.derived.calculatorRows.filter((row) => row.recordId).map((row) => row.vesselName).filter(Boolean))].sort();
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

function calculatorInputValue(row, column) {
  const value = row[column.key];
  return value ?? "";
}

function renderCalculatorCell(row, inputRow, column, stickyLeft) {
  const classes = ["calculator-cell"];
  const styleParts = [`min-width:${column.width}px`, `width:${column.width}px`];

  if (column.kind.startsWith("sticky")) {
    classes.push("sticky-cell");
    styleParts.push(`left:${stickyLeft}px`);
  }

  if (column.kind.startsWith("editable")) {
    classes.push("editable-cell");
  }

  if (column.kind.startsWith("calculated")) {
    classes.push("calculated-cell");
  }

  if (column.kind.includes("number") || column.kind.includes("currency") || column.key === "rowActions") {
    classes.push("number-cell");
  }

  if (column.key === "rowActions") {
    return `
      <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
        <button class="inline-button compact-button" type="button" data-action="delete-calculator-row" data-row-id="${row.id}">Delete</button>
      </td>
    `;
  }

  if (column.kind.startsWith("editable")) {
    return `
      <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
        <input
          class="calculator-grid-input input-orange"
          data-calc-cell="${column.key}"
          data-row-id="${row.id}"
          type="${column.input}"
          value="${calculatorInputValue(inputRow, column)}"
          ${column.list ? `list="${column.list}"` : ""}
          ${column.step ? `step="${column.step}"` : ""}
        >
      </td>
    `;
  }

  return `
    <td class="${classes.join(" ")}" style="${styleParts.join(";")}">
      ${calculatorCellValue(row, column)}
    </td>
  `;
}

function openDrilldown(title, subtitle, columns, rows) {
  stateStore.ui.drilldown = { title, subtitle, columns, rows };
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
      ])
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
      ])
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
      ])
    );
    return;
  }

  if (kpiKey === "penalty") {
    openDrilldown(
      "FuelEU penalty",
      "Penalty-bearing rows for the current filter.",
      ["Record", "Vessel", "Route", "Compliance Balance", "Penalty"],
      activeRows
        .filter((row) => numberOrZero(row.fuelEuPenaltyEur) > 0)
        .map((row) => [
          row.recordId,
          row.vesselName,
          row.route,
          formatNumber(row.complianceBalanceT, 3),
          formatCurrency(row.fuelEuPenaltyEur),
        ])
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
      ])
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
      })
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
      ])
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
      ])
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

function renderProjectionVoyageTable() {
  const projectedRows = getProjectionRows(getActiveRows().filter((row) => row.type === "Voyage"));
  const showProjection = projectionIsActive();
  return renderCollapsibleSection({
    action: "toggle-voyage-table",
    title: "Voyage Results Table",
    badges: `${renderSectionBadge(`${projectedRows.length} voyages`)}${showProjection ? renderSectionBadge("Projected columns", "purple") : ""}`,
    note: "Collapse when you want more chart space",
    open: stateStore.ui.voyageTableOpen,
    body: `
      <article class="table-card compact-table-card">
        <div class="table-wrap dense-table-wrap">
          <table>
            <thead>
              <tr>
                <th>Record</th>
                <th>Vessel</th>
                <th>Route</th>
                <th>EUAs</th>
                <th>GHG</th>
                <th>Balance</th>
                ${showProjection ? `<th class="projection-col-head">Proj EUAs</th><th class="projection-col-head">Proj GHG</th><th class="projection-col-head">Proj Balance</th>` : ""}
              </tr>
            </thead>
            <tbody>
              ${projectedRows
                .map((row) => {
                  const euaDelta = formatDelta(row.deltaEuasRequiredT, 2, "", true);
                  const ghgDelta = formatDelta(row.deltaAttainedGhgIntensity, 2, "", true);
                  const balanceDelta = formatDelta(row.deltaComplianceBalanceT, 2, "", false);
                  return `
                    <tr>
                      <td>${row.recordId}</td>
                      <td>${row.vesselName}</td>
                      <td>${row.route}</td>
                      <td>${formatNumber(row.euasRequiredT, 3)}</td>
                      <td>${formatNumber(row.attainedGhgIntensity, 3)}</td>
                      <td class="${toneClass(row.complianceBalanceT || 0)}">${formatNumber(row.complianceBalanceT, 3)}</td>
                      ${
                        showProjection
                          ? `<td class="projection-cell">${formatNumber(row.projectedEuasRequiredT, 3)} <span class="projection-delta ${euaDelta.tone}">${euaDelta.text}</span></td>
                             <td class="projection-cell">${formatNumber(row.projectedAttainedGhgIntensity, 3)} <span class="projection-delta ${ghgDelta.tone}">${ghgDelta.text}</span></td>
                             <td class="projection-cell">${formatNumber(row.projectedComplianceBalanceT, 3)} <span class="projection-delta ${balanceDelta.tone}">${balanceDelta.text}</span></td>`
                          : ""
                      }
                    </tr>
                  `;
                })
                .join("")}
            </tbody>
          </table>
        </div>
      </article>
    `,
  });
}

function renderDashboard() {
  return `
    ${renderProjectionPanel()}
    <section class="dashboard-layout">
      <div class="dashboard-main">
        ${renderCollapsibleSection({
          action: "toggle-charts",
          title: "Visual Analytics",
          badges: renderSectionBadge("6 charts"),
          note: "Collapse when you need more space",
          open: stateStore.ui.chartsOpen,
          body: `
            <div class="chart-grid">
          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>EUAs Required Trend by Voyage (t CO2eq)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="vesselEuaChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>GHG Intensity Attained vs Target (g/MJ)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="voyageGhgChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Voyage</p>
                <h3>EUAs Required by Voyage (t CO2eq)</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="voyageEuaChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">Cost Split</p>
                <h3>EUA Cost Split — Voyage vs Port Stay</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="costSplitChart"></canvas></div>
          </article>
          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">Projection</p>
                <h3>EUAs Actual vs Projected by Voyage</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="projectedEuaChart"></canvas></div>
          </article>

          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">Projection</p>
                <h3>Compliance Balance Actual vs Projected</h3>
              </div>
            </div>
            <div class="chart-canvas-wrap"><canvas id="projectedBalanceChart"></canvas></div>
          </article>
            </div>
          `,
        })}
        ${renderProjectionVoyageTable()}
      </div>
      ${renderDrilldownPane()}
    </section>
  `;
}

function renderDashboardCharts() {
  destroyCharts();
  const activeRows = getActiveRows();
  const dashboard = computeFilteredDashboard(activeRows);
  const voyageRows = dashboard.voyageRows.filter((row) => row.attainedGhgIntensity !== null);
  const euaTrendRows = dashboard.voyageRows.filter((row) => row.euasRequiredT !== null);
  const projectedRows = getProjectionRows(dashboard.voyageRows);

  const vesselCanvas = document.getElementById("vesselEuaChart");
  const ghgCanvas = document.getElementById("voyageGhgChart");
  const voyageEuaCanvas = document.getElementById("voyageEuaChart");
  const costSplitCanvas = document.getElementById("costSplitChart");
  const projectedEuaCanvas = document.getElementById("projectedEuaChart");
  const projectedBalanceCanvas = document.getElementById("projectedBalanceChart");

  if (vesselCanvas) {
    stateStore.charts.vesselEua = new Chart(vesselCanvas, {
      type: "line",
      data: {
        labels: euaTrendRows.map((row) => row.recordId),
        datasets: [
          {
            label: "EUAs required",
            data: euaTrendRows.map((row) => row.euasRequiredT),
            borderColor: "#d78a1f",
            backgroundColor: "rgba(215, 138, 31, 0.16)",
            pointBackgroundColor: "#d78a1f",
            pointBorderColor: "#d78a1f",
            pointRadius: 4,
            pointHoverRadius: 6,
            borderWidth: 3,
            tension: 0.28,
            fill: true,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const row = euaTrendRows[index];
          if (!row) return;
          openDrilldown(
            `Voyage EUA trend ${row.recordId}`,
            row.route,
            ["Record", "Vessel", "Type", "EUAs Required", "ETS Cost", "GHG Intensity"],
            [[
              row.recordId,
              row.vesselName,
              row.type,
              formatNumber(row.euasRequiredT, 3),
              formatCurrency(row.euasCostEur),
              formatNumber(row.attainedGhgIntensity, 3),
            ]]
          );
        },
        scales: {
          y: {
            beginAtZero: true,
          },
        },
      },
    });
  }

  if (ghgCanvas) {
    stateStore.charts.voyageGhg = new Chart(ghgCanvas, {
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
        maintainAspectRatio: false,
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const row = voyageRows[index];
          openDrilldown(
            `GHG Intensity for ${row.recordId}`,
            row.route,
            ["Record", "Vessel", "Attained", "Target", "Compliance Balance", "FuelEU Penalty"],
            [[
              row.recordId,
              row.vesselName,
              formatNumber(row.attainedGhgIntensity, 3),
              formatNumber(row.targetGhgIntensity, 3),
              formatNumber(row.complianceBalanceT, 3),
              formatCurrency(row.fuelEuPenaltyEur),
            ]]
          );
        },
      },
    });
  }

  if (voyageEuaCanvas) {
    stateStore.charts.voyageEua = new Chart(voyageEuaCanvas, {
      type: "bar",
      data: {
        labels: voyageRows.map((row) => row.recordId),
        datasets: [
          {
            label: "EUAs required",
            data: voyageRows.map((row) => row.euasRequiredT),
            backgroundColor: voyageRows.map((row) => (row.scopePercent === 1 ? "#178c18" : "#78a641")),
            borderRadius: 6,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const row = voyageRows[index];
          openDrilldown(
            `Voyage EUA Requirement ${row.recordId}`,
            row.route,
            ["Vessel", "Scope", "ETS CO2eq", "EUAs Required", "ETS Cost", "Fuel Consumed (MT)"],
            [[
              row.vesselName,
              formatPercent(row.scopePercent),
              formatNumber(row.etsInScopeCo2eqT, 3),
              formatNumber(row.euasRequiredT, 3),
              formatCurrency(row.euasCostEur),
              formatNumber(numberOrZero(row.fuel1ConsumptionMt) + numberOrZero(row.fuel2ConsumptionMt) + numberOrZero(row.bioFuelConsumptionMt), 2),
            ]]
          );
        },
      },
    });
  }

  if (costSplitCanvas) {
    const voyageCost = dashboard.voyageRows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0);
    const portStayCost = dashboard.portStayRows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0);
    stateStore.charts.costSplit = new Chart(costSplitCanvas, {
      type: "doughnut",
      data: {
        labels: ["Voyages", "Port stays"],
        datasets: [
          {
            data: [voyageCost, portStayCost],
            backgroundColor: ["#4288d6", "#178c18"],
            borderWidth: 0,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: { legend: { position: "bottom" } },
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const label = index === 0 ? "Voyages" : "Port stays";
          const rows = index === 0 ? dashboard.voyageRows : dashboard.portStayRows;
          openDrilldown(
            `${label} cost split`,
            "Rows included in the selected cost segment.",
            ["Record", "Vessel", "Route", "Type", "EUAs Required", "ETS Cost"],
            rows.map((row) => [
              row.recordId,
              row.vesselName,
              row.route,
              row.type,
              formatNumber(row.euasRequiredT, 3),
              formatCurrency(row.euasCostEur),
            ])
          );
        },
      },
    });
  }

  if (projectedEuaCanvas) {
    stateStore.charts.projectedEua = new Chart(projectedEuaCanvas, {
      type: "bar",
      data: {
        labels: projectedRows.map((row) => row.recordId),
        datasets: [
          {
            label: "Actual",
            data: projectedRows.map((row) => row.euasRequiredT),
            backgroundColor: "#4288d6",
            borderRadius: 6,
          },
          {
            label: "Projected",
            data: projectedRows.map((row) => row.projectedEuasRequiredT),
            backgroundColor: "#7f77dd",
            borderRadius: 6,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const row = projectedRows[index];
          openDrilldown(
            `Actual vs projected EUAs ${row.recordId}`,
            row.route,
            ["Vessel", "Actual EUAs", "Projected EUAs", "Delta", "Actual Cost", "Projected Cost"],
            [[
              row.vesselName,
              formatNumber(row.euasRequiredT, 3),
              formatNumber(row.projectedEuasRequiredT, 3),
              formatNumber(row.deltaEuasRequiredT, 3),
              formatCurrency(row.euasCostEur),
              formatCurrency(row.projectedEuasCostEur),
            ]]
          );
        },
      },
    });
  }

  if (projectedBalanceCanvas) {
    stateStore.charts.projectedBalance = new Chart(projectedBalanceCanvas, {
      type: "line",
      data: {
        labels: projectedRows.map((row) => row.recordId),
        datasets: [
          {
            label: "Actual",
            data: projectedRows.map((row) => row.complianceBalanceT),
            borderColor: "#2b8a3e",
            backgroundColor: "#2b8a3e",
            pointRadius: 4,
            tension: 0.2,
          },
          {
            label: "Projected",
            data: projectedRows.map((row) => row.projectedComplianceBalanceT),
            borderColor: "#7f77dd",
            backgroundColor: "#7f77dd",
            borderDash: [6, 6],
            pointRadius: 3,
            tension: 0.2,
          },
          {
            label: "Zero reference",
            data: projectedRows.map(() => 0),
            borderColor: "rgba(24, 44, 77, 0.35)",
            borderDash: [4, 4],
            pointRadius: 0,
            tension: 0,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const row = projectedRows[index];
          openDrilldown(
            `Balance projection ${row.recordId}`,
            row.route,
            ["Vessel", "Actual Balance", "Projected Balance", "Delta", "Actual GHG", "Projected GHG"],
            [[
              row.vesselName,
              formatNumber(row.complianceBalanceT, 3),
              formatNumber(row.projectedComplianceBalanceT, 3),
              formatNumber(row.deltaComplianceBalanceT, 3),
              formatNumber(row.attainedGhgIntensity, 3),
              formatNumber(row.projectedAttainedGhgIntensity, 3),
            ]]
          );
        },
      },
    });
  }
}

function ensureCalculatorSelection() {
  const rows = stateStore.derived.calculatorRows;
  const selected = rows.find((row) => row.id === stateStore.ui.calculatorSelectedId);
  if (selected) return selected;
  const fallback = rows.find((row) => row.recordId) || rows[0];
  stateStore.ui.calculatorSelectedId = fallback?.id || null;
  return fallback || null;
}

function renderCalculator() {
  ensureCalculatorSelection();
  const searchTerm = lower(stateStore.ui.calculatorSearch);
  const filteredRows = stateStore.derived.calculatorRows.filter((row) => {
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

  const stickyColumns = CALCULATOR_COLUMNS.filter((column) => column.kind.startsWith("sticky"));
  let stickyOffset = 0;
  const stickyOffsets = new Map();
  stickyColumns.forEach((column) => {
    stickyOffsets.set(column.key, stickyOffset);
    stickyOffset += column.width;
  });

  const totalRows = filteredRows.length;
  const activeRows = filteredRows.filter((row) => row.recordId);

  return `
    <section class="calculator-shell">
      <div class="calculator-toolbar">
        <div class="calculator-toolbar-copy">
          <h2>Calculator</h2>
          <p class="helper-text">Orange cells are editable inputs. Blue cells are workbook-driven outputs. Vessel detail columns stay frozen while you scroll across the calculation sheet.</p>
        </div>
        <div class="calculator-actions">
          <input
            class="search-input"
            type="search"
            data-action="calculator-search"
            value="${stateStore.ui.calculatorSearch}"
            placeholder="Search vessel, route, port, IMO, or fuel"
          >
          <button class="inline-button" type="button" data-action="add-calculator-row">Add row</button>
        </div>
      </div>

      ${renderCollapsibleSection({
        action: "toggle-calculator",
        title: `Calculator Workbook${stateStore.ui.vesselFilter !== "all" ? ` for ${stateStore.ui.vesselFilter}` : ""}`,
        badges: `${renderSectionBadge(`${totalRows} rows`)}${renderSectionBadge(`${activeRows.length} active`, "purple")}`,
        note: "Collapse when you need a cleaner view",
        open: stateStore.ui.calculatorOpen,
        body: `
          <article class="table-card compact-table-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Workbook Grid</p>
            <h3>Calculator rows${stateStore.ui.vesselFilter !== "all" ? ` for ${stateStore.ui.vesselFilter}` : ""}</h3>
          </div>
          <span class="chip">${totalRows} rows · ${activeRows.length} active records</span>
        </div>
        <div class="calculator-table-wrap">
          <table class="calculator-table">
            <thead>
              <tr>
                ${CALCULATOR_COLUMNS.map((column) => {
                  const classes = ["calculator-header"];
                  const styleParts = [`min-width:${column.width}px`, `width:${column.width}px`];
                  if (column.kind.startsWith("sticky")) {
                    classes.push("sticky-header");
                    styleParts.push(`left:${stickyOffsets.get(column.key) || 0}px`);
                  } else if (column.kind.startsWith("editable")) {
                    classes.push("editable-header");
                  } else if (column.kind.startsWith("calculated")) {
                    classes.push("calculated-header");
                  } else if (column.kind === "actions") {
                    classes.push("actions-header");
                  }
                  return `<th class="${classes.join(" ")}" style="${styleParts.join(";")}">${column.label}</th>`;
                }).join("")}
              </tr>
            </thead>
            <tbody>
              ${filteredRows
                .map((row) => {
                  const inputRow = stateStore.state.calculatorRows.find((item) => item.id === row.id) || blankCalculatorRow();
                  return `
                    <tr class="${row.id === stateStore.ui.calculatorSelectedId ? "selected-row" : ""}" data-action="select-calculator-row" data-row-id="${row.id}">
                      ${CALCULATOR_COLUMNS.map((column) => renderCalculatorCell(row, inputRow, column, stickyOffsets.get(column.key) || 0)).join("")}
                    </tr>
                  `;
                })
                .join("")}
            </tbody>
          </table>
        </div>
          </article>
        `,
      })}
    </section>
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

function renderLibraryContent() {
  const sheetKey = stateStore.ui.librarySheet;
  const allRows = getSheetRowsForDisplay(sheetKey);
  const search = lower(stateStore.ui.librarySearch);
  const filteredRows = allRows.filter((row) => {
    if (!search) return true;
    return Object.values(row).some((value) => lower(value).includes(search));
  });
  const pageCount = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  if (stateStore.ui.libraryPage > pageCount) {
    stateStore.ui.libraryPage = pageCount;
  }
  const pageRows = filteredRows.slice((stateStore.ui.libraryPage - 1) * PAGE_SIZE, stateStore.ui.libraryPage * PAGE_SIZE);
  const columns = SHEET_COLUMNS[sheetKey];
  const visibleColumns = columns.slice(0, Math.min(columns.length, 5));

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
            ${visibleColumns.map((column) => `<th>${column}</th>`).join("")}
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          ${pageRows
            .map(
              (row) => `
                <tr>
                  ${visibleColumns.map((column) => `<td>${row[column] ?? "-"}</td>`).join("")}
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
  renderLibraryTabs();
  renderLibraryContent();
}

function render() {
  renderViewTabs();
  renderVesselFilter();
  if (stateStore.ui.activeView === "dashboard") {
    elements.kpiGrid.classList.remove("hidden");
    renderKpis();
  } else {
    elements.kpiGrid.classList.add("hidden");
    elements.kpiGrid.innerHTML = "";
  }
  buildDataLists();
  renderContent();
  renderLibraryDrawer();
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"', '""')}"`;
}

function downloadText(filename, content, type = "text/plain") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function exportFilteredData() {
  const activeRows = getActiveRows();
  const vesselSummary = computeFilteredDashboard(activeRows).byVessel;
  const slug = stateStore.ui.vesselFilter === "all" ? "all-vessels" : stateStore.ui.vesselFilter.toLowerCase().replaceAll(/\s+/g, "-");

  const calculatorHeaders = ["Record", "Vessel", "Route", "Type", "EUAs Required", "ETS Cost", "GHG Intensity", "Compliance Balance", "FuelEU Penalty"];
  const calculatorCsv = [
    calculatorHeaders.map(csvEscape).join(","),
    ...activeRows.map((row) =>
      [
        row.recordId,
        row.vesselName,
        row.route,
        row.type,
        formatNumber(row.euasRequiredT, 3),
        formatCurrency(row.euasCostEur),
        formatNumber(row.attainedGhgIntensity, 3),
        formatNumber(row.complianceBalanceT, 3),
        formatCurrency(row.fuelEuPenaltyEur),
      ]
        .map(csvEscape)
        .join(",")
    ),
  ].join("\n");

  const summaryHeaders = ["Vessel", "EUAs Required", "EUA Cost", "Avg GHG Intensity"];
  const summaryCsv = [
    summaryHeaders.map(csvEscape).join(","),
    ...vesselSummary.map((row) =>
      [row.vesselName, formatNumber(row.euasRequired, 3), formatCurrency(row.euasCost), formatNumber(row.averageIntensity, 3)]
        .map(csvEscape)
        .join(",")
    ),
  ].join("\n");

  downloadText(`fuel-ets-${slug}-calculator.csv`, calculatorCsv, "text/csv");
  downloadText(`fuel-ets-${slug}-summary.csv`, summaryCsv, "text/csv");
}

function updateCalculatorField(field, rawValue) {
  const row = stateStore.state.calculatorRows.find((item) => item.id === stateStore.ui.calculatorSelectedId);
  if (!row) return;
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
  recomputeAndRender();
}

function updateCalculatorCell(rowId, field, rawValue) {
  const row = stateStore.state.calculatorRows.find((item) => item.id === rowId);
  if (!row) return;
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
  stateStore.ui.calculatorSelectedId = rowId;
  recomputeAndRender();
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
      const inputType =
        numericColumns[sheetKey]?.has(column) || (sheetKey === "parameters" && column === "value" && draft.type === "number")
          ? "number"
          : "text";
      return `
        <div class="editor-field">
          <label for="editor-${column}">${column}</label>
          ${
            multiline
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
    stateStore.ui.calculatorSelectedId = rowId;
    render();
    return;
  }

  if (action === "add-calculator-row") {
    const row = blankCalculatorRow();
    stateStore.state.calculatorRows.unshift(row);
    stateStore.ui.calculatorSelectedId = row.id;
    recomputeAndRender();
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
    recomputeAndRender();
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
    updateCalculatorCell(event.target.dataset.rowId, calcCell, event.target.value);
    return;
  }

  const calcField = event.target.dataset.calcField;
  if (calcField) {
    updateCalculatorField(calcField, event.target.value);
    return;
  }

  if (event.target.dataset.action === "calculator-search") {
    stateStore.ui.calculatorSearch = event.target.value;
    render();
  }
}

async function bootstrap() {
  const response = await fetch("/data/workbook-seed.json");
  const seed = await response.json();
  stateStore.seedState = createStateFromSeed(seed);
  stateStore.state = hydrateFromStorage(stateStore.seedState);
  stateStore.derived = recalculateWorkbook(stateStore.state);
  stateStore.state.parameters = deepClone(stateStore.derived.parameters);
  ensureCalculatorSelection();

  elements.viewTabs.addEventListener("click", handleMainClick);
  elements.contentView.addEventListener("click", handleMainClick);
  elements.contentView.addEventListener("input", handleMainInput);
  elements.contentView.addEventListener("change", handleMainInput);
  elements.vesselFilter.addEventListener("change", (event) => {
    stateStore.ui.vesselFilter = event.target.value;
    stateStore.ui.calculatorSelectedId = null;
    stateStore.ui.drilldown = null;
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
  elements.resetWorkbookButton.addEventListener("click", () => {
    stateStore.state = deepClone(stateStore.seedState);
    stateStore.ui.calculatorSelectedId = null;
    stateStore.ui.drilldown = null;
    stateStore.ui.projection = { ...PROJECTION_BASELINE };
    stateStore.ui.projectionPreset = "baseline";
    recomputeAndRender();
  });
  elements.closeEditorButton.addEventListener("click", () => {
    stateStore.ui.dialog = null;
    elements.rowEditorDialog.close();
  });
  elements.saveEditorButton.addEventListener("click", saveEditorDialog);

  render();
}

bootstrap().catch((error) => {
  elements.contentView.innerHTML = `<div class="empty-state">Failed to load dashboard: ${error.message}</div>`;
});
