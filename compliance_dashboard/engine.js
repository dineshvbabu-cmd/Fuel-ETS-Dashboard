export const STORAGE_KEY = "eu-ets-fueleu-dashboard-state-v1";

export const SHEETS = [
  { key: "dashboard", label: "Dashboard", kind: "derived" },
  { key: "calculator", label: "Calculator", kind: "calculator" },
  { key: "vesselSummary", label: "Vessel Summary", kind: "derived" },
  { key: "parameters", label: "Parameters", kind: "editor" },
  { key: "fuelReference", label: "Fuel Reference", kind: "editor" },
  { key: "fleet", label: "Fleet DB", kind: "editor" },
  { key: "ports", label: "Port DB", kind: "editor" },
  { key: "flags", label: "Flag States", kind: "editor" },
  { key: "derogations", label: "Derogations", kind: "editor" },
  { key: "methodology", label: "Methodology", kind: "editor" },
  { key: "formulaGuide", label: "Formula Guide", kind: "editor" },
];

export const SHEET_COLUMNS = {
  parameters: ["section", "key", "label", "value", "note", "editable", "type"],
  fuelReference: [
    "fuelPathway",
    "fuelClass",
    "lcvMjPerG",
    "wtWPerMj",
    "rwd",
    "etsCo2Cf",
    "notes",
    "alias",
    "cfCo2PerG",
    "cfCh4PerG",
    "cfN2oPerG",
    "cslipPercent",
    "consumerSource",
  ],
  fleet: ["imoNo", "vesselName", "shipType", "flag", "className", "gt", "nt", "summerDwt", "built"],
  ports: ["unlocode", "portName", "country", "countryCode", "euEeaInScope", "outermostRegion", "specialCategory"],
  flags: ["flagState", "iso", "registryType", "euEeaFlag", "notes"],
  derogations: ["serialNo", "euMemberState", "outermostRegion", "omrPortName", "unlocode"],
  methodology: ["detail"],
  formulaGuide: ["stepField", "resultColumn", "formulaPlainEnglish"],
};

export const CALCULATOR_FIELD_GROUPS = [
  {
    title: "Vessel and Route Inputs",
    fields: [
      { key: "imoNo", label: "IMO No.", type: "number", hint: "Use an IMO from Fleet DB for auto-fill." },
      { key: "departureDate", label: "Departure Date", type: "date" },
      { key: "fromPortCode", label: "From Port UN/LOCODE", type: "text", list: "portCodes" },
      { key: "arrivalDate", label: "Arrival Date", type: "date" },
      { key: "toPortCode", label: "To Port UN/LOCODE", type: "text", list: "portCodes" },
      { key: "distanceNm", label: "Total Distance (nm)", type: "number", step: "0.01" },
      { key: "cargoTonnes", label: "Cargo Carried (t)", type: "number", step: "0.01" },
      { key: "timeAtSeaHours", label: "Time at Sea (h)", type: "number", step: "0.01" },
      { key: "berthHours", label: "Hours at Berth or Anchor", type: "number", step: "0.01" },
      { key: "opsElectricityMj", label: "OPS Electricity (MJ)", type: "number", step: "0.01" },
    ],
  },
  {
    title: "Fuel Inputs",
    fields: [
      { key: "fuel1Type", label: "Fossil Fuel 1 Type", type: "text", list: "fuelTypes" },
      { key: "fuel1ConsumptionMt", label: "Fossil Fuel 1 Consumption (MT)", type: "number", step: "0.0001" },
      { key: "fuel2Type", label: "Fossil Fuel 2 Type", type: "text", list: "fuelTypes" },
      { key: "fuel2ConsumptionMt", label: "Fossil Fuel 2 Consumption (MT)", type: "number", step: "0.0001" },
      { key: "bioFuelType", label: "Biofuel or RFNBO Type", type: "text", list: "fuelTypes" },
      { key: "bioFuelConsumptionMt", label: "Biofuel or RFNBO Consumption (MT)", type: "number", step: "0.0001" },
      { key: "sustainabilityFactor", label: "Sustainability Factor (0-1)", type: "number", step: "0.01" },
      { key: "windFactor", label: "WASP or f_wind", type: "number", step: "0.01" },
    ],
  },
];

export function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

export function numberOrZero(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function maybeNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

export function normalizeText(value) {
  return String(value ?? "").trim();
}

export function normalizeUpper(value) {
  return normalizeText(value).toUpperCase();
}

function round(value, digits = 4) {
  const factor = 10 ** digits;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function percentTone(value, warnAt, riskAt) {
  if (value >= riskAt) return "risk";
  if (value >= warnAt) return "warn";
  return "good";
}

function formatEmptyFuel(value) {
  return normalizeText(value) || "(none)";
}

function dateToInputValue(value) {
  if (!value) {
    return "";
  }
  if (typeof value === "string") {
    return value.slice(0, 10);
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return date.toISOString().slice(0, 10);
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
    fuelPathway: formatEmptyFuel(row["Fuel Pathway"]),
    fuelClass: normalizeText(row.Class),
    lcvMjPerG: numberOrZero(row["LCV (MJ/g)"]),
    wtWPerMj: numberOrZero(row["WtT (gCO₂eq/MJ)"]),
    rwd: numberOrZero(row["RWD (RFNBO reward)"]) || 1,
    etsCo2Cf: numberOrZero(row["ETS CO₂ Cf (tCO₂/t)"]),
    notes: normalizeText(row.Notes),
    alias: formatEmptyFuel(row["Column 11"]),
    cfCo2PerG: numberOrZero(row["Cf CO₂ (g/gFuel)"]),
    cfCh4PerG: numberOrZero(row["Cf CH₄ (g/gFuel)"]),
    cfN2oPerG: numberOrZero(row["Cf N₂O (g/gFuel)"]),
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
    imoNo: maybeNumber(row.imoNo),
    departureDate: dateToInputValue(row.departureDate),
    fromPortCode: normalizeUpper(row.fromPortCode),
    arrivalDate: dateToInputValue(row.arrivalDate),
    toPortCode: normalizeUpper(row.toPortCode),
    fuel1Type: formatEmptyFuel(row.fuel1Type),
    fuel1ConsumptionMt: maybeNumber(row.fuel1ConsumptionMt),
    fuel2Type: formatEmptyFuel(row.fuel2Type),
    fuel2ConsumptionMt: maybeNumber(row.fuel2ConsumptionMt),
    bioFuelType: formatEmptyFuel(row.bioFuelType),
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

export function createStateFromSeed(seed) {
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

function deriveParameterValues(parameterRows) {
  const map = Object.fromEntries(parameterRows.map((row) => [row.key, row.value]));
  const reportYear = numberOrZero(map.reportYear);
  const etsPhaseIn = reportYear <= 2024 ? 0.4 : reportYear === 2025 ? 0.7 : 1;
  const etsGasScope = reportYear >= 2026 ? "CO2 + CH4 + N2O" : "CO2 only";
  const fueleuRef = Number(map.fueleuRef ?? 91.16);
  const fueleuRedux =
    reportYear < 2030
      ? 0.02
      : reportYear < 2035
        ? 0.06
        : reportYear < 2040
          ? 0.145
          : reportYear < 2045
            ? 0.31
            : reportYear < 2050
              ? 0.62
              : 0.8;
  const fueleuTarget = fueleuRef * (1 - fueleuRedux);
  const normalized = {
    reportYear,
    etsPhaseIn,
    etsGasScope,
    bioZero: /^yes$/i.test(normalizeText(map.bioZero)) ? "Yes" : "No",
    euaPrice: numberOrZero(map.euaPrice),
    fueleuRef,
    fueleuRedux,
    fueleuTarget,
    vlsfoMj: numberOrZero(map.vlsfoMj),
    penRate: numberOrZero(map.penRate),
    rfnboWindow: normalizeText(map.rfnboWindow),
    gwpCo2: numberOrZero(map.gwpCo2 || 1),
    gwpCh4: numberOrZero(map.gwpCh4),
    gwpN2o: numberOrZero(map.gwpN2o),
    gwpBasis: normalizeText(map.gwpBasis),
    penMultiplier: numberOrZero(map.penMultiplier || 1),
    elecWtw: numberOrZero(map.elecWtw),
    gwpCh4Ets: numberOrZero(map.gwpCh4Ets),
    gwpN2oEts: numberOrZero(map.gwpN2oEts),
  };

  const rows = parameterRows.map((row) => {
    if (row.key === "etsPhaseIn") return { ...row, value: normalized.etsPhaseIn };
    if (row.key === "etsGasScope") return { ...row, value: normalized.etsGasScope };
    if (row.key === "fueleuRedux") return { ...row, value: normalized.fueleuRedux };
    if (row.key === "fueleuTarget") return { ...row, value: normalized.fueleuTarget };
    if (row.key === "bioZero") return { ...row, value: normalized.bioZero };
    return { ...row };
  });

  return { rows, values: normalized };
}

function deriveFuelReferenceRows(fuelReference, params) {
  return fuelReference.map((row) => {
    const nonSlipCfCo2 = (1 - row.cslipPercent / 100) * row.cfCo2PerG;
    const ttwCo2eqPerG =
      (1 - row.cslipPercent / 100) * (row.cfCo2PerG + row.cfCh4PerG * params.gwpCh4 + row.cfN2oPerG * params.gwpN2o) +
      (row.cslipPercent / 100) * params.gwpCh4;
    const wtwIntensity = row.lcvMjPerG === 0 ? 0 : row.wtWPerMj + ttwCo2eqPerG / row.lcvMjPerG;
    const etsTtwAr5 =
      (1 - row.cslipPercent / 100) * (row.cfCo2PerG + row.cfCh4PerG * params.gwpCh4Ets + row.cfN2oPerG * params.gwpN2oEts) +
      (row.cslipPercent / 100) * params.gwpCh4Ets;
    return {
      ...row,
      ttwCo2eqPerG: round(ttwCo2eqPerG, 6),
      wtwIntensity: round(wtwIntensity, 6),
      etsTtwAr5: round(etsTtwAr5, 6),
      etsNonCo2Ar5: round(etsTtwAr5 - nonSlipCfCo2, 6),
    };
  });
}

function buildReferenceMaps(state, params) {
  const fuelRows = deriveFuelReferenceRows(state.fuelReference, params);
  const fuelByName = new Map();
  for (const row of fuelRows) {
    fuelByName.set(normalizeUpper(row.fuelPathway), row);
    if (row.alias) {
      fuelByName.set(normalizeUpper(row.alias), row);
    }
  }

  const portByCode = new Map(state.ports.map((row) => [normalizeUpper(row.unlocode), row]));
  const fleetByImo = new Map(state.fleet.map((row) => [String(row.imoNo), row]));

  return { fuelRows, fuelByName, portByCode, fleetByImo };
}

function getFuel(referenceMaps, fuelName) {
  return referenceMaps.fuelByName.get(normalizeUpper(fuelName)) || referenceMaps.fuelByName.get("(NONE)");
}

function routeLabel(fromPort, toPort) {
  if (!fromPort && !toPort) return "Awaiting route";
  return [fromPort || "Unknown", toPort || "Unknown"].join(" to ");
}

function buildScopeNote(type, scope, fromPort, toPort, omrInvolved) {
  if (!type) return "";
  if (scope === 0) return "Out of scope - no EU or EEA port";
  if (type === "Port Stay") return "At-berth in EU or EEA - 100%";
  if (omrInvolved) return "Outermost-region leg - 50%";
  if (scope === 0.5) return "International leg with one EU or EEA port - 50%";
  if (scope === 1 && fromPort && toPort) return "Intra-EU or EEA voyage - 100%";
  return "";
}

function calculateCalculatorRows(state, params, referenceMaps) {
  const counters = { Voyage: 0, "Port Stay": 0 };

  return state.calculatorRows.map((row) => {
    const imoKey = row.imoNo ? String(Math.trunc(row.imoNo)) : "";
    const vessel = referenceMaps.fleetByImo.get(imoKey);
    const fromPort = referenceMaps.portByCode.get(normalizeUpper(row.fromPortCode));
    const toPort = referenceMaps.portByCode.get(normalizeUpper(row.toPortCode));
    const derivedType =
      !row.fromPortCode && !row.toPortCode
        ? ""
        : normalizeUpper(row.fromPortCode) === normalizeUpper(row.toPortCode)
          ? "Port Stay"
          : "Voyage";

    const type = row.type === "Voyage" || row.type === "Port Stay" ? row.type : derivedType;

    const generatedRecordId = type && row.imoNo
      ? `${type === "Port Stay" ? "P" : "V"}${String(++counters[type]).padStart(3, "0")}`
      : "";
    const recordId = normalizeUpper(row.recordId) || generatedRecordId;

    const fromEu = fromPort?.euEeaInScope === "Yes";
    const toEu = toPort?.euEeaInScope === "Yes";
    const omrInvolved = fromPort?.outermostRegion === "Yes" || toPort?.outermostRegion === "Yes";
    const scope =
      !recordId
        ? 0
        : type === "Port Stay"
          ? (fromEu ? 1 : 0)
          : omrInvolved
            ? 0.5
            : fromEu && toEu
              ? 1
              : fromEu || toEu
                ? 0.5
                : 0;

    const fuel1 = getFuel(referenceMaps, row.fuel1Type);
    const fuel2 = getFuel(referenceMaps, row.fuel2Type);
    const bioFuel = getFuel(referenceMaps, row.bioFuelType);

    const fuel1Mt = numberOrZero(row.fuel1ConsumptionMt);
    const fuel2Mt = numberOrZero(row.fuel2ConsumptionMt);
    const bioMt = numberOrZero(row.bioFuelConsumptionMt);
    const windFactor = row.windFactor === null ? 1 : numberOrZero(row.windFactor) || 1;

    const fuel1EnergyMj = fuel1Mt * 1_000_000 * numberOrZero(fuel1?.lcvMjPerG);
    const fuel2EnergyMj = fuel2Mt * 1_000_000 * numberOrZero(fuel2?.lcvMjPerG);
    const bioEnergyMj = bioMt * 1_000_000 * numberOrZero(bioFuel?.lcvMjPerG);
    const totalEnergyMj = fuel1EnergyMj + fuel2EnergyMj + bioEnergyMj;
    const inScopeEnergyMj = totalEnergyMj * scope;
    const opsElectricityInScopeMj = numberOrZero(row.opsElectricityMj) * scope;

    const transportWork = type === "Voyage" ? numberOrZero(row.distanceNm) * numberOrZero(row.cargoTonnes) : 0;

    const fuel1TtwG = fuel1Mt * 1_000_000 * numberOrZero(fuel1?.ttwCo2eqPerG);
    const fuel2TtwG = fuel2Mt * 1_000_000 * numberOrZero(fuel2?.ttwCo2eqPerG);
    const bioTtwG = bioMt * 1_000_000 * numberOrZero(bioFuel?.ttwCo2eqPerG);

    const fuel1WtwG = fuel1EnergyMj * numberOrZero(fuel1?.wtWPerMj) + fuel1TtwG;
    const fuel2WtwG = fuel2EnergyMj * numberOrZero(fuel2?.wtWPerMj) + fuel2TtwG;
    const bioEffectiveWtw =
      row.bioFuelType === "(none)"
        ? 0
        : row.sustainabilityFactor === null
          ? numberOrZero(bioFuel?.wtwIntensity)
          : numberOrZero(row.sustainabilityFactor) * numberOrZero(bioFuel?.wtwIntensity) +
            (1 - numberOrZero(row.sustainabilityFactor)) * params.fueleuRef;
    const bioWtwG = bioEnergyMj * bioEffectiveWtw;

    const elecWtwG = opsElectricityInScopeMj * params.elecWtw;

    const etsCo2eqTotalT =
      params.reportYear >= 2026
        ? (
            fuel1Mt * 1_000_000 * numberOrZero(fuel1?.etsTtwAr5) +
            fuel2Mt * 1_000_000 * numberOrZero(fuel2?.etsTtwAr5) +
            bioMt *
              1_000_000 *
              (params.bioZero === "Yes" ? numberOrZero(bioFuel?.etsNonCo2Ar5) : numberOrZero(bioFuel?.etsTtwAr5))
          ) /
          1_000_000
        : fuel1Mt * numberOrZero(fuel1?.etsCo2Cf) + fuel2Mt * numberOrZero(fuel2?.etsCo2Cf) + bioMt * numberOrZero(bioFuel?.etsCo2Cf);

    const fossilCo2InScopeT = (fuel1Mt * numberOrZero(fuel1?.etsCo2Cf) + fuel2Mt * numberOrZero(fuel2?.etsCo2Cf)) * scope;
    const etsInScopeCo2eqT = params.reportYear >= 2026 ? etsCo2eqTotalT * scope : fossilCo2InScopeT;
    const euasRequiredT = etsInScopeCo2eqT * params.etsPhaseIn;
    const euasCostEur = euasRequiredT * params.euaPrice;

    const fuelEuWtwEmissionsG = (fuel1WtwG + fuel2WtwG + bioWtwG) * windFactor * scope + elecWtwG;
    const fuelEuDenomStep1Mj =
      (fuel1EnergyMj * numberOrZero(fuel1?.rwd) + fuel2EnergyMj * numberOrZero(fuel2?.rwd) + bioEnergyMj * numberOrZero(bioFuel?.rwd)) * scope +
      opsElectricityInScopeMj;
    const fuelEuEnergyStep2Mj = inScopeEnergyMj + opsElectricityInScopeMj;
    const attainedGhgIntensity = fuelEuDenomStep1Mj > 0 ? fuelEuWtwEmissionsG / fuelEuDenomStep1Mj : null;
    const targetGhgIntensity = fuelEuDenomStep1Mj > 0 ? params.fueleuTarget : null;
    const complianceBalanceT =
      fuelEuDenomStep1Mj > 0 && attainedGhgIntensity !== null
        ? (params.fueleuTarget - attainedGhgIntensity) * fuelEuEnergyStep2Mj / 1_000_000
        : null;
    const fuelEuPenaltyEur =
      attainedGhgIntensity && complianceBalanceT !== null && complianceBalanceT < 0
        ? Math.abs(complianceBalanceT) * 1_000_000 / (attainedGhgIntensity * params.vlsfoMj) * params.penRate * params.penMultiplier
        : 0;
    const eeoi =
      type === "Voyage" && numberOrZero(row.cargoTonnes) > 0 && numberOrZero(row.distanceNm) > 0
        ? ((fuel1Mt * numberOrZero(fuel1?.cfCo2PerG) + fuel2Mt * numberOrZero(fuel2?.cfCo2PerG) + bioMt * numberOrZero(bioFuel?.cfCo2PerG)) * 1_000_000) /
          (numberOrZero(row.cargoTonnes) * numberOrZero(row.distanceNm))
        : null;

    return {
      ...row,
      recordId,
      type,
      vesselName: vessel?.vesselName || (row.imoNo ? "IMO not in fleet" : ""),
      shipType: vessel?.shipType || "",
      flagState: vessel?.flag || "",
      deadweightTonnes: vessel?.summerDwt || 0,
      netTonnage: vessel?.nt || 0,
      grossTonnage: vessel?.gt || 0,
      fromPortName: fromPort?.portName || (row.fromPortCode ? "Code not found" : ""),
      toPortName: toPort?.portName || (row.toPortCode ? "Code not found" : ""),
      fromEuEea: fromPort?.euEeaInScope || "",
      toEuEea: toPort?.euEeaInScope || "",
      omrInvolved: omrInvolved ? "Yes" : "No",
      scopePercent: scope,
      scopeNote: buildScopeNote(type, scope, row.fromPortCode, row.toPortCode, omrInvolved),
      route: routeLabel(fromPort?.portName, toPort?.portName),
      totalEnergyMj,
      inScopeEnergyMj,
      transportWork,
      etsCo2eqTotalT,
      etsInScopeCo2eqT,
      euasRequiredT,
      euasCostEur,
      fuelEuWtwEmissionsG,
      fuelEuDenomStep1Mj,
      fuelEuEnergyStep2Mj,
      attainedGhgIntensity,
      targetGhgIntensity,
      complianceBalanceT,
      fuelEuPenaltyEur,
      eeoi,
      fuel1EnergyMj,
      fuel2EnergyMj,
      bioEnergyMj,
      opsElectricityInScopeMj,
      elecWtwG,
      fuel1TtwG,
      fuel2TtwG,
      bioTtwG,
      fuel1WtwG,
      fuel2WtwG,
      bioEffectiveWtw,
      bioWtwG,
      fossilCo2InScopeT,
      fromSpecial: fromPort?.specialCategory || "",
      toSpecial: toPort?.specialCategory || "",
    };
  });
}

function calculateVesselSummary(state, calculatorRows) {
  return state.fleet.map((vessel) => {
    const rows = calculatorRows.filter((row) => String(row.imoNo || "") === String(vessel.imoNo));
    const totalEuasRequired = rows.reduce((sum, row) => sum + numberOrZero(row.euasRequiredT), 0);
    const totalEuasCost = rows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0);
    const totalComplianceBalance = rows.reduce((sum, row) => sum + numberOrZero(row.complianceBalanceT), 0);
    const voyageCount = rows.filter((row) => row.type === "Voyage").length;
    return {
      id: `summary-${vessel.imoNo}`,
      imoNo: vessel.imoNo,
      vesselName: vessel.vesselName,
      shipType: vessel.shipType,
      flag: vessel.flag,
      gt: vessel.gt,
      nt: vessel.nt,
      summerDwt: vessel.summerDwt,
      voyageCount,
      totalEuasRequired,
      totalEuasCost,
      totalComplianceBalance,
      status:
        totalComplianceBalance > 0
          ? "Surplus"
          : totalComplianceBalance < 0
            ? "Deficit"
            : voyageCount > 0
              ? "Neutral"
              : "No activity",
    };
  });
}

function calculateDashboard(params, calculatorRows, vesselSummary) {
  const activeRows = calculatorRows.filter((row) => row.recordId);
  const totalEuasRequired = activeRows.reduce((sum, row) => sum + numberOrZero(row.euasRequiredT), 0);
  const totalEuasCost = activeRows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0);
  const totalFuelEuNumerator = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuWtwEmissionsG), 0);
  const totalFuelEuDenominator = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuDenomStep1Mj), 0);
  const totalFuelEuEnergy = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuEnergyStep2Mj), 0);
  const attainedFleetIntensity = totalFuelEuDenominator === 0 ? 0 : totalFuelEuNumerator / totalFuelEuDenominator;
  const fleetComplianceBalance =
    totalFuelEuDenominator === 0 ? 0 : (params.fueleuTarget - attainedFleetIntensity) * totalFuelEuEnergy / 1_000_000;
  const fleetPenalty =
    fleetComplianceBalance < 0 && attainedFleetIntensity > 0
      ? Math.abs(fleetComplianceBalance) * 1_000_000 / (attainedFleetIntensity * params.vlsfoMj) * params.penRate * params.penMultiplier
      : 0;
  const totalExposure = totalEuasCost + fleetPenalty;

  const inScopeEnergyMj = activeRows.reduce((sum, row) => sum + numberOrZero(row.inScopeEnergyMj), 0);
  const voyageCount = activeRows.filter((row) => row.type === "Voyage").length;
  const portStayCount = activeRows.filter((row) => row.type === "Port Stay").length;
  const highPenaltyRows = [...activeRows].sort((a, b) => numberOrZero(b.fuelEuPenaltyEur) - numberOrZero(a.fuelEuPenaltyEur)).slice(0, 8);
  const topExposureVessels = [...vesselSummary]
    .filter((row) => row.voyageCount > 0)
    .sort((a, b) => numberOrZero(b.totalEuasCost) - numberOrZero(a.totalEuasCost))
    .slice(0, 8);
  const fuelMix = activeRows.reduce((acc, row) => {
    const fuels = [
      [row.fuel1Type, row.fuel1EnergyMj],
      [row.fuel2Type, row.fuel2EnergyMj],
      [row.bioFuelType, row.bioEnergyMj],
    ];
    fuels.forEach(([name, value]) => {
      const label = formatEmptyFuel(name);
      if (label === "(none)" || numberOrZero(value) === 0) return;
      acc[label] = (acc[label] || 0) + numberOrZero(value);
    });
    return acc;
  }, {});

  return {
    kpis: [
      {
        label: "EUAs required",
        value: round(totalEuasRequired, 2),
        detail: "Fleet total from Calculator column AM.",
        tone: "risk",
      },
      {
        label: "ETS cost exposure",
        value: round(totalEuasCost, 0),
        detail: `At EUR ${params.euaPrice} per allowance.`,
        tone: "warn",
      },
      {
        label: "Fleet FuelEU balance",
        value: round(fleetComplianceBalance, 3),
        detail: fleetComplianceBalance >= 0 ? "Positive compliance surplus." : "Deficit that triggers a penalty.",
        tone: fleetComplianceBalance >= 0 ? "good" : "risk",
      },
      {
        label: "Total regulatory exposure",
        value: round(totalExposure, 0),
        detail: "EU ETS cost plus FuelEU penalty.",
        tone: "risk",
      },
      {
        label: "Voyages",
        value: voyageCount,
        detail: "Rows classified as voyages.",
        tone: "neutral",
      },
      {
        label: "Port stays",
        value: portStayCount,
        detail: "Rows classified as port stays.",
        tone: "neutral",
      },
      {
        label: "In-scope energy (MJ)",
        value: round(inScopeEnergyMj, 0),
        detail: "Calculator column AI total.",
        tone: "neutral",
      },
      {
        label: "FuelEU attained intensity",
        value: round(attainedFleetIntensity, 3),
        detail: `Target ${round(params.fueleuTarget, 3)} gCO2eq/MJ.`,
        tone: percentTone(Math.max(0, attainedFleetIntensity - params.fueleuTarget), 1, 5),
      },
    ],
    topPenaltyRows: highPenaltyRows,
    topExposureVessels,
    fuelMix,
    totals: {
      totalEuasRequired,
      totalEuasCost,
      fleetComplianceBalance,
      fleetPenalty,
      totalExposure,
      attainedFleetIntensity,
      inScopeEnergyMj,
    },
  };
}

export function recalculateWorkbook(state) {
  const parameterState = deriveParameterValues(state.parameters);
  const referenceMaps = buildReferenceMaps(state, parameterState.values);
  const calculatorRows = calculateCalculatorRows(state, parameterState.values, referenceMaps);
  const vesselSummary = calculateVesselSummary(state, calculatorRows);
  const dashboard = calculateDashboard(parameterState.values, calculatorRows, vesselSummary);

  return {
    parameters: parameterState.rows,
    parameterValues: parameterState.values,
    fuelReference: referenceMaps.fuelRows,
    calculatorRows,
    vesselSummary,
    dashboard,
  };
}

export function persistableState(state) {
  return deepClone({
    meta: state.meta,
    parameters: state.parameters,
    fuelReference: state.fuelReference,
    fleet: state.fleet,
    ports: state.ports,
    flags: state.flags,
    derogations: state.derogations,
    methodology: state.methodology,
    formulaGuide: state.formulaGuide,
    calculatorRows: state.calculatorRows,
  });
}

export function blankCalculatorRow() {
  return {
    id: `calc-${Date.now()}`,
    recordId: "",
    type: "Voyage",
    imoNo: null,
    departureDate: "",
    fromPortCode: "",
    arrivalDate: "",
    toPortCode: "",
    fuel1Type: "(none)",
    fuel1ConsumptionMt: null,
    fuel2Type: "(none)",
    fuel2ConsumptionMt: null,
    bioFuelType: "(none)",
    bioFuelConsumptionMt: null,
    sustainabilityFactor: null,
    windFactor: 1,
    distanceNm: null,
    cargoTonnes: null,
    timeAtSeaHours: null,
    berthHours: null,
    opsElectricityMj: null,
  };
}

export function blankRowForSheet(sheetKey) {
  const columns = SHEET_COLUMNS[sheetKey] || [];
  const row = { id: `${sheetKey}-${Date.now()}` };
  for (const column of columns) {
    row[column] = "";
  }
  if (sheetKey === "parameters") {
    row.editable = true;
    row.type = "text";
  }
  if (sheetKey === "fuelReference") {
    row.fuelPathway = "(none)";
    row.rwd = 1;
  }
  return row;
}
