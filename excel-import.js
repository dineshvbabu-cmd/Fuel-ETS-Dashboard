const ExcelJS = require("exceljs");

const MAX_WORKBOOK_BYTES = 20 * 1024 * 1024;

const PARAMETER_DEFINITIONS = [
  { row: 5, key: "reportYear", section: "EU ETS - Emissions Trading System", editable: true, type: "number" },
  { row: 6, key: "etsPhaseIn", section: "EU ETS - Emissions Trading System", editable: false, type: "number" },
  { row: 7, key: "etsGasScope", section: "EU ETS - Emissions Trading System", editable: false, type: "text" },
  { row: 8, key: "bioZero", section: "EU ETS - Emissions Trading System", editable: true, type: "text" },
  { row: 9, key: "euaPrice", section: "EU ETS - Emissions Trading System", editable: true, type: "number" },
  { row: 12, key: "fueleuRef", section: "FuelEU Maritime", editable: true, type: "number" },
  { row: 13, key: "fueleuRedux", section: "FuelEU Maritime", editable: false, type: "number" },
  { row: 14, key: "fueleuTarget", section: "FuelEU Maritime", editable: false, type: "number" },
  { row: 15, key: "vlsfoMj", section: "FuelEU Maritime", editable: true, type: "number" },
  { row: 16, key: "penRate", section: "FuelEU Maritime", editable: true, type: "number" },
  { row: 17, key: "rfnboWindow", section: "FuelEU Maritime", editable: true, type: "text" },
  { row: 20, key: "gwpCo2", section: "Global Warming Potentials", editable: true, type: "number" },
  { row: 21, key: "gwpCh4", section: "Global Warming Potentials", editable: true, type: "number" },
  { row: 22, key: "gwpN2o", section: "Global Warming Potentials", editable: true, type: "number" },
  { row: 23, key: "gwpBasis", section: "Global Warming Potentials", editable: true, type: "text" },
  { row: 24, key: "penMultiplier", section: "Global Warming Potentials", editable: true, type: "number" },
  { row: 25, key: "elecWtw", section: "Global Warming Potentials", editable: true, type: "number" },
  { row: 26, key: "gwpCh4Ets", section: "Global Warming Potentials", editable: true, type: "number" },
  { row: 27, key: "gwpN2oEts", section: "Global Warming Potentials", editable: true, type: "number" },
];

const SHEET_ALIASES = {
  calculatorRows: ["calculator", "voyageinputs", "voyageinput", "euetsfueleucalculator", "euetsfuelcalculator"],
  parameters: ["parameters", "parameter"],
  fuelReference: ["fuelreference", "fuelreferences"],
  fleet: ["fleetdb", "fleetdatabase", "fleet"],
  ports: ["portdb", "portdatabase", "ports"],
  flags: ["flagstates", "flagstate", "flags"],
  derogations: ["derogations", "derogation"],
  methodology: ["methodology"],
  formulaGuide: ["formulaguide", "formulas"],
};

const SECTION_LABELS = {
  calculatorRows: "Voyage Inputs",
  parameters: "Parameters",
  fuelReference: "Fuel Reference",
  fleet: "Fleet DB",
  ports: "Port DB",
  flags: "Flag States",
  derogations: "Derogations",
  methodology: "Methodology",
  formulaGuide: "Formula Guide",
};

const COLUMN_DEFINITIONS = {
  fuelReference: [
    { key: "fuelPathway", aliases: ["Fuel Pathway"], type: "text" },
    { key: "fuelClass", aliases: ["Class"], type: "text" },
    { key: "lcvMjPerG", aliases: ["LCV (MJ/g)"], type: "number" },
    { key: "wtWPerMj", aliases: ["WtT (gCO2eq/MJ)", "WtT (gCO₂eq/MJ)"], type: "number" },
    { key: "ttwCo2eqPerG", aliases: ["TtW CO2eq (gCO2eq/gFuel)", "TtW CO₂eq (gCO₂eq/gFuel)"], type: "number" },
    { key: "wtwIntensity", aliases: ["WtW intensity (gCO2eq/MJ)", "WtW intensity (gCO₂eq/MJ)"], type: "number" },
    { key: "rwd", aliases: ["RWD (RFNBO reward)"], type: "number" },
    { key: "etsCo2Cf", aliases: ["ETS CO2 Cf (tCO2/t)", "ETS CO₂ Cf (tCO₂/t)"], type: "number" },
    { key: "notes", aliases: ["Notes"], type: "text" },
    { key: "etsTtwAr5", aliases: ["ETS TtW CO2eq AR5 (gCO2eq/gFuel)", "ETS TtW CO₂eq AR5 (gCO₂eq/gFuel)"], type: "number" },
    { key: "etsNonCo2Ar5", aliases: ["ETS non-CO2 TtW AR5 (gCO2eq/gFuel)", "ETS non-CO₂ TtW AR5 (gCO₂eq/gFuel)"], type: "number" },
    { key: "cfCo2PerG", aliases: ["Cf CO2 (g/gFuel)", "Cf CO₂ (g/gFuel)"], type: "number" },
    { key: "cfCh4PerG", aliases: ["Cf CH4 (g/gFuel)", "Cf CH₄ (g/gFuel)"], type: "number" },
    { key: "cfN2oPerG", aliases: ["Cf N2O (g/gFuel)", "Cf N₂O (g/gFuel)"], type: "number" },
    { key: "cslipPercent", aliases: ["Cslip (% mass)"], type: "number" },
    { key: "consumerSource", aliases: ["Fuel consumer / source (Annex II)"], type: "text" },
  ],
  fleet: [
    { key: "imoNo", aliases: ["IMO No."], type: "number" },
    { key: "vesselName", aliases: ["Vessel Name"], type: "text" },
    { key: "shipType", aliases: ["Ship Type"], type: "text" },
    { key: "flag", aliases: ["Flag"], type: "text" },
    { key: "className", aliases: ["Class"], type: "text" },
    { key: "gt", aliases: ["GT (GRT)"], type: "number" },
    { key: "nt", aliases: ["NT (NRT)"], type: "number" },
    { key: "summerDwt", aliases: ["Summer DWT"], type: "number" },
    { key: "built", aliases: ["Built"], type: "number" },
    { key: "wapsFwindFactor", aliases: ["WAPS Fwind factor", "WAPS Fwind"], type: "number" },
  ],
  ports: [
    { key: "unlocode", aliases: ["UN/LOCODE"], type: "text" },
    { key: "portName", aliases: ["Port Name"], type: "text" },
    { key: "country", aliases: ["Country"], type: "text" },
    { key: "countryCode", aliases: ["Country Code"], type: "text" },
    { key: "euEeaInScope", aliases: ["EU/EEA In-Scope"], type: "text" },
    { key: "outermostRegion", aliases: ["Outermost Region"], type: "text" },
    { key: "specialCategory", aliases: ["Special Category", "Small Islands"], type: "text" },
  ],
  flags: [
    { key: "flagState", aliases: ["Flag State"], type: "text" },
    { key: "iso", aliases: ["ISO"], type: "text" },
    { key: "registryType", aliases: ["Registry Type"], type: "text" },
    { key: "euEeaFlag", aliases: ["EU/EEA Flag"], type: "text" },
    { key: "notes", aliases: ["Notes"], type: "text" },
  ],
  derogations: [
    { key: "serialNo", aliases: ["#", "S.No"], type: "number" },
    { key: "derogationRule", aliases: ["Derogation / Rule"], type: "text" },
    { key: "regime", aliases: ["Regime"], type: "text" },
    { key: "whatItCovers", aliases: ["What it covers"], type: "text" },
    { key: "conditionsEligibility", aliases: ["Conditions / Eligibility"], type: "text" },
    { key: "effect", aliases: ["Effect"], type: "text" },
    { key: "expires", aliases: ["Expires"], type: "date" },
    { key: "legalBasis", aliases: ["Legal basis"], type: "text" },
  ],
  formulaGuide: [
    { key: "stepField", aliases: ["Step / Field"], type: "text" },
    { key: "resultColumn", aliases: ["Result column"], type: "text" },
    { key: "formulaPlainEnglish", aliases: ["Formula (plain English)"], type: "text" },
  ],
};

const SUBSCRIPT_MAP = {
  "₀": "0",
  "₁": "1",
  "₂": "2",
  "₃": "3",
  "₄": "4",
  "₅": "5",
  "₆": "6",
  "₇": "7",
  "₈": "8",
  "₉": "9",
};

function normalizeName(value) {
  return String(value ?? "")
    .replace(/[₀-₉]/g, (character) => SUBSCRIPT_MAP[character] || character)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function unwrapCellValue(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value;
  if (typeof value !== "object") return value;
  if (Object.prototype.hasOwnProperty.call(value, "result")) return unwrapCellValue(value.result);
  if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
  if (Object.prototype.hasOwnProperty.call(value, "text")) return value.text;
  if (Object.prototype.hasOwnProperty.call(value, "hyperlink")) return value.text || value.hyperlink;
  return null;
}

function cellValue(worksheet, row, column) {
  if (!worksheet || !row || !column || column < 1) return null;
  return unwrapCellValue(worksheet.getCell(row, column).value);
}

function textValue(value) {
  return String(value ?? "").trim();
}

function numberValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const numeric = Number(String(value).replaceAll(",", ""));
  return Number.isFinite(numeric) ? numeric : null;
}

function dateValue(value) {
  if (!value) return "";
  let date = value;
  if (typeof value === "number" && value > 20000) {
    date = new Date(Date.UTC(1899, 11, 30) + Math.round(value * 86400000));
  } else if (!(value instanceof Date)) {
    const text = String(value).trim();
    const dayFirst = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (dayFirst) {
      date = new Date(Date.UTC(Number(dayFirst[3]), Number(dayFirst[2]) - 1, Number(dayFirst[1])));
    } else {
      date = new Date(text);
    }
  }
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function portableValue(value) {
  const unwrapped = unwrapCellValue(value);
  if (unwrapped === null || unwrapped === undefined || unwrapped === "") return "";
  if (unwrapped instanceof Date) return dateValue(unwrapped);
  if (typeof unwrapped === "number" || typeof unwrapped === "boolean") return unwrapped;
  return textValue(unwrapped);
}

function uniqueColumnKey(baseKey, usedKeys) {
  let key = baseKey || "column";
  let suffix = 2;
  while (usedKeys.has(key)) {
    key = `${baseKey || "column"}_${suffix}`;
    suffix += 1;
  }
  usedKeys.add(key);
  return key;
}

function findMatchingDefinition(headerText, definitions = []) {
  const rawHeader = textValue(headerText);
  const normalizedHeader = normalizeName(headerText);
  return (
    definitions.find((definition) =>
      (definition.aliases || []).some((alias) => {
        const rawAlias = textValue(alias);
        if (rawAlias && rawAlias === rawHeader) return true;
        const normalizedAlias = normalizeName(alias);
        return normalizedAlias && normalizedHeader === normalizedAlias;
      })
    ) || null
  );
}

function inferColumnType(values) {
  const nonEmpty = values.filter((value) => value !== null && value !== undefined && value !== "");
  if (!nonEmpty.length) return "text";
  if (nonEmpty.every((value) => typeof value === "number")) return "number";
  if (nonEmpty.every((value) => typeof value === "boolean")) return "boolean";
  if (nonEmpty.every((value) => typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value))) return "date";
  return "text";
}

function buildDynamicColumns(worksheet, headerRow, sectionKey) {
  const definitions = COLUMN_DEFINITIONS[sectionKey] || [];
  const usedKeys = new Set(["id"]);
  const columns = [];

  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
    const label = textValue(cellValue(worksheet, headerRow, columnNumber));
    if (!label) continue;
    const definition = findMatchingDefinition(label, definitions);
    const key = uniqueColumnKey(definition?.key || normalizeName(label), usedKeys);
    columns.push({
      key,
      label,
      type: definition?.type || "auto",
      sourceColumn: columnNumber,
    });
  }

  return columns;
}

function parseDynamicTable(worksheet, { sectionKey, headerRow, dataStartRow, idPrefix }) {
  const columns = buildDynamicColumns(worksheet, headerRow, sectionKey);
  const rows = [];
  const warnings = [];

  for (let rowNumber = dataStartRow; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = { id: makeId(idPrefix, rows.length) };
    let hasData = false;

    columns.forEach((column) => {
      const rawValue = cellValue(worksheet, rowNumber, column.sourceColumn);
      let value;
      if (column.type === "number") {
        value = numberValue(rawValue);
      } else if (column.type === "date") {
        value = dateValue(rawValue);
      } else {
        value = portableValue(rawValue);
      }
      if (value !== null && value !== undefined && value !== "") {
        hasData = true;
      }
      row[column.key] = value;
    });

    if (!hasData) continue;
    rows.push(row);
  }

  const finalizedColumns = columns.map((column) => ({
    key: column.key,
    label: column.label,
    type: column.type === "auto" ? inferColumnType(rows.map((row) => row[column.key])) : column.type,
  }));

  return { rows, warnings, columns: finalizedColumns };
}

function findHeaderRowByAliases(worksheet, aliases, maxRow = 12) {
  let bestRow = 1;
  let bestScore = -1;
  for (let rowNumber = 1; rowNumber <= Math.min(maxRow, worksheet.rowCount); rowNumber += 1) {
    let score = 0;
    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const label = textValue(cellValue(worksheet, rowNumber, columnNumber));
      if (!label) continue;
      const normalized = normalizeName(label);
      if (aliases.some((alias) => normalized === normalizeName(alias) || normalized.includes(normalizeName(alias)))) {
        score += 1;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = rowNumber;
    }
  }
  return bestRow;
}

function makeId(prefix, index) {
  return `${prefix}-${Date.now()}-${index + 1}`;
}

function findWorksheet(workbook, sectionKey) {
  const aliases = new Set(SHEET_ALIASES[sectionKey]);
  return workbook.worksheets.find((worksheet) => aliases.has(normalizeName(worksheet.name))) || null;
}

function headerMap(worksheet, rowNumber) {
  const result = new Map();
  for (let column = 1; column <= worksheet.columnCount; column += 1) {
    const normalized = normalizeName(cellValue(worksheet, rowNumber, column));
    if (normalized && !result.has(normalized)) result.set(normalized, column);
  }
  return result;
}

function findColumn(headers, aliases, fallback) {
  for (const alias of aliases) {
    const exact = headers.get(normalizeName(alias));
    if (exact) return exact;
  }
  for (const [header, column] of headers.entries()) {
    if (aliases.some((alias) => header.includes(normalizeName(alias)))) return column;
  }
  return fallback;
}

function parseCalculator(worksheet) {
  const headers = headerMap(worksheet, 4);
  const columns = {
    recordId: findColumn(headers, ["Voyage / Port-Stay ID"], 1),
    type: findColumn(headers, ["Type"], 2),
    imoNo: findColumn(headers, ["IMO No."], 3),
    departureDate: findColumn(headers, ["Departure Date"], 10),
    fromPortCode: findColumn(headers, ["From Port UN/LOCODE"], 11),
    arrivalDate: findColumn(headers, ["Arrival Date"], 13),
    toPortCode: findColumn(headers, ["To Port UN/LOCODE"], 14),
    fuel1Type: findColumn(headers, ["Fossil Fuel 1 Type"], 16),
    fuel1ConsumptionMt: findColumn(headers, ["Fossil Fuel 1 Cons. (MT)"], 17),
    fuel2Type: findColumn(headers, ["Fossil Fuel 2 Type"], 18),
    fuel2ConsumptionMt: findColumn(headers, ["Fossil Fuel 2 Cons. (MT)"], 19),
    bioFuelType: findColumn(headers, ["Biofuel/RFNBO Type"], 20),
    bioFuelConsumptionMt: findColumn(headers, ["Biofuel/RFNBO Cons. (MT)"], 21),
    sustainabilityFactor: findColumn(headers, ["Sustain. Factor WtW (0-1)"], 22),
    windFactor: findColumn(headers, ["WASP Factor (f_wind)"], 23),
    distanceNm: findColumn(headers, ["Total Distance (nm)"], null),
    cargoTonnes: findColumn(headers, ["Cargo Carried (t)"], null),
    timeAtSeaHours: findColumn(headers, ["Time at Sea (h)"], null),
    berthHours: findColumn(headers, ["Hours at Berth/Anchor"], null),
    opsElectricityMj: findColumn(headers, ["OPS Electricity (MJ)"], 24),
  };
  const rows = [];
  const warnings = [];

  for (let rowNumber = 5; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = {
      id: makeId("calc-import", rows.length),
      recordId: textValue(cellValue(worksheet, rowNumber, columns.recordId)).toUpperCase(),
      type: textValue(cellValue(worksheet, rowNumber, columns.type)),
      imoNo: numberValue(cellValue(worksheet, rowNumber, columns.imoNo)),
      departureDate: dateValue(cellValue(worksheet, rowNumber, columns.departureDate)),
      fromPortCode: textValue(cellValue(worksheet, rowNumber, columns.fromPortCode)).toUpperCase(),
      arrivalDate: dateValue(cellValue(worksheet, rowNumber, columns.arrivalDate)),
      toPortCode: textValue(cellValue(worksheet, rowNumber, columns.toPortCode)).toUpperCase(),
      fuel1Type: textValue(cellValue(worksheet, rowNumber, columns.fuel1Type)) || "(none)",
      fuel1ConsumptionMt: numberValue(cellValue(worksheet, rowNumber, columns.fuel1ConsumptionMt)),
      fuel2Type: textValue(cellValue(worksheet, rowNumber, columns.fuel2Type)) || "(none)",
      fuel2ConsumptionMt: numberValue(cellValue(worksheet, rowNumber, columns.fuel2ConsumptionMt)),
      bioFuelType: textValue(cellValue(worksheet, rowNumber, columns.bioFuelType)) || "(none)",
      bioFuelConsumptionMt: numberValue(cellValue(worksheet, rowNumber, columns.bioFuelConsumptionMt)),
      sustainabilityFactor: numberValue(cellValue(worksheet, rowNumber, columns.sustainabilityFactor)),
      windFactor: numberValue(cellValue(worksheet, rowNumber, columns.windFactor)),
      distanceNm: numberValue(cellValue(worksheet, rowNumber, columns.distanceNm)),
      cargoTonnes: numberValue(cellValue(worksheet, rowNumber, columns.cargoTonnes)),
      timeAtSeaHours: numberValue(cellValue(worksheet, rowNumber, columns.timeAtSeaHours)),
      berthHours: numberValue(cellValue(worksheet, rowNumber, columns.berthHours)),
      opsElectricityMj: numberValue(cellValue(worksheet, rowNumber, columns.opsElectricityMj)),
      entrySource: "excel",
      sourceSystem: "Excel import",
      sourceRecordId: `Calculator!${rowNumber}`,
      sourceUpdatedAt: new Date().toISOString(),
    };

    const hasManualInput = [
      row.imoNo,
      row.departureDate,
      row.fromPortCode,
      row.arrivalDate,
      row.toPortCode,
      row.fuel1ConsumptionMt,
      row.fuel2ConsumptionMt,
      row.bioFuelConsumptionMt,
      row.distanceNm,
      row.cargoTonnes,
      row.timeAtSeaHours,
      row.berthHours,
      row.opsElectricityMj,
    ].some((value) => value !== null && value !== "");
    if (!hasManualInput) continue;

    if (row.type !== "Voyage" && row.type !== "Port Stay") {
      row.type = row.fromPortCode && row.toPortCode && row.fromPortCode === row.toPortCode ? "Port Stay" : "Voyage";
    }
    row.storageYear =
      Number(row.departureDate.slice(0, 4)) ||
      Number(row.arrivalDate.slice(0, 4)) ||
      new Date().getFullYear();
    if (!row.imoNo) warnings.push(`Calculator row ${rowNumber} has no IMO number.`);
    rows.push(row);
  }

  return { rows, warnings };
}

function parseParameters(worksheet) {
  const rows = PARAMETER_DEFINITIONS.map((definition, index) => ({
    id: makeId("parameter-import", index),
    section: definition.section,
    key: definition.key,
    label: textValue(cellValue(worksheet, definition.row, 1)),
    value: unwrapCellValue(worksheet.getCell(definition.row, 2).value),
    note: textValue(cellValue(worksheet, definition.row, 3)),
    editable: definition.editable,
    type: definition.type,
  })).filter((row) => row.label || row.value !== null);
  return {
    rows,
    warnings: [],
    columns: [
      { key: "section", label: "Section", type: "text" },
      { key: "key", label: "Key", type: "text" },
      { key: "label", label: "Label", type: "text" },
      { key: "value", label: "Value", type: "text" },
      { key: "note", label: "Note", type: "text" },
      { key: "editable", label: "Editable", type: "boolean" },
      { key: "type", label: "Type", type: "text" },
    ],
  };
}

function parseFuelReference(worksheet) {
  return parseDynamicTable(worksheet, {
    sectionKey: "fuelReference",
    headerRow: 4,
    dataStartRow: 5,
    idPrefix: "fuel-import",
  });
}

function parseFleet(worksheet) {
  return parseDynamicTable(worksheet, {
    sectionKey: "fleet",
    headerRow: 4,
    dataStartRow: 5,
    idPrefix: "fleet-import",
  });
}

function parsePorts(worksheet) {
  return parseDynamicTable(worksheet, {
    sectionKey: "ports",
    headerRow: 4,
    dataStartRow: 5,
    idPrefix: "port-import",
  });
}

function parseFlags(worksheet) {
  return parseDynamicTable(worksheet, {
    sectionKey: "flags",
    headerRow: 4,
    dataStartRow: 5,
    idPrefix: "flag-import",
  });
}

function parseDerogations(worksheet) {
  const headerRow = findHeaderRowByAliases(worksheet, ["Derogation / Rule", "Regime", "Effect", "Legal basis"]);
  return parseDynamicTable(worksheet, {
    sectionKey: "derogations",
    headerRow,
    dataStartRow: headerRow + 1,
    idPrefix: "derogation-import",
  });
}

function parseMethodology(worksheet) {
  const rows = [];
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const detail = textValue(cellValue(worksheet, rowNumber, 2));
    if (!detail) continue;
    rows.push({ id: makeId("methodology-import", rows.length), detail });
  }
  return {
    rows,
    warnings: [],
    columns: [{ key: "detail", label: "Detail", type: "text" }],
  };
}

function parseFormulaGuide(worksheet) {
  return parseDynamicTable(worksheet, {
    sectionKey: "formulaGuide",
    headerRow: 3,
    dataStartRow: 4,
    idPrefix: "formula-import",
  });
}

const PARSERS = {
  calculatorRows: parseCalculator,
  parameters: parseParameters,
  fuelReference: parseFuelReference,
  fleet: parseFleet,
  ports: parsePorts,
  flags: parseFlags,
  derogations: parseDerogations,
  methodology: parseMethodology,
  formulaGuide: parseFormulaGuide,
};

async function parseComplianceWorkbook(buffer, fileName = "workbook.xlsx") {
  if (!Buffer.isBuffer(buffer) || !buffer.length) throw new Error("The uploaded workbook is empty.");
  if (buffer.length > MAX_WORKBOOK_BYTES) throw new Error("The workbook exceeds the 20 MB upload limit.");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer, {
    ignoreNodes: ["dataValidations", "conditionalFormatting", "extLst"],
  });
  const totalCells = workbook.worksheets.reduce(
    (total, worksheet) => total + worksheet.rowCount * Math.max(worksheet.columnCount, 1),
    0
  );
  if (workbook.worksheets.length > 30 || totalCells > 1_000_000) {
    throw new Error("The workbook is larger than the supported compliance template.");
  }

  const sections = {};
  const warnings = [];
  for (const sectionKey of Object.keys(PARSERS)) {
    const worksheet = findWorksheet(workbook, sectionKey);
    if (!worksheet) continue;
    const parsed = PARSERS[sectionKey](worksheet);
    if (!parsed.rows.length) {
      warnings.push(`${worksheet.name} was found but contained no importable rows.`);
      continue;
    }
    sections[sectionKey] = {
      key: sectionKey,
      label: SECTION_LABELS[sectionKey],
      sourceSheet: worksheet.name,
      rowCount: parsed.rows.length,
      columns: parsed.columns || [],
      rows: parsed.rows,
      warnings: parsed.warnings,
    };
    warnings.push(...parsed.warnings);
  }

  if (!Object.keys(sections).length) {
    throw new Error("No supported sheets were found. Use the Fuel ETS workbook format with Calculator and/or reference-library sheets.");
  }

  return {
    fileName: pathSafeName(fileName),
    importedAt: new Date().toISOString(),
    workbookSheets: workbook.worksheets.map((worksheet) => worksheet.name),
    sections,
    warnings,
  };
}

function pathSafeName(fileName) {
  return String(fileName || "workbook.xlsx").replace(/[^\w.\- ()]/g, "_").slice(0, 160);
}

module.exports = {
  MAX_WORKBOOK_BYTES,
  parseComplianceWorkbook,
};
