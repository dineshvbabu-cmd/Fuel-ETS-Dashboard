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
  return { rows, warnings: [] };
}

function parseFuelReference(worksheet) {
  const rows = [];
  for (let rowNumber = 5; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const fuelPathway = textValue(cellValue(worksheet, rowNumber, 1));
    if (!fuelPathway) continue;
    rows.push({
      id: makeId("fuel-import", rows.length),
      fuelPathway,
      fuelClass: textValue(cellValue(worksheet, rowNumber, 2)),
      lcvMjPerG: numberValue(cellValue(worksheet, rowNumber, 3)) || 0,
      wtWPerMj: numberValue(cellValue(worksheet, rowNumber, 4)) || 0,
      rwd: numberValue(cellValue(worksheet, rowNumber, 7)) || 1,
      etsCo2Cf: numberValue(cellValue(worksheet, rowNumber, 8)) || 0,
      notes: textValue(cellValue(worksheet, rowNumber, 9)),
      alias: textValue(cellValue(worksheet, rowNumber, 11)) || "(none)",
      cfCo2PerG: numberValue(cellValue(worksheet, rowNumber, 13)) || 0,
      cfCh4PerG: numberValue(cellValue(worksheet, rowNumber, 14)) || 0,
      cfN2oPerG: numberValue(cellValue(worksheet, rowNumber, 15)) || 0,
      cslipPercent: numberValue(cellValue(worksheet, rowNumber, 16)) || 0,
      consumerSource: textValue(cellValue(worksheet, rowNumber, 17)),
    });
  }
  return { rows, warnings: [] };
}

function parseFleet(worksheet) {
  const headers = headerMap(worksheet, 4);
  const wapsColumn = findColumn(headers, ["WAPS Fwind factor", "WASP Factor"], 10);
  const rows = [];
  for (let rowNumber = 5; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const imoNo = numberValue(cellValue(worksheet, rowNumber, 1));
    const vesselName = textValue(cellValue(worksheet, rowNumber, 2));
    if (!imoNo && !vesselName) continue;
    rows.push({
      id: makeId("fleet-import", rows.length),
      imoNo,
      vesselName,
      shipType: textValue(cellValue(worksheet, rowNumber, 3)),
      flag: textValue(cellValue(worksheet, rowNumber, 4)),
      className: textValue(cellValue(worksheet, rowNumber, 5)),
      gt: numberValue(cellValue(worksheet, rowNumber, 6)) || 0,
      nt: numberValue(cellValue(worksheet, rowNumber, 7)) || 0,
      summerDwt: numberValue(cellValue(worksheet, rowNumber, 8)) || 0,
      built: numberValue(cellValue(worksheet, rowNumber, 9)),
      wapsFwindFactor: numberValue(cellValue(worksheet, rowNumber, wapsColumn)),
    });
  }
  return { rows, warnings: [] };
}

function parsePorts(worksheet) {
  const rows = [];
  for (let rowNumber = 5; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const unlocode = textValue(cellValue(worksheet, rowNumber, 1)).toUpperCase();
    if (!unlocode) continue;
    rows.push({
      id: makeId("port-import", rows.length),
      unlocode,
      portName: textValue(cellValue(worksheet, rowNumber, 2)),
      country: textValue(cellValue(worksheet, rowNumber, 3)),
      countryCode: textValue(cellValue(worksheet, rowNumber, 4)),
      euEeaInScope: textValue(cellValue(worksheet, rowNumber, 5)),
      outermostRegion: textValue(cellValue(worksheet, rowNumber, 6)),
      specialCategory: textValue(cellValue(worksheet, rowNumber, 7)),
    });
  }
  return { rows, warnings: [] };
}

function parseFlags(worksheet) {
  const rows = [];
  for (let rowNumber = 5; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const flagState = textValue(cellValue(worksheet, rowNumber, 1));
    if (!flagState) continue;
    rows.push({
      id: makeId("flag-import", rows.length),
      flagState,
      iso: textValue(cellValue(worksheet, rowNumber, 2)),
      registryType: textValue(cellValue(worksheet, rowNumber, 3)),
      euEeaFlag: textValue(cellValue(worksheet, rowNumber, 4)),
      notes: textValue(cellValue(worksheet, rowNumber, 5)),
    });
  }
  return { rows, warnings: [] };
}

function parseDerogations(worksheet) {
  const rows = [];
  for (let rowNumber = 8; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const unlocode = textValue(cellValue(worksheet, rowNumber, 5)).toUpperCase();
    const region = textValue(cellValue(worksheet, rowNumber, 3));
    if (!unlocode && !region) continue;
    rows.push({
      id: makeId("derogation-import", rows.length),
      serialNo: numberValue(cellValue(worksheet, rowNumber, 1)) || rows.length + 1,
      euMemberState: textValue(cellValue(worksheet, rowNumber, 2)),
      outermostRegion: region,
      omrPortName: textValue(cellValue(worksheet, rowNumber, 4)),
      unlocode,
    });
  }
  return { rows, warnings: [] };
}

function parseMethodology(worksheet) {
  const rows = [];
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const detail = textValue(cellValue(worksheet, rowNumber, 2));
    if (!detail) continue;
    rows.push({ id: makeId("methodology-import", rows.length), detail });
  }
  return { rows, warnings: [] };
}

function parseFormulaGuide(worksheet) {
  const rows = [];
  for (let rowNumber = 4; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const stepField = textValue(cellValue(worksheet, rowNumber, 1));
    if (!stepField) continue;
    rows.push({
      id: makeId("formula-import", rows.length),
      stepField,
      resultColumn: textValue(cellValue(worksheet, rowNumber, 2)),
      formulaPlainEnglish: textValue(cellValue(worksheet, rowNumber, 3)),
    });
  }
  return { rows, warnings: [] };
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
