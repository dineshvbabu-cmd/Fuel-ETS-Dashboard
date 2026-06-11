const PDFDocument = require("pdfkit");

const COLORS = {
  navy: "#073E49",
  cyan: "#1C88AF",
  ink: "#143744",
  muted: "#536C7A",
  pale: "#E9F1FC",
  line: "#BCCBD5",
  white: "#FFFFFF",
  good: "#187A3E",
  risk: "#B83A32",
};

function number(value, digits = 2) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return "-";
  return new Intl.NumberFormat("en-GB", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(numeric);
}

function integer(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return "-";
  return new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 }).format(numeric);
}

function dateLabel(value) {
  if (!value) return "-";
  const parsed = new Date(`${String(value).slice(0, 10)}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime())) return String(value);
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    timeZone: "UTC",
  }).format(parsed);
}

function safeText(value, fallback = "-") {
  const text = String(value ?? "").trim();
  return text || fallback;
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
}

function tableCell(doc, text, x, y, width, height, options = {}) {
  const {
    fill = null,
    color = COLORS.ink,
    bold = false,
    align = "left",
    size = 8,
    padding = 8,
  } = options;
  if (fill) {
    doc.rect(x, y, width, height).fill(fill);
  }
  doc.rect(x, y, width, height).strokeColor(COLORS.line).lineWidth(0.5).stroke();
  doc
    .font(bold ? "Helvetica-Bold" : "Helvetica")
    .fontSize(size)
    .fillColor(color)
    .text(safeText(text), x + padding, y + (height - size) / 2 - 1, {
      width: width - padding * 2,
      align,
      lineBreak: false,
      ellipsis: true,
    });
}

function drawStatementHeader(doc, reportYear) {
  const x = 51;
  const y = 51;
  const width = 493;
  doc.rect(x, y, width, 58).fill(COLORS.navy);
  doc.rect(x, y + 54, width, 4).fill(COLORS.cyan);
  doc.font("Helvetica-Bold").fontSize(17).fillColor(COLORS.white).text("ATLANTAS SHIP MANAGEMENT", x + 35, y + 11);
  doc.font("Helvetica").fontSize(8).fillColor("#CFE4E9").text("EU ETS & FuelEU Maritime  -  Vessel Compliance Statement", x + 35, y + 36);
  doc.font("Helvetica-Bold").fontSize(8).fillColor(COLORS.white).text("VESSEL STATEMENT", x + 380, y + 15, { width: 95, align: "right" });
  doc.font("Helvetica").fontSize(8).fillColor("#CFE4E9").text(`Reporting Period ${reportYear}`, x + 380, y + 35, { width: 95, align: "right" });
}

function drawVesselStatementPage(doc, payload) {
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const vessel = payload.vessel || {};
  const reportYear = Number(payload.reportYear) || new Date().getFullYear();
  const issued = payload.issuedAt ? new Date(payload.issuedAt) : new Date();
  const issuedLabel = new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "short", year: "numeric" }).format(issued);
  const imo = safeText(vessel.imoNo, "UNKNOWN");
  const reference = `ASM-CS-${reportYear}-${imo}`;
  const euaPrice = Number(payload.euaPrice) || 0;
  const totalEuas = sum(rows, "euasRequiredT");
  const totalCost = sum(rows, "euasCostEur");
  const balance = sum(rows, "complianceBalanceT");
  const penalty = sum(rows, "fuelEuPenaltyEur");
  const status = balance > 0 ? "Surplus" : balance < 0 ? "Deficit" : rows.length ? "Neutral" : "No activity";

  drawStatementHeader(doc, reportYear);
  doc.font("Helvetica-Bold").fontSize(16).fillColor(COLORS.navy).text("Vessel Compliance Statement", 85, 127);
  doc
    .font("Helvetica")
    .fontSize(8)
    .fillColor(COLORS.muted)
    .text(`Reporting year ${reportYear}    |    Issued ${issuedLabel}    |    Ref: ${reference}`, 85, 154);

  const left = 72;
  const top = 180;
  const labelWidth = 96;
  const valueWidth = 143;
  const rightLabelWidth = 96;
  const rightValueWidth = 153;
  const rowHeight = 26;
  const vesselRows = [
    ["VESSEL NAME", vessel.vesselName, "IMO NUMBER", imo],
    ["SHIP TYPE", vessel.shipType, "FLAG", vessel.flag],
    ["CLASS", vessel.className, "GROSS TONNAGE", integer(vessel.gt)],
    ["YEAR BUILT", safeText(vessel.built), "IN-SCOPE VOYAGES", integer(rows.filter((row) => row.type === "Voyage" && Number(row.scopePercent) > 0).length)],
  ];
  vesselRows.forEach((item, index) => {
    const y = top + index * rowHeight;
    tableCell(doc, item[0], left, y, labelWidth, rowHeight, { fill: COLORS.pale, color: COLORS.muted, size: 6.5 });
    tableCell(doc, item[1], left + labelWidth, y, valueWidth, rowHeight, { fill: COLORS.pale, bold: true });
    tableCell(doc, item[2], left + labelWidth + valueWidth, y, rightLabelWidth, rowHeight, { fill: COLORS.pale, color: COLORS.muted, size: 6.5 });
    tableCell(doc, item[3], left + labelWidth + valueWidth + rightLabelWidth, y, rightValueWidth, rowHeight, { fill: COLORS.pale, bold: true });
  });

  const summaryTop = top + vesselRows.length * rowHeight + 39;
  doc.font("Helvetica-Bold").fontSize(9).fillColor(COLORS.navy).text("COMPLIANCE SUMMARY", left + 13, summaryTop - 18);
  tableCell(doc, "Metric", left, summaryTop, 416, 24, { fill: COLORS.navy, color: COLORS.white, bold: true });
  tableCell(doc, "Value", left + 416, summaryTop, 72, 24, { fill: COLORS.navy, color: COLORS.white, bold: true, align: "right" });
  const summaryRows = [
    ["EUAs Required - EU ETS obligation (t CO2e)", number(totalEuas, 2)],
    [`Indicative EUA Cost (at EUR ${number(euaPrice, 2)} / EUA)`, `EUR ${number(totalCost, 0)}`],
    ["FuelEU Compliance Balance (t CO2e)", number(balance, 2)],
    ["FuelEU Penalty", `EUR ${number(penalty, 0)}`],
    ["Compliance Status", status],
  ];
  summaryRows.forEach((item, index) => {
    const y = summaryTop + 24 + index * 26;
    tableCell(doc, item[0], left, y, 416, 26);
    tableCell(doc, item[1], left + 416, y, 72, 26, {
      bold: true,
      align: "right",
      color: item[0] === "Compliance Status" ? (balance < 0 ? COLORS.risk : COLORS.good) : COLORS.ink,
    });
  });

  const notesTop = summaryTop + 24 + summaryRows.length * 26 + 19;
  doc.font("Helvetica-Bold").fontSize(8).fillColor(COLORS.navy).text("Basis & Notes", left + 13, notesTop);
  const notes = [
    "EUAs Required is the EU ETS allowance obligation for in-scope EU/EEA voyages at the applicable phase-in share for the reporting year.",
    `Indicative EUA Cost uses the configured allowance price of EUR ${number(euaPrice, 2)} per EUA and is for guidance only.`,
    "FuelEU Compliance Balance: a positive value is a Surplus, a negative value a Deficit. Figures are indicative and subject to MRV verification.",
    "This statement is not legal, tax or financial advice; cost allocation between owners and charterers is governed by the charter party.",
  ];
  notes.forEach((note, index) => {
    doc.circle(left + 14, notesTop + 20 + index * 34, 1.4).fill(COLORS.muted);
    doc.font("Helvetica").fontSize(7.5).fillColor(COLORS.muted).text(note, left + 23, notesTop + 15 + index * 34, { width: 465, lineGap: 2 });
  });

  doc.font("Helvetica").fontSize(7).fillColor(COLORS.muted).text(`Selection: ${safeText(payload.selectionLabel, "All vessel records")}`, left, 805, { width: 488, align: "right" });
}

function drawVoyageDetails(doc, payload) {
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) return;

  doc.addPage({ size: "A4", layout: "landscape", margin: 32 });
  const pageWidth = doc.page.width;
  doc.rect(32, 32, pageWidth - 64, 42).fill(COLORS.navy);
  doc.font("Helvetica-Bold").fontSize(15).fillColor(COLORS.white).text("Selected Voyage & Port-Stay Detail", 50, 45);
  doc.font("Helvetica").fontSize(8).fillColor("#CFE4E9").text(safeText(payload.selectionLabel), pageWidth - 310, 48, { width: 260, align: "right" });

  const columns = [
    ["ID", 48, (row) => row.recordId],
    ["Type", 58, (row) => row.type],
    ["Departure", 68, (row) => dateLabel(row.departureDate)],
    ["From", 92, (row) => row.fromPortName || row.fromPortCode],
    ["Arrival", 68, (row) => dateLabel(row.arrivalDate)],
    ["To", 92, (row) => row.toPortName || row.toPortCode],
    ["EUAs", 58, (row) => number(row.euasRequiredT, 2)],
    ["EUA Cost", 72, (row) => `EUR ${number(row.euasCostEur, 0)}`],
    ["GHG", 58, (row) => number(row.attainedGhgIntensity, 2)],
    ["Balance", 66, (row) => number(row.complianceBalanceT, 2)],
    ["Penalty", 72, (row) => `EUR ${number(row.fuelEuPenaltyEur, 0)}`],
  ];
  const rowHeight = 25;
  let y = 92;

  function drawHeader() {
    let x = 32;
    columns.forEach(([label, width]) => {
      tableCell(doc, label, x, y, width, rowHeight, { fill: COLORS.navy, color: COLORS.white, bold: true, size: 7, align: "center", padding: 4 });
      x += width;
    });
    y += rowHeight;
  }

  drawHeader();
  rows.forEach((row) => {
    if (y + rowHeight > doc.page.height - 38) {
      doc.addPage({ size: "A4", layout: "landscape", margin: 32 });
      y = 42;
      drawHeader();
    }
    let x = 32;
    columns.forEach(([, width, getter], index) => {
      tableCell(doc, getter(row), x, y, width, rowHeight, {
        size: 6.8,
        align: index >= 6 ? "right" : "center",
        padding: 4,
        fill: y / rowHeight % 2 ? "#F8FAFC" : COLORS.white,
      });
      x += width;
    });
    y += rowHeight;
  });
}

function createComplianceStatement(payload) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", layout: "portrait", margin: 0, info: { Title: "Vessel Compliance Statement" } });
    const chunks = [];
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
    drawVesselStatementPage(doc, payload);
    drawVoyageDetails(doc, payload);
    doc.end();
  });
}

module.exports = {
  createComplianceStatement,
};
