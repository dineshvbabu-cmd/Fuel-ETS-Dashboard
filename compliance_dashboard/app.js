import {
  STORAGE_KEY,
  SHEETS,
  SHEET_COLUMNS,
  CALCULATOR_FIELD_GROUPS,
  blankCalculatorRow,
  blankRowForSheet,
  createStateFromSeed,
  deepClone,
  normalizeText,
  numberOrZero,
  persistableState,
  recalculateWorkbook,
} from "./engine.js";

const PAGE_SIZE = 20;

const elements = {
  metaGrid: document.getElementById("metaGrid"),
  sheetTabs: document.getElementById("sheetTabs"),
  resetWorkbookButton: document.getElementById("resetWorkbookButton"),
  exportStateButton: document.getElementById("exportStateButton"),
  viewTitle: document.getElementById("viewTitle"),
  viewDescription: document.getElementById("viewDescription"),
  sourceWorkbook: document.getElementById("sourceWorkbook"),
  kpiGrid: document.getElementById("kpiGrid"),
  contentView: document.getElementById("contentView"),
  rowEditorDialog: document.getElementById("rowEditorDialog"),
  editorDialogTitle: document.getElementById("editorDialogTitle"),
  editorDialogBody: document.getElementById("editorDialogBody"),
  closeEditorButton: document.getElementById("closeEditorButton"),
  saveEditorButton: document.getElementById("saveEditorButton"),
  portCodes: document.getElementById("portCodes"),
  fuelTypes: document.getElementById("fuelTypes"),
  imoNumbers: document.getElementById("imoNumbers"),
};

const viewMeta = {
  dashboard: {
    title: "Fleet Exposure Dashboard",
    description: "Review the same allowance, intensity, balance, and penalty outputs that the Excel dashboard rolls up from the Calculator sheet.",
  },
  calculator: {
    title: "Calculator Input and Output",
    description: "Orange fields are editable voyage inputs. Blue cards and tables are spreadsheet-style calculated outputs from the same row.",
  },
  vesselSummary: {
    title: "Vessel Summary",
    description: "Aggregate the calculator data per vessel, just like the workbook summary, with allowance totals and compliance balance status.",
  },
  parameters: {
    title: "Parameters Sheet",
    description: "Maintain reporting year, EUA price, FuelEU targets, GWP assumptions, and electricity factors that drive the rest of the workbook.",
  },
  fuelReference: {
    title: "Fuel Reference Sheet",
    description: "Edit fuel pathways, calorific values, WtT assumptions, ETS factors, and fuel-specific coefficients used by the calculator.",
  },
  fleet: {
    title: "Fleet Database",
    description: "Maintain the IMO master data used to auto-fill vessel name, ship type, flag, and tonnages from the calculator.",
  },
  ports: {
    title: "Port Database",
    description: "Edit UN/LOCODE, country, EU or EEA scope, outermost-region status, and other route classification data.",
  },
  flags: {
    title: "Flag States",
    description: "Edit the reference list used for ship registry context and dropdown support.",
  },
  derogations: {
    title: "Derogations",
    description: "Maintain the derogation reference list for outermost regions and policy guidance context.",
  },
  methodology: {
    title: "Methodology Notes",
    description: "The regulatory methodology notes are editable in the dashboard so the guidance can stay in line with your internal model.",
  },
  formulaGuide: {
    title: "Formula Guide",
    description: "Review and update the plain-English explanation of the workbook formula logic from directly inside the application.",
  },
};

const numericColumns = {
  fuelReference: new Set(["lcvMjPerG", "wtwTankToWakePerG", "rwd", "etsCo2Cf", "cfCo2PerG", "cfCh4PerG", "cfN2oPerG", "cslipPercent"]),
  fleet: new Set(["imoNo", "gt", "nt", "summerDwt", "built"]),
  derogations: new Set(["serialNo"]),
};

const stateStore = {
  seedState: null,
  state: null,
  derived: null,
  ui: {
    activeSheet: "dashboard",
    calculatorSearch: "",
    calculatorSelectedId: null,
    editorSearch: "",
    editorPage: 1,
    dialog: null,
  },
  charts: {},
};

function formatNumber(value, digits = 2) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits === 0 ? 0 : 0,
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

function formatPercent(value) {
  return `${formatNumber((Number(value) || 0) * 100, 0)}%`;
}

function lower(value) {
  return normalizeText(value).toLowerCase();
}

function getCollection(sheetKey) {
  return stateStore.state[sheetKey];
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(persistableState(stateStore.state)));
}

function recomputeAndRender() {
  stateStore.derived = recalculateWorkbook(stateStore.state);
  stateStore.state.parameters = deepClone(stateStore.derived.parameters);
  saveState();
  render();
}

function hydrateFromStorage(seedState) {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return deepClone(seedState);
    }
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.calculatorRows) || !Array.isArray(parsed.parameters)) {
      return deepClone(seedState);
    }
    return parsed;
  } catch {
    return deepClone(seedState);
  }
}

function downloadJson(filename, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function renderMetaGrid() {
  const generatedAt = new Date(stateStore.state.meta.generatedAt);
  const rows = [
    ["Generated", Number.isNaN(generatedAt.getTime()) ? stateStore.state.meta.generatedAt : generatedAt.toLocaleString()],
    ["Calculator rows", formatInteger(stateStore.state.calculatorRows.length)],
    ["Fleet records", formatInteger(stateStore.state.fleet.length)],
    ["Port records", formatInteger(stateStore.state.ports.length)],
    ["Fuel rows", formatInteger(stateStore.state.fuelReference.length)],
    ["Storage", "Local browser cache"],
  ];

  elements.metaGrid.innerHTML = rows
    .map(
      ([label, value]) => `
        <div class="meta-row">
          <dt>${label}</dt>
          <dd>${value}</dd>
        </div>
      `
    )
    .join("");
}

function renderSheetTabs() {
  elements.sheetTabs.innerHTML = SHEETS.map(
    (sheet) => `
      <button
        class="sheet-tab ${stateStore.ui.activeSheet === sheet.key ? "active" : ""}"
        type="button"
        data-action="select-sheet"
        data-sheet="${sheet.key}"
      >
        ${sheet.label}
      </button>
    `
  ).join("");
}

function renderKpis() {
  elements.kpiGrid.innerHTML = stateStore.derived.dashboard.kpis
    .map(
      (kpi) => `
        <article class="kpi-card tone-${kpi.tone || "neutral"}">
          <div class="kpi-label">${kpi.label}</div>
          <div class="kpi-value">${typeof kpi.value === "number" ? formatNumber(kpi.value, 2) : kpi.value}</div>
          <div class="kpi-detail">${kpi.detail}</div>
        </article>
      `
    )
    .join("");
}

function ensureCalculatorSelection() {
  const rows = stateStore.derived.calculatorRows;
  const selected = rows.find((row) => row.id === stateStore.ui.calculatorSelectedId);
  if (selected) {
    return selected;
  }
  const firstActive = rows.find((row) => row.recordId) || rows[0];
  stateStore.ui.calculatorSelectedId = firstActive?.id || null;
  return firstActive || null;
}

function buildDataLists() {
  elements.portCodes.innerHTML = stateStore.state.ports
    .slice(0, 10000)
    .map((row) => `<option value="${row.unlocode}">${row.portName}</option>`)
    .join("");

  elements.fuelTypes.innerHTML = stateStore.derived.fuelReference
    .map((row) => `<option value="${row.fuelPathway}">${row.fuelClass}</option>`)
    .join("");

  elements.imoNumbers.innerHTML = stateStore.state.fleet
    .map((row) => `<option value="${row.imoNo}">${row.vesselName}</option>`)
    .join("");
}

function toneClass(value) {
  if (value > 0) return "tag-good";
  if (value < 0) return "tag-risk";
  return "muted";
}

function renderDashboard() {
  const topRows = stateStore.derived.dashboard.topPenaltyRows;
  const vesselRows = stateStore.derived.dashboard.topExposureVessels;
  return `
    <section class="dashboard-grid">
      <article class="chart-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Exposure by Vessel</p>
            <h3>Top ETS Cost Contributors</h3>
          </div>
        </div>
        <div class="chart-canvas-wrap">
          <canvas id="vesselExposureChart"></canvas>
        </div>
      </article>
      <article class="chart-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Fuel Mix</p>
            <h3>Energy Share by Fuel</h3>
          </div>
        </div>
        <div class="chart-canvas-wrap">
          <canvas id="fuelMixChart"></canvas>
        </div>
      </article>
    </section>

    <section class="dashboard-grid">
      <article class="table-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Calculator Hotspots</p>
            <h3>Highest FuelEU Penalty Rows</h3>
          </div>
          <span class="chip">${topRows.length} rows</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Record</th>
                <th>Vessel</th>
                <th>Route</th>
                <th>Scope</th>
                <th>Balance (t)</th>
                <th>Penalty</th>
              </tr>
            </thead>
            <tbody>
              ${topRows
                .map(
                  (row) => `
                    <tr>
                      <td>${row.recordId || "-"}</td>
                      <td>${row.vesselName || "-"}</td>
                      <td>${row.route}</td>
                      <td>${formatPercent(row.scopePercent)}</td>
                      <td class="${toneClass(row.complianceBalanceT)}">${formatNumber(row.complianceBalanceT, 3)}</td>
                      <td class="number-cell">${formatCurrency(row.fuelEuPenaltyEur)}</td>
                    </tr>
                  `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      </article>

      <article class="card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Workbook Logic</p>
            <h3>What Drives These KPIs</h3>
          </div>
        </div>
        <ul class="insight-list">
          <li>Calculator rows auto-classify into Voyage or Port Stay from the From and To port codes.</li>
          <li>Fleet lookups resolve vessel name, type, flag, GT, NT, and DWT from the Fleet DB using the IMO number.</li>
          <li>Port DB rules determine EU or EEA scope, outermost-region handling, and route percentage for both ETS and FuelEU.</li>
          <li>FuelEU intensity, balance, and penalty update from fuel factors, wind factor, sustainability factor, and OPS electricity.</li>
          <li>The dashboard totals are recomputed every time any sheet data is edited, so the workbook behaves like an application.</li>
        </ul>
      </article>
    </section>

    <section class="table-card">
      <div class="table-head">
        <div>
          <p class="eyebrow">Vessel Rollup</p>
          <h3>Top Vessel Exposure Summary</h3>
        </div>
        <span class="chip">${vesselRows.length} vessels</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Vessel</th>
              <th>Type</th>
              <th>Voyages</th>
              <th>EUAs Required</th>
              <th>ETS Cost</th>
              <th>FuelEU Balance</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${vesselRows
              .map(
                (row) => `
                  <tr>
                    <td>${row.vesselName}</td>
                    <td>${row.shipType}</td>
                    <td class="number-cell">${formatInteger(row.voyageCount)}</td>
                    <td class="number-cell">${formatNumber(row.totalEuasRequired, 2)}</td>
                    <td class="number-cell">${formatCurrency(row.totalEuasCost)}</td>
                    <td class="${toneClass(row.totalComplianceBalance)}">${formatNumber(row.totalComplianceBalance, 3)}</td>
                    <td>${row.status}</td>
                  </tr>
                `
              )
              .join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderDashboardCharts() {
  const vesselRows = stateStore.derived.dashboard.topExposureVessels;
  const fuelMixEntries = Object.entries(stateStore.derived.dashboard.fuelMix);

  Object.values(stateStore.charts).forEach((chart) => chart.destroy());
  stateStore.charts = {};

  const vesselCanvas = document.getElementById("vesselExposureChart");
  const fuelCanvas = document.getElementById("fuelMixChart");

  if (vesselCanvas && vesselRows.length) {
    stateStore.charts.vessel = new Chart(vesselCanvas, {
      type: "bar",
      data: {
        labels: vesselRows.map((row) => row.vesselName),
        datasets: [
          {
            label: "ETS Cost (EUR)",
            data: vesselRows.map((row) => row.totalEuasCost),
            backgroundColor: "#214b79",
            borderRadius: 8,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
        },
        scales: {
          y: {
            ticks: { callback: (value) => formatInteger(value) },
          },
        },
      },
    });
  }

  if (fuelCanvas && fuelMixEntries.length) {
    stateStore.charts.fuel = new Chart(fuelCanvas, {
      type: "doughnut",
      data: {
        labels: fuelMixEntries.map(([label]) => label),
        datasets: [
          {
            data: fuelMixEntries.map(([, value]) => value),
            backgroundColor: ["#214b79", "#ee8c2b", "#2c8c6c", "#9d6cd5", "#7f99b2", "#e1b641"],
            borderWidth: 0,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom" },
        },
      },
    });
  }
}

function renderCalculatorField(field, row) {
  const value = row[field.key] ?? "";
  const common = [
    `class="calculator-input input-orange"`,
    `data-calc-field="${field.key}"`,
    field.list ? `list="${field.list}"` : "",
    field.step ? `step="${field.step}"` : "",
  ].filter(Boolean).join(" ");

  return `
    <div class="field">
      <label for="calc-${field.key}">${field.label}</label>
      <input id="calc-${field.key}" type="${field.type}" ${common} value="${value}">
      ${field.hint ? `<div class="helper-text">${field.hint}</div>` : ""}
    </div>
  `;
}

function renderCalculator() {
  const selectedDerived = ensureCalculatorSelection();
  const selectedInput = stateStore.state.calculatorRows.find((row) => row.id === selectedDerived?.id) || blankCalculatorRow();
  const searchTerm = lower(stateStore.ui.calculatorSearch);
  const activeRows = stateStore.derived.calculatorRows.filter((row) => row.recordId);
  const filteredRows = activeRows.filter((row) => {
    if (!searchTerm) return true;
    return [
      row.recordId,
      row.vesselName,
      row.route,
      row.fromPortCode,
      row.toPortCode,
      row.fuel1Type,
      row.fuel2Type,
      row.bioFuelType,
      row.type,
    ]
      .filter(Boolean)
      .some((value) => lower(value).includes(searchTerm));
  });
  const visibleRows = filteredRows.slice(0, 50);

  return `
    <section class="calculator-shell">
      <div class="calculator-toolbar">
        <div>
          <p class="eyebrow">Row Search</p>
          <input
            class="search-input"
            type="search"
            data-action="calculator-search"
            value="${stateStore.ui.calculatorSearch}"
            placeholder="Search by vessel, route, port code, or fuel"
          >
        </div>
        <div class="calculator-actions">
          <button class="inline-button" type="button" data-action="add-calculator-row">Add row</button>
          <button class="inline-button" type="button" data-action="duplicate-calculator-row">Duplicate selected</button>
          <button class="danger-button" type="button" data-action="delete-calculator-row">Delete selected</button>
        </div>
      </div>

      <div class="calculator-grid">
        <article class="calculator-form-card">
          <div class="table-head">
            <div>
              <p class="eyebrow">Input Form</p>
              <h3>Editable Orange Fields</h3>
            </div>
            <span class="chip">${selectedDerived?.recordId || "Draft row"}</span>
          </div>
          ${CALCULATOR_FIELD_GROUPS.map(
            (group) => `
              <section class="calculator-form-group">
                <h4>${group.title}</h4>
                <div class="field-grid">
                  ${group.fields.map((field) => renderCalculatorField(field, selectedInput)).join("")}
                </div>
              </section>
            `
          ).join("")}
        </article>

        <article class="calculator-output-card">
          <div class="table-head">
            <div>
              <p class="eyebrow">Calculated Output</p>
              <h3>Blue Read-Only Result Cards</h3>
            </div>
            <span class="chip">${selectedDerived?.type || "Awaiting route"}</span>
          </div>

          <div class="record-summary">
            <div class="record-chip"><strong>${selectedDerived?.vesselName || "No vessel yet"}</strong><span>${selectedDerived?.shipType || "Unresolved fleet lookup"}</span></div>
            <div class="record-chip"><strong>${selectedDerived?.route || "Route will appear here"}</strong><span>${selectedDerived?.scopeNote || "Scope note will appear here"}</span></div>
          </div>

          <div class="output-grid">
            <div class="output-card"><span class="output-label">Type</span><strong class="output-value">${selectedDerived?.type || "-"}</strong></div>
            <div class="output-card"><span class="output-label">Scope</span><strong class="output-value">${formatPercent(selectedDerived?.scopePercent || 0)}</strong></div>
            <div class="output-card"><span class="output-label">Vessel Name</span><strong class="output-value">${selectedDerived?.vesselName || "-"}</strong></div>
            <div class="output-card"><span class="output-label">Flag State</span><strong class="output-value">${selectedDerived?.flagState || "-"}</strong></div>
            <div class="output-card"><span class="output-label">Total Energy (MJ)</span><strong class="output-value">${formatInteger(selectedDerived?.totalEnergyMj || 0)}</strong></div>
            <div class="output-card"><span class="output-label">In-Scope Energy (MJ)</span><strong class="output-value">${formatInteger(selectedDerived?.inScopeEnergyMj || 0)}</strong></div>
            <div class="output-card"><span class="output-label">ETS CO2eq (t)</span><strong class="output-value">${formatNumber(selectedDerived?.etsInScopeCo2eqT || 0, 3)}</strong></div>
            <div class="output-card"><span class="output-label">EUAs Required</span><strong class="output-value">${formatNumber(selectedDerived?.euasRequiredT || 0, 3)}</strong></div>
            <div class="output-card"><span class="output-label">ETS Cost</span><strong class="output-value">${formatCurrency(selectedDerived?.euasCostEur || 0)}</strong></div>
            <div class="output-card"><span class="output-label">Attained GHG Intensity</span><strong class="output-value">${formatNumber(selectedDerived?.attainedGhgIntensity || 0, 3)}</strong></div>
            <div class="output-card"><span class="output-label">Compliance Balance</span><strong class="output-value ${toneClass(selectedDerived?.complianceBalanceT || 0)}">${formatNumber(selectedDerived?.complianceBalanceT || 0, 3)}</strong></div>
            <div class="output-card"><span class="output-label">FuelEU Penalty</span><strong class="output-value">${formatCurrency(selectedDerived?.fuelEuPenaltyEur || 0)}</strong></div>
          </div>

          <p class="helper-text">
            This panel is intentionally user-readable. The spreadsheet formulas still run in the application logic, but the output is shown as operations-ready values instead of raw workbook JSON.
          </p>
        </article>
      </div>

      <article class="table-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Calculator Table</p>
            <h3>Rows in the Workbook Model</h3>
          </div>
          <span class="chip">${filteredRows.length} matching rows${filteredRows.length > 50 ? " · showing first 50" : ""}</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Record</th>
                <th>Vessel</th>
                <th>Route</th>
                <th>Type</th>
                <th>Scope</th>
                <th>EUAs</th>
                <th>ETS Cost</th>
                <th>FuelEU Balance</th>
                <th>Penalty</th>
              </tr>
            </thead>
            <tbody>
              ${visibleRows
                .map(
                  (row) => `
                    <tr class="${row.id === stateStore.ui.calculatorSelectedId ? "selected-row" : ""}" data-action="select-calculator-row" data-row-id="${row.id}">
                      <td>${row.recordId || "-"}</td>
                      <td>${row.vesselName || "-"}</td>
                      <td>${row.route}</td>
                      <td>${row.type || "-"}</td>
                      <td>${formatPercent(row.scopePercent || 0)}</td>
                      <td class="number-cell">${formatNumber(row.euasRequiredT || 0, 3)}</td>
                      <td class="number-cell">${formatCurrency(row.euasCostEur || 0)}</td>
                      <td class="${toneClass(row.complianceBalanceT || 0)}">${formatNumber(row.complianceBalanceT || 0, 3)}</td>
                      <td class="number-cell">${formatCurrency(row.fuelEuPenaltyEur || 0)}</td>
                    </tr>
                  `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      </article>
    </section>
  `;
}

function renderVesselSummary() {
  return `
    <section class="table-card">
      <div class="table-head">
        <div>
          <p class="eyebrow">Computed Summary Sheet</p>
          <h3>Per-Vessel Rollup</h3>
        </div>
        <span class="chip">${stateStore.derived.vesselSummary.length} vessels</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>IMO</th>
              <th>Vessel</th>
              <th>Type</th>
              <th>Flag</th>
              <th>Voyages</th>
              <th>EUAs Required</th>
              <th>ETS Cost</th>
              <th>FuelEU Balance</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${stateStore.derived.vesselSummary
              .map(
                (row) => `
                  <tr>
                    <td>${row.imoNo}</td>
                    <td>${row.vesselName}</td>
                    <td>${row.shipType}</td>
                    <td>${row.flag}</td>
                    <td class="number-cell">${formatInteger(row.voyageCount)}</td>
                    <td class="number-cell">${formatNumber(row.totalEuasRequired, 2)}</td>
                    <td class="number-cell">${formatCurrency(row.totalEuasCost)}</td>
                    <td class="${toneClass(row.totalComplianceBalance)}">${formatNumber(row.totalComplianceBalance, 3)}</td>
                    <td>${row.status}</td>
                  </tr>
                `
              )
              .join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function getSheetRowsForDisplay(sheetKey) {
  if (sheetKey === "fuelReference") {
    return stateStore.derived.fuelReference;
  }
  return getCollection(sheetKey) || [];
}

function renderEditorSheet(sheetKey) {
  const allRows = getSheetRowsForDisplay(sheetKey);
  const search = lower(stateStore.ui.editorSearch);
  const filteredRows = allRows.filter((row) => {
    if (!search) return true;
    return Object.values(row).some((value) => lower(value).includes(search));
  });
  const pageCount = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  if (stateStore.ui.editorPage > pageCount) {
    stateStore.ui.editorPage = pageCount;
  }
  const startIndex = (stateStore.ui.editorPage - 1) * PAGE_SIZE;
  const pageRows = filteredRows.slice(startIndex, startIndex + PAGE_SIZE);
  const columns = SHEET_COLUMNS[sheetKey];
  const visibleColumns = columns.slice(0, Math.min(columns.length, 6));

  return `
    <section class="editor-grid">
      <div class="editor-toolbar">
        <input
          class="search-input"
          type="search"
          data-action="editor-search"
          value="${stateStore.ui.editorSearch}"
          placeholder="Search this sheet"
        >
        <div class="toolbar-actions">
          <button class="inline-button" type="button" data-action="open-row-editor" data-sheet="${sheetKey}">Add row</button>
          <span class="page-label">${filteredRows.length} rows</span>
        </div>
      </div>

      <article class="table-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Editable Sheet</p>
            <h3>${viewMeta[sheetKey].title}</h3>
          </div>
          <div class="pagination-row">
            <button class="inline-button" type="button" data-action="editor-page-prev">Previous</button>
            <span class="page-label">Page ${stateStore.ui.editorPage} of ${pageCount}</span>
            <button class="inline-button" type="button" data-action="editor-page-next">Next</button>
          </div>
        </div>
        ${
          pageRows.length
            ? `
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
                      .join("")}
                  </tbody>
                </table>
              </div>
            `
            : `<div class="empty-state">No rows match the current search.</div>`
        }
      </article>
    </section>
  `;
}

function renderContent() {
  const sheetKey = stateStore.ui.activeSheet;
  const meta = viewMeta[sheetKey];
  elements.viewTitle.textContent = meta.title;
  elements.viewDescription.textContent = meta.description;
  elements.sourceWorkbook.textContent = stateStore.state.meta.sourceWorkbook;

  if (sheetKey === "dashboard") {
    elements.contentView.innerHTML = renderDashboard();
    renderDashboardCharts();
    return;
  }

  if (sheetKey === "calculator") {
    elements.contentView.innerHTML = renderCalculator();
    return;
  }

  if (sheetKey === "vesselSummary") {
    elements.contentView.innerHTML = renderVesselSummary();
    return;
  }

  elements.contentView.innerHTML = renderEditorSheet(sheetKey);
}

function render() {
  renderMetaGrid();
  renderSheetTabs();
  renderKpis();
  buildDataLists();
  renderContent();
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
  elements.editorDialogTitle.textContent = `${dialogState.rowId ? "Edit" : "Add"} ${viewMeta[sheetKey].title} row`;
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

function handleContentClick(event) {
  const actionTarget = event.target.closest("[data-action]");
  if (!actionTarget) return;
  const { action, sheet, rowId } = actionTarget.dataset;

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
    const index = stateStore.state.calculatorRows.findIndex((row) => row.id === stateStore.ui.calculatorSelectedId);
    if (index === -1) return;
    stateStore.state.calculatorRows.splice(index, 1);
    stateStore.ui.calculatorSelectedId = null;
    recomputeAndRender();
    return;
  }

  if (action === "open-row-editor") {
    openEditorDialog(sheet, rowId);
    return;
  }

  if (action === "delete-sheet-row") {
    deleteSheetRow(sheet, rowId);
    return;
  }

  if (action === "editor-page-prev") {
    stateStore.ui.editorPage = Math.max(1, stateStore.ui.editorPage - 1);
    render();
    return;
  }

  if (action === "editor-page-next") {
    stateStore.ui.editorPage += 1;
    render();
  }
}

function handleContentInput(event) {
  const calcField = event.target.dataset.calcField;
  if (calcField) {
    updateCalculatorField(calcField, event.target.value);
    return;
  }

  const action = event.target.dataset.action;
  if (action === "calculator-search") {
    stateStore.ui.calculatorSearch = event.target.value;
    render();
    return;
  }

  if (action === "editor-search") {
    stateStore.ui.editorSearch = event.target.value;
    stateStore.ui.editorPage = 1;
    render();
  }
}

function handleTabClick(event) {
  const button = event.target.closest("[data-sheet]");
  if (!button) return;
  stateStore.ui.activeSheet = button.dataset.sheet;
  stateStore.ui.editorSearch = "";
  stateStore.ui.editorPage = 1;
  render();
}

async function bootstrap() {
  const response = await fetch("/data/workbook-seed.json");
  const seed = await response.json();
  stateStore.seedState = createStateFromSeed(seed);
  stateStore.state = hydrateFromStorage(stateStore.seedState);
  stateStore.derived = recalculateWorkbook(stateStore.state);
  stateStore.state.parameters = deepClone(stateStore.derived.parameters);
  ensureCalculatorSelection();

  elements.sheetTabs.addEventListener("click", handleTabClick);
  elements.contentView.addEventListener("click", handleContentClick);
  elements.contentView.addEventListener("input", handleContentInput);
  elements.resetWorkbookButton.addEventListener("click", () => {
    stateStore.state = deepClone(stateStore.seedState);
    stateStore.ui.calculatorSelectedId = null;
    recomputeAndRender();
  });
  elements.exportStateButton.addEventListener("click", () => {
    downloadJson("eu-ets-fueleu-dashboard-state.json", {
      state: persistableState(stateStore.state),
      derived: stateStore.derived,
    });
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
