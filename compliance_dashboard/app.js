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

const PAGE_SIZE = 18;
const REFERENCE_SHEETS = SHEETS.filter((sheet) => !["dashboard", "calculator", "vesselSummary"].includes(sheet.key));

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
  },
};

const numericColumns = {
  fuelReference: new Set(["lcvMjPerG", "wtWPerMj", "rwd", "etsCo2Cf", "cfCo2PerG", "cfCh4PerG", "cfN2oPerG", "cslipPercent"]),
  fleet: new Set(["imoNo", "gt", "nt", "summerDwt", "built"]),
  derogations: new Set(["serialNo"]),
};

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
  const totalEuasRequired = activeRows.reduce((sum, row) => sum + numberOrZero(row.euasRequiredT), 0);
  const totalEuasCost = activeRows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0);
  const totalPenalty = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuPenaltyEur), 0);
  const totalFuelConsumed = activeRows.reduce(
    (sum, row) =>
      sum +
      numberOrZero(row.fuel1ConsumptionMt) +
      numberOrZero(row.fuel2ConsumptionMt) +
      numberOrZero(row.bioFuelConsumptionMt),
    0
  );
  const totalNumerator = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuWtwEmissionsG), 0);
  const totalDenominator = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuDenomStep1Mj), 0);
  const totalEnergy = activeRows.reduce((sum, row) => sum + numberOrZero(row.fuelEuEnergyStep2Mj), 0);
  const averageIntensity = totalDenominator > 0 ? totalNumerator / totalDenominator : 0;
  const target = stateStore.derived.parameterValues.fueleuTarget;
  const complianceBalance = totalDenominator > 0 ? (target - averageIntensity) * totalEnergy / 1_000_000 : 0;
  const voyageRows = activeRows.filter((row) => row.type === "Voyage");
  const portStayRows = activeRows.filter((row) => row.type === "Port Stay");

  const byVessel = [...new Set(activeRows.map((row) => row.vesselName).filter(Boolean))]
    .map((vesselName) => {
      const rows = activeRows.filter((row) => row.vesselName === vesselName);
      return {
        vesselName,
        euasRequired: rows.reduce((sum, row) => sum + numberOrZero(row.euasRequiredT), 0),
        euasCost: rows.reduce((sum, row) => sum + numberOrZero(row.euasCostEur), 0),
        averageIntensity: (() => {
          const numerator = rows.reduce((sum, row) => sum + numberOrZero(row.fuelEuWtwEmissionsG), 0);
          const denominator = rows.reduce((sum, row) => sum + numberOrZero(row.fuelEuDenomStep1Mj), 0);
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
  };
}

function renderKpis() {
  const dashboard = computeFilteredDashboard(getActiveRows());
  const cards = [
    {
      label: "Total EUAs required",
      value: `${formatNumber(dashboard.totalEuasRequired, 1)}`,
      detail: "t CO2eq",
      note: "Click charts for record-level rows",
      tone: "risk",
    },
    {
      label: "Total EUA cost",
      value: `${formatCurrency(dashboard.totalEuasCost)}`,
      detail: `@ EUR ${formatInteger(stateStore.derived.parameterValues.euaPrice)} / EUA`,
      note: "Filtered by current vessel",
      tone: "warn",
    },
    {
      label: "Compliance balance",
      value: `${dashboard.complianceBalance >= 0 ? "+" : ""}${formatNumber(dashboard.complianceBalance, 1)}`,
      detail: "t CO2eq surplus / deficit",
      note: "Based on FuelEU target",
      tone: dashboard.complianceBalance >= 0 ? "good" : "risk",
    },
    {
      label: "FuelEU penalty",
      value: `${formatCurrency(dashboard.totalPenalty)}`,
      detail: dashboard.totalPenalty > 0 ? "Penalty triggered by deficits" : "No penalty due",
      note: "Calculator-derived total",
      tone: dashboard.totalPenalty > 0 ? "risk" : "good",
    },
    {
      label: "Avg GHG intensity",
      value: `${formatNumber(dashboard.averageIntensity, 2)}`,
      detail: `g/MJ vs ${formatNumber(stateStore.derived.parameterValues.fueleuTarget, 2)} target`,
      note: "Voyage and port-stay weighted",
      tone: dashboard.averageIntensity <= stateStore.derived.parameterValues.fueleuTarget ? "good" : "warn",
    },
    {
      label: "Total fuel consumed",
      value: `${formatNumber(dashboard.totalFuelConsumed, 1)}`,
      detail: "MT all fuel types",
      note: "Fossil plus biofuel",
      tone: "neutral",
    },
    {
      label: "Voyage records",
      value: `${formatInteger(dashboard.voyageRows.length)}`,
      detail: "Click charts to view voyages",
      note: "Current filter only",
      tone: "neutral",
    },
    {
      label: "Port stay records",
      value: `${formatInteger(dashboard.portStayRows.length)}`,
      detail: "Click charts to view port stays",
      note: "Current filter only",
      tone: "neutral",
    },
  ];

  elements.kpiGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="kpi-card tone-${card.tone}">
          <div class="kpi-label">${card.label}</div>
          <div class="kpi-value">${card.value}</div>
          <div class="kpi-detail">${card.detail}</div>
          <div class="kpi-note">${card.note}</div>
        </article>
      `
    )
    .join("");
}

function toneClass(value) {
  if (value > 0) return "tag-good";
  if (value < 0) return "tag-risk";
  return "muted";
}

function openDrilldown(title, subtitle, columns, rows) {
  stateStore.ui.drilldown = { title, subtitle, columns, rows };
  render();
}

function closeDrilldown() {
  stateStore.ui.drilldown = null;
  render();
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

function renderDashboard() {
  return `
    <section class="analytics-header">
      <div class="analytics-title">
        <h2>Visual Analytics</h2>
        <span class="chip">4 charts</span>
      </div>
      <p class="helper-text">Click any chart element to open its underlying table in the right pane.</p>
    </section>

    <section class="dashboard-layout">
      <div class="dashboard-main">
        <div class="chart-grid">
          <article class="chart-card">
            <div class="table-head">
              <div>
                <p class="eyebrow">By Vessel</p>
                <h3>EUAs Required by Vessel (t CO2eq)</h3>
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
        </div>
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

  const vesselCanvas = document.getElementById("vesselEuaChart");
  const ghgCanvas = document.getElementById("voyageGhgChart");
  const voyageEuaCanvas = document.getElementById("voyageEuaChart");
  const costSplitCanvas = document.getElementById("costSplitChart");

  if (vesselCanvas) {
    stateStore.charts.vesselEua = new Chart(vesselCanvas, {
      type: "bar",
      data: {
        labels: dashboard.byVessel.map((row) => row.vesselName),
        datasets: [
          {
            label: "EUAs required",
            data: dashboard.byVessel.map((row) => row.euasRequired),
            backgroundColor: "#4288d6",
            borderRadius: 8,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        onClick: (_, elementsClicked) => {
          if (!elementsClicked.length) return;
          const index = elementsClicked[0].index;
          const vesselName = dashboard.byVessel[index]?.vesselName;
          const rows = activeRows.filter((row) => row.vesselName === vesselName);
          openDrilldown(
            `EUAs for ${vesselName}`,
            "Voyage and port stay rows contributing to the vessel total.",
            ["Record", "Type", "Route", "EUAs Required", "ETS Cost", "GHG Intensity"],
            rows.map((row) => [
              row.recordId,
              row.type,
              row.route,
              formatNumber(row.euasRequiredT, 3),
              formatCurrency(row.euasCostEur),
              formatNumber(row.attainedGhgIntensity, 3),
            ])
          );
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
}

function ensureCalculatorSelection() {
  const activeRows = getActiveRows();
  const rows = activeRows.length ? activeRows : stateStore.derived.calculatorRows;
  const selected = rows.find((row) => row.id === stateStore.ui.calculatorSelectedId);
  if (selected) return selected;
  const fallback = rows.find((row) => row.recordId) || rows[0];
  stateStore.ui.calculatorSelectedId = fallback?.id || null;
  return fallback || null;
}

function renderCalculatorField(field, row) {
  const value = row[field.key] ?? "";
  return `
    <div class="field">
      <label for="calc-${field.key}">${field.label}</label>
      <input
        id="calc-${field.key}"
        class="calculator-input input-orange"
        data-calc-field="${field.key}"
        type="${field.type}"
        value="${value}"
        ${field.list ? `list="${field.list}"` : ""}
        ${field.step ? `step="${field.step}"` : ""}
      >
      ${field.hint ? `<div class="helper-text">${field.hint}</div>` : ""}
    </div>
  `;
}

function renderCalculator() {
  const selectedDerived = ensureCalculatorSelection();
  const selectedInput = stateStore.state.calculatorRows.find((row) => row.id === selectedDerived?.id) || blankCalculatorRow();
  const searchTerm = lower(stateStore.ui.calculatorSearch);
  const filteredRows = getActiveRows().filter((row) => {
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
  const visibleRows = filteredRows.slice(0, 60);

  return `
    <section class="calculator-shell">
      <div class="calculator-toolbar">
        <div class="calculator-toolbar-copy">
          <h2>Calculator</h2>
          <p class="helper-text">Enter route and fuel data in orange fields. The blue output panel recalculates EU ETS and FuelEU results instantly.</p>
        </div>
        <div class="calculator-actions">
          <input
            class="search-input"
            type="search"
            data-action="calculator-search"
            value="${stateStore.ui.calculatorSearch}"
            placeholder="Search records, routes, or fuels"
          >
          <button class="inline-button" type="button" data-action="add-calculator-row">Add row</button>
          <button class="inline-button" type="button" data-action="duplicate-calculator-row">Duplicate</button>
          <button class="danger-button" type="button" data-action="delete-calculator-row">Delete</button>
        </div>
      </div>

      <div class="calculator-grid">
        <article class="calculator-form-card">
          <div class="table-head">
            <div>
              <p class="eyebrow">Operator Inputs</p>
              <h3>${selectedDerived?.recordId || "New scenario row"}</h3>
            </div>
            <span class="chip">${selectedDerived?.vesselName || "Unassigned vessel"}</span>
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
              <h3>${selectedDerived?.route || "Awaiting route classification"}</h3>
            </div>
            <span class="chip">${selectedDerived?.type || "-"}</span>
          </div>

          <div class="record-summary">
            <div class="record-chip"><strong>${selectedDerived?.vesselName || "-"}</strong><span>${selectedDerived?.shipType || "Resolve IMO to auto-fill vessel data"}</span></div>
            <div class="record-chip"><strong>${selectedDerived?.scopeNote || "Scope note pending"}</strong><span>${selectedDerived?.flagState || "No flag yet"}</span></div>
          </div>

          <div class="output-grid">
            <div class="output-card"><span class="output-label">EUAs Required</span><strong class="output-value">${formatNumber(selectedDerived?.euasRequiredT, 3)}</strong></div>
            <div class="output-card"><span class="output-label">ETS Cost</span><strong class="output-value">${formatCurrency(selectedDerived?.euasCostEur)}</strong></div>
            <div class="output-card"><span class="output-label">GHG Intensity</span><strong class="output-value">${formatNumber(selectedDerived?.attainedGhgIntensity, 3)}</strong></div>
            <div class="output-card"><span class="output-label">Target Intensity</span><strong class="output-value">${formatNumber(selectedDerived?.targetGhgIntensity, 3)}</strong></div>
            <div class="output-card"><span class="output-label">Compliance Balance</span><strong class="output-value ${toneClass(selectedDerived?.complianceBalanceT || 0)}">${formatNumber(selectedDerived?.complianceBalanceT, 3)}</strong></div>
            <div class="output-card"><span class="output-label">FuelEU Penalty</span><strong class="output-value">${formatCurrency(selectedDerived?.fuelEuPenaltyEur)}</strong></div>
            <div class="output-card"><span class="output-label">ETS CO2eq</span><strong class="output-value">${formatNumber(selectedDerived?.etsInScopeCo2eqT, 3)}</strong></div>
            <div class="output-card"><span class="output-label">Scope</span><strong class="output-value">${formatPercent(selectedDerived?.scopePercent || 0)}</strong></div>
            <div class="output-card"><span class="output-label">Fuel Consumed</span><strong class="output-value">${formatNumber(numberOrZero(selectedDerived?.fuel1ConsumptionMt) + numberOrZero(selectedDerived?.fuel2ConsumptionMt) + numberOrZero(selectedDerived?.bioFuelConsumptionMt), 2)} MT</strong></div>
            <div class="output-card"><span class="output-label">Transport Work</span><strong class="output-value">${formatInteger(selectedDerived?.transportWork || 0)}</strong></div>
            <div class="output-card"><span class="output-label">In-Scope Energy</span><strong class="output-value">${formatInteger(selectedDerived?.inScopeEnergyMj || 0)} MJ</strong></div>
            <div class="output-card"><span class="output-label">EEOI</span><strong class="output-value">${formatNumber(selectedDerived?.eeoi, 3)}</strong></div>
          </div>
        </article>
      </div>

      <article class="table-card">
        <div class="table-head">
          <div>
            <p class="eyebrow">Scenario Table</p>
            <h3>Filtered rows${stateStore.ui.vesselFilter !== "all" ? ` for ${stateStore.ui.vesselFilter}` : ""}</h3>
          </div>
          <span class="chip">${filteredRows.length} rows${filteredRows.length > 60 ? " · showing first 60" : ""}</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Record</th>
                <th>Vessel</th>
                <th>Route</th>
                <th>Type</th>
                <th>EUAs</th>
                <th>GHG Intensity</th>
                <th>Balance</th>
                <th>Penalty</th>
              </tr>
            </thead>
            <tbody>
              ${visibleRows
                .map(
                  (row) => `
                    <tr class="${row.id === stateStore.ui.calculatorSelectedId ? "selected-row" : ""}" data-action="select-calculator-row" data-row-id="${row.id}">
                      <td>${row.recordId}</td>
                      <td>${row.vesselName}</td>
                      <td>${row.route}</td>
                      <td>${row.type}</td>
                      <td class="number-cell">${formatNumber(row.euasRequiredT, 3)}</td>
                      <td class="number-cell">${formatNumber(row.attainedGhgIntensity, 3)}</td>
                      <td class="${toneClass(row.complianceBalanceT || 0)}">${formatNumber(row.complianceBalanceT, 3)}</td>
                      <td class="number-cell">${formatCurrency(row.fuelEuPenaltyEur)}</td>
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
  renderKpis();
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
    const index = stateStore.state.calculatorRows.findIndex((row) => row.id === stateStore.ui.calculatorSelectedId);
    if (index === -1) return;
    stateStore.state.calculatorRows.splice(index, 1);
    stateStore.ui.calculatorSelectedId = null;
    recomputeAndRender();
    return;
  }

  if (action === "close-drilldown") {
    closeDrilldown();
  }
}

function handleMainInput(event) {
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
