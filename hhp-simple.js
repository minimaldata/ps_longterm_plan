const YEARS = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const FIRST_FORECAST_YEAR = 2025;
const STORAGE_KEY = "ps-hhp-simple-scenarios-v1";

const planRevenue = [5546199, 10450032, 16513061, 25376000, 36444600, 46813700, 65766000, 83585600, 128908000];

const baseDrivers = {
  retention: [0.4469, 0.5037, 0.583, 0.5931, 0.607, 0.6, 0.6, 0.6, 0.6],
  newCustomers: [186245, 259510, 360475, 422060, 536412, 629192, 725961, 829158, 937328],
  profilesReturning: [1.23, 1.24, 1.22, 1.22, 1.21, 1.21, 1.21, 1.21, 1.21],
  profilesNew: [1.25, 1.24, 1.24, 1.24, 1.22, 1.22, 1.22, 1.22, 1.22],
  revReturningProfile: [14.16, 18.32, 20.76, 24.65, 26.21, 26.21, 26.21, 26.21, 26.21],
  revNewProfile: [19.39, 23.82, 24.44, 28.63, 29.77, 29.77, 29.77, 29.77, 29.77],
};

const historicalReturningCustomers = [48172, 118071, 220125, 344335];

const driverMeta = {
  retention: { label: "Retention", format: "percent", chartId: "chart-retention", yMax: 1, color: "#d39b2a" },
  newCustomers: { label: "New Customers", format: "integer", chartId: "chart-new-customers", yMax: 1000, scale: 1000, suffix: "000s", color: "#6fa76b" },
  profilesReturning: { label: "Profiles / Returning Cust", format: "decimal", chartId: "chart-profiles-returning", yMax: 1.6, color: "#7b6fb8" },
  profilesNew: { label: "Profiles / New Cust", format: "decimal", chartId: "chart-profiles-new", yMax: 1.6, color: "#a96c50" },
  revReturningProfile: { label: "Rev / Returning Cust Profile", format: "currency2", chartId: "chart-rev-returning", yMax: 40, color: "#4f7fb8" },
  revNewProfile: { label: "Rev / New Cust Profile", format: "currency2", chartId: "chart-rev-new", yMax: 40, color: "#4b9b8e" },
};

const outputRows = [
  ["returningCustomers", "Returning Cust", "integer"],
  ["totalCustomers", "Total Customers", "integer"],
  ["revenueExisting", "Revenue Existing Customers", "currency0"],
  ["revenueNew", "Revenue New Customers", "currency0"],
  ["revenue", "Revenue", "currency0"],
  ["planRevenue", "Plan Revenue", "currency0"],
  ["delta", "Delta", "currency0"],
  ["paidProfiles", "Paid Profiles", "integer"],
];

const state = {
  scenarios: loadScenarios(),
  activeScenarioId: "base",
  compareScenarioId: "",
  compareScenarioSnapshot: null,
  selectedVersionId: "",
  charts: {},
  outputCharts: {},
  syncingCharts: false,
  undoStack: [],
  redoStack: [],
  pendingDialogAction: null,
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function makeBaseScenario(name = "Finance Base Case") {
  const scenario = {
    id: "base",
    name,
    drivers: clone(baseDrivers),
  };
  addScenarioVersion(scenario, "Initial");
  return scenario;
}

function loadScenarios() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (parsed && parsed.base && parsed.base.drivers) {
      Object.values(parsed).forEach(normalizeScenario);
      return parsed;
    }
  } catch (error) {
    console.warn("Could not load saved scenarios", error);
  }
  return { base: makeBaseScenario() };
}

function normalizeScenario(scenario) {
  if (!Array.isArray(scenario.versions) || scenario.versions.length === 0) {
    scenario.versions = [];
    addScenarioVersion(scenario, "Initial");
  }
  return scenario;
}

function addScenarioVersion(scenario, label = "Update") {
  if (!Array.isArray(scenario.versions)) scenario.versions = [];
  const createdAt = new Date().toISOString();
  const version = {
    id: `version-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    label,
    createdAt,
    drivers: clone(scenario.drivers),
  };
  scenario.versions.push(version);
  return version;
}

function saveScenarios() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.scenarios));
}

function activeScenario() {
  return state.scenarios[state.activeScenarioId];
}

function compareScenario() {
  if (!state.compareScenarioId || !state.compareScenarioSnapshot) return null;
  return state.compareScenarioSnapshot;
}

function setCompareScenario(scenarioId) {
  state.compareScenarioId = scenarioId;
  state.compareScenarioSnapshot = scenarioId && state.scenarios[scenarioId]
    ? clone(state.scenarios[scenarioId])
    : null;
}

function snapshotState() {
  return {
    scenarios: clone(state.scenarios),
    activeScenarioId: state.activeScenarioId,
    compareScenarioId: state.compareScenarioId,
    compareScenarioSnapshot: clone(state.compareScenarioSnapshot),
    selectedVersionId: state.selectedVersionId,
  };
}

function snapshotsMatch(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function pushUndoSnapshot(snapshot = snapshotState()) {
  const previous = state.undoStack[state.undoStack.length - 1];
  if (previous && snapshotsMatch(previous, snapshot)) return;
  state.undoStack.push(snapshot);
  state.redoStack = [];
  updateHistoryControls();
}

function restoreSnapshot(snapshot) {
  state.scenarios = clone(snapshot.scenarios);
  state.activeScenarioId = snapshot.activeScenarioId;
  state.compareScenarioId = snapshot.compareScenarioId || "";
  state.compareScenarioSnapshot = snapshot.compareScenarioSnapshot ? clone(snapshot.compareScenarioSnapshot) : null;
  state.selectedVersionId = snapshot.selectedVersionId || "";
  saveScenarios();
  renderScenarioSelect();
  syncDriverCharts();
  renderAll();
  updateHistoryControls();
}

function undoChange() {
  const snapshot = state.undoStack.pop();
  if (!snapshot) return;
  state.redoStack.push(snapshotState());
  restoreSnapshot(snapshot);
}

function redoChange() {
  const snapshot = state.redoStack.pop();
  if (!snapshot) return;
  state.undoStack.push(snapshotState());
  restoreSnapshot(snapshot);
}

function updateHistoryControls() {
  const undoButton = document.getElementById("undo-change");
  const redoButton = document.getElementById("redo-change");
  if (undoButton) undoButton.disabled = state.undoStack.length === 0;
  if (redoButton) redoButton.disabled = state.redoStack.length === 0;
}

function editableYear(year) {
  return year >= FIRST_FORECAST_YEAR;
}

function formatValue(value, format) {
  if (!Number.isFinite(value)) return "-";
  if (format === "percent") return `${(value * 100).toFixed(value >= 0.1 ? 0 : 1)}%`;
  if (format === "decimal") return value.toFixed(2);
  if (format === "currency2") return `$${value.toFixed(2)}`;
  if (format === "currency0") return formatCurrency(value, 0);
  return Math.round(value).toLocaleString("en-US");
}

function formatCurrency(value, decimals = 0) {
  const sign = value < 0 ? "-" : "";
  return `${sign}$${Math.abs(value).toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  })}`;
}

function formatCompactCurrency(value) {
  const abs = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (abs >= 1000000000) return `${sign}$${trimNumber(abs / 1000000000, 1)}B`;
  if (abs >= 1000000) return `${sign}$${trimNumber(abs / 1000000, 1)}M`;
  if (abs >= 1000) return `${sign}$${trimNumber(abs / 1000, 0)}k`;
  return formatCurrency(value, 0);
}

function formatCompactNumber(value) {
  const abs = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (abs >= 1000000) return `${sign}${trimNumber(abs / 1000000, 1)}M`;
  if (abs >= 1000) return `${sign}${trimNumber(abs / 1000, 0)}k`;
  return trimNumber(value, 1);
}

function trimNumber(value, decimals) {
  return Number(value).toFixed(decimals).replace(/\.0+$/, "").replace(/(\.\d*[1-9])0+$/, "$1");
}

function formatDriverAxisValue(value, meta) {
  const actualValue = value * (meta.scale || 1);
  if (meta.format === "percent") return `${trimNumber(actualValue * 100, 0)}%`;
  if (meta.format === "currency2") return formatCompactCurrency(actualValue);
  if (meta.format === "integer") return formatCompactNumber(actualValue);
  if (meta.format === "decimal") return trimNumber(actualValue, 2);
  return formatCompactNumber(actualValue);
}

function formatDriverTooltipValue(value, meta) {
  const actualValue = value * (meta.scale || 1);
  if (meta.format === "percent") return `${trimNumber(actualValue * 100, 1)}%`;
  if (meta.format === "currency2") return formatCurrency(actualValue, 2);
  if (meta.format === "integer") return Math.round(actualValue).toLocaleString("en-US");
  if (meta.format === "decimal") return actualValue.toFixed(2);
  return String(actualValue);
}

function formatAxisYear(value) {
  const year = Number(value);
  return Number.isInteger(year) ? String(year) : "";
}

function tooltipHeader(axisValue) {
  const year = Array.isArray(axisValue) ? axisValue[0] : axisValue;
  return `<strong>${year}</strong>`;
}

function parseInput(value, format) {
  const cleaned = String(value).replace(/[$,%\s,]/g, "");
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return null;
  return format === "percent" ? parsed / 100 : parsed;
}

function calculateOutputs(scenario) {
  const drivers = scenario.drivers;
  const returningCustomers = [];
  const totalCustomers = [];
  const revenueExisting = [];
  const revenueNew = [];
  const revenue = [];
  const paidProfiles = [];
  const delta = [];

  YEARS.forEach((year, index) => {
    const returning = index < historicalReturningCustomers.length
      ? historicalReturningCustomers[index]
      : totalCustomers[index - 1] * drivers.retention[index];
    const newCust = drivers.newCustomers[index];
    const existingRev = returning * drivers.profilesReturning[index] * drivers.revReturningProfile[index];
    const newRev = newCust * drivers.profilesNew[index] * drivers.revNewProfile[index];

    returningCustomers.push(returning);
    totalCustomers.push(returning + newCust);
    revenueExisting.push(existingRev);
    revenueNew.push(newRev);
    revenue.push(existingRev + newRev);
    paidProfiles.push((returning * drivers.profilesReturning[index]) + (newCust * drivers.profilesNew[index]));
    delta.push((existingRev + newRev) - planRevenue[index]);
  });

  return {
    returningCustomers,
    totalCustomers,
    revenueExisting,
    revenueNew,
    revenue,
    planRevenue,
    delta,
    paidProfiles,
  };
}

function chartPairs(values, meta) {
  const scale = meta.scale || 1;
  return YEARS.map((year, index) => [year, values[index] / scale]);
}

function makeEditableLine(key, scenario) {
  const meta = driverMeta[key];
  return {
    name: scenario.name,
    color: meta.color,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: chartPairs(scenario.drivers[key], meta),
  };
}

function makeComparisonLine(key, scenario) {
  const meta = driverMeta[key];
  return {
    name: `Comparison: ${scenario.name}`,
    color: "#98a2b3",
    editable: false,
    data: chartPairs(scenario.drivers[key], meta),
  };
}

function driverChartLines(key) {
  const lines = [makeEditableLine(key, activeScenario())];
  const comparison = compareScenario();
  if (comparison) lines.push(makeComparisonLine(key, comparison));
  return lines;
}

function styleComparisonSeries(chart) {
  const option = chart.chart.getOption();
  const series = (option.series || []).map(seriesItem => {
    if (!String(seriesItem.name || "").startsWith("Comparison:")) {
      return {
        ...seriesItem,
        z: 4,
      };
    }
    return {
      ...seriesItem,
      silent: true,
      showSymbol: false,
      symbolSize: 0,
      z: 2,
      itemStyle: {
        ...(seriesItem.itemStyle || {}),
        color: "#98a2b3",
        opacity: 1,
      },
      lineStyle: {
        ...(seriesItem.lineStyle || {}),
        color: "#98a2b3",
        width: 3,
        opacity: 1,
        type: "solid",
      },
      tooltip: { show: true },
    };
  });
  chart.chart.setOption({ series }, false);
}

function initDriverChart(key) {
  const meta = driverMeta[key];
  const chart = new NapkinChart(
    meta.chartId,
    driverChartLines(key),
    true,
    {
      animation: false,
      xAxis: {
        type: "value",
        min: YEARS[0],
        max: YEARS[YEARS.length - 1],
        minInterval: 1,
        axisLabel: { formatter: formatAxisYear },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: meta.yMax,
        axisLabel: { formatter: value => formatDriverAxisValue(value, meta) },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: value => formatDriverTooltipValue(value, meta),
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const seen = new Set();
          const rows = [];
          let year = items[0]?.axisValue;
          items
            .filter(item => item && item.seriesType === "line" && item.seriesIndex % 2 === 1)
            .forEach(item => {
              if (seen.has(item.seriesName)) return;
              seen.add(item.seriesName);
              const data = Array.isArray(item.data) ? item.data : item.value;
              year = Array.isArray(data) ? data[0] : item.axisValue;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatDriverTooltipValue(Number(rawValue), meta)}`);
            });
          if (!rows.length && items[0]) {
            const item = items[0];
            const data = Array.isArray(item.data) ? item.data : item.value;
            year = Array.isArray(data) ? data[0] : item.axisValue;
            const rawValue = Array.isArray(data) ? data[1] : item.value;
            rows.push(`${item.marker || ""} ${item.seriesName || meta.label}: ${formatDriverTooltipValue(Number(rawValue), meta)}`);
          }
          return [tooltipHeader(year), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );

  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._refreshChart();
  styleComparisonSeries(chart);
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const next = Array(YEARS.length).fill(null);
    chart.lines[0].data.forEach(([year, value]) => {
      const index = YEARS.indexOf(Number(year));
      if (index >= 0) next[index] = Number(value) * (meta.scale || 1);
    });
    next.forEach((value, index) => {
      if (value !== null && editableYear(YEARS[index])) activeScenario().drivers[key][index] = value;
    });
    saveScenarios();
    styleComparisonSeries(chart);
    renderAll();
  };
  state.charts[key] = chart;
}

function syncDriverCharts() {
  state.syncingCharts = true;
  Object.entries(driverMeta).forEach(([key, meta]) => {
    const chart = state.charts[key];
    if (!chart) return;
    chart.lines = driverChartLines(key);
    chart._refreshChart();
    styleComparisonSeries(chart);
  });
  state.syncingCharts = false;
}

function initOutputCharts() {
  state.outputCharts.revenue = echarts.init(document.getElementById("revenue-chart"));
  state.outputCharts.gap = echarts.init(document.getElementById("gap-chart"));
  window.addEventListener("resize", () => {
    state.outputCharts.revenue.resize();
    state.outputCharts.gap.resize();
  });
}

function renderOutputCharts(outputs) {
  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
  const revenueSeries = [
    { name: `${activeScenario().name} Existing Customers`, type: "bar", stack: "revenue", data: outputs.revenueExisting, itemStyle: { color: "#4f7fb8" } },
    { name: `${activeScenario().name} New Customers`, type: "bar", stack: "revenue", data: outputs.revenueNew, itemStyle: { color: "#8ebf86" } },
    { name: "Plan Revenue", type: "line", data: outputs.planRevenue, symbolSize: 6, lineStyle: { color: "#111827", width: 2 }, itemStyle: { color: "#111827" } },
  ];
  if (comparisonOutputs) {
    revenueSeries.push({
      name: `${comparison.name} Revenue`,
      type: "line",
      data: comparisonOutputs.revenue,
      symbolSize: 5,
      lineStyle: { color: "#98a2b3", width: 2 },
      itemStyle: { color: "#98a2b3" },
    });
  }

  const gapSeries = [{
    name: `${activeScenario().name} Delta to Plan`,
    type: "bar",
    data: outputs.delta,
    itemStyle: { color: value => value.value < 0 ? "#b42318" : "#4f7f52" },
  }];
  if (comparisonOutputs) {
    gapSeries.push({
      name: `${comparison.name} Delta to Plan`,
      type: "line",
      data: comparisonOutputs.delta,
      symbolSize: 5,
      lineStyle: { color: "#98a2b3", width: 2 },
      itemStyle: { color: "#98a2b3" },
    });
  }

  state.outputCharts.revenue.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const lines = [tooltipHeader(params[0]?.axisValue)];
        params.forEach(item => {
          lines.push(`${item.marker} ${item.seriesName}: ${formatCurrency(Number(item.value), 0)}`);
        });
        const total = params
          .filter(item => item.seriesType === "bar" && item.seriesName.startsWith(activeScenario().name))
          .reduce((sum, item) => sum + Number(item.value || 0), 0);
        lines.push(`<strong>${activeScenario().name} Revenue: ${formatCurrency(total, 0)}</strong>`);
        return lines.join("<br/>");
      },
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
    series: revenueSeries,
  }, true);

  state.outputCharts.gap.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const item = params[0];
        return [
          tooltipHeader(item.axisValue),
          ...params.map(item => `${item.marker} ${item.seriesName}: ${formatCurrency(Number(item.value), 0)}`),
        ].join("<br/>");
      },
    },
    grid: { left: 12, right: 18, top: 18, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
    series: gapSeries,
  }, true);
}

function renderKpis(outputs) {
  const last = YEARS.length - 1;
  const firstForecast = YEARS.indexOf(2025);
  const cagr = Math.pow(outputs.revenue[last] / outputs.revenue[firstForecast], 1 / (YEARS[last] - YEARS[firstForecast])) - 1;
  document.getElementById("kpi-revenue-2029").textContent = formatCurrency(outputs.revenue[last], 0);
  document.getElementById("kpi-gap-2029").textContent = formatCurrency(outputs.delta[last], 0);
  document.getElementById("kpi-gap-2029").classList.toggle("negative", outputs.delta[last] < 0);
  document.getElementById("kpi-cagr").textContent = `${(cagr * 100).toFixed(1)}%`;
  document.getElementById("kpi-paid-profiles").textContent = Math.round(outputs.paidProfiles[last]).toLocaleString("en-US");
}

function renderTable(outputs) {
  const table = document.getElementById("driver-table");
  const driverRows = Object.entries(driverMeta).map(([key, meta]) => ({
    key,
    label: meta.label,
    format: meta.format,
    values: activeScenario().drivers[key],
    editable: true,
  }));

  const rows = [
    { section: "Inputs" },
    ...driverRows,
    { section: "Calculated Outputs" },
    ...outputRows.map(([key, label, format]) => ({ key, label, format, values: outputs[key], editable: false })),
  ];

  table.innerHTML = `
    <thead>
      <tr>
        <th>Metric</th>
        ${YEARS.map(year => `<th>${year}</th>`).join("")}
      </tr>
    </thead>
    <tbody>
      ${rows.map(row => {
        if (row.section) return `<tr class="section-row"><td colspan="${YEARS.length + 1}">${row.section}</td></tr>`;
        return `
          <tr data-key="${row.key}">
            <td>${row.label}</td>
            ${YEARS.map((year, index) => renderCell(row, year, index)).join("")}
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  table.querySelectorAll("input[data-key]").forEach(input => {
    input.addEventListener("change", event => {
      const key = event.target.dataset.key;
      const index = Number(event.target.dataset.index);
      const meta = driverMeta[key];
      const parsed = parseInput(event.target.value, meta.format);
      if (parsed === null) {
        event.target.value = formatValue(activeScenario().drivers[key][index], meta.format);
        return;
      }
      if (parsed === activeScenario().drivers[key][index]) return;
      pushUndoSnapshot();
      activeScenario().drivers[key][index] = parsed;
      saveScenarios();
      syncDriverCharts();
      renderAll();
    });
  });
}

function renderCell(row, year, index) {
  const value = row.values[index];
  if (row.editable && editableYear(year)) {
    return `
      <td class="input">
        <input data-key="${row.key}" data-index="${index}" value="${formatValue(value, row.format)}" aria-label="${row.label} ${year}" />
      </td>
    `;
  }
  const classes = [
    row.editable ? "historical" : "output",
    row.key === "delta" && value < 0 ? "negative" : "",
  ].filter(Boolean).join(" ");
  return `<td class="${classes}">${formatValue(value, row.format)}</td>`;
}

function renderScenarioSelect() {
  const scenarios = Object.values(state.scenarios);
  const select = document.getElementById("scenario-select");
  const compareSelect = document.getElementById("compare-scenario-select");
  const versionSelect = document.getElementById("version-select");
  const restoreButton = document.getElementById("restore-version");
  select.innerHTML = scenarios
    .map(scenario => `<option value="${scenario.id}">${scenario.name}</option>`)
    .join("");
  select.value = state.activeScenarioId;
  compareSelect.innerHTML = [
    `<option value="">None</option>`,
    ...scenarios.map(scenario => `<option value="${scenario.id}">${scenario.name}</option>`),
  ].join("");
  if (state.compareScenarioId && !state.scenarios[state.compareScenarioId]) {
    setCompareScenario("");
  }
  compareSelect.value = state.compareScenarioId;

  const versions = activeScenario().versions || [];
  versionSelect.innerHTML = versions.length
    ? versions.slice().reverse().map(version => `<option value="${version.id}">${formatVersionLabel(version)}</option>`).join("")
    : `<option value="">No versions</option>`;
  if (!state.selectedVersionId || !versions.some(version => version.id === state.selectedVersionId)) {
    state.selectedVersionId = versions.length ? versions[versions.length - 1].id : "";
  }
  versionSelect.value = state.selectedVersionId;
  restoreButton.disabled = !state.selectedVersionId;
}

function formatVersionLabel(version) {
  const date = new Date(version.createdAt);
  const timestamp = Number.isNaN(date.getTime())
    ? version.createdAt
    : date.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  return `${timestamp} - ${version.label}`;
}

function renderAll() {
  const outputs = calculateOutputs(activeScenario());
  renderKpis(outputs);
  renderOutputCharts(outputs);
  renderTable(outputs);
}

function openScenarioDialog({ title, label, defaultValue, confirmText, onConfirm }) {
  const dialog = document.getElementById("scenario-dialog");
  const titleEl = document.getElementById("scenario-dialog-title");
  const labelEl = document.getElementById("scenario-dialog-label");
  const input = document.getElementById("scenario-dialog-input");
  const confirmButton = document.getElementById("scenario-dialog-confirm");

  state.pendingDialogAction = onConfirm;
  titleEl.textContent = title;
  labelEl.textContent = label;
  confirmButton.textContent = confirmText;
  input.value = defaultValue;
  dialog.showModal();
  setTimeout(() => {
    input.focus();
    input.select();
  }, 0);
}

function closeScenarioDialog() {
  const dialog = document.getElementById("scenario-dialog");
  state.pendingDialogAction = null;
  dialog.close();
}

function confirmScenarioDialog() {
  const input = document.getElementById("scenario-dialog-input");
  const value = input.value.trim();
  if (!value || typeof state.pendingDialogAction !== "function") return;
  const action = state.pendingDialogAction;
  state.pendingDialogAction = null;
  document.getElementById("scenario-dialog").close();
  action(value);
}

function saveAsScenario() {
  openScenarioDialog({
    title: "Save Scenario As",
    label: "Scenario name",
    defaultValue: `${activeScenario().name} copy`,
    confirmText: "Save As",
    onConfirm: createScenarioCopy,
  });
}

function createScenarioCopy(name) {
  if (!name) return;
  const id = `scenario-${Date.now()}`;
  const previousActiveId = state.activeScenarioId;
  pushUndoSnapshot();
  state.scenarios[id] = {
    id,
    name,
    drivers: clone(activeScenario().drivers),
  };
  addScenarioVersion(state.scenarios[id], "Save As");
  state.activeScenarioId = id;
  state.selectedVersionId = state.scenarios[id].versions[state.scenarios[id].versions.length - 1].id;
  setCompareScenario(previousActiveId);
  saveScenarios();
  renderScenarioSelect();
  syncDriverCharts();
  renderAll();
}

function updateScenario() {
  openScenarioDialog({
    title: "Update Scenario",
    label: "Version note",
    defaultValue: "Update",
    confirmText: "Update",
    onConfirm: saveScenarioVersion,
  });
}

function saveScenarioVersion(label) {
  pushUndoSnapshot();
  const version = addScenarioVersion(activeScenario(), label.trim() || "Update");
  state.selectedVersionId = version.id;
  saveScenarios();
  renderScenarioSelect();
}

function restoreVersion() {
  const scenario = activeScenario();
  const version = (scenario.versions || []).find(item => item.id === state.selectedVersionId);
  if (!version) return;
  pushUndoSnapshot();
  scenario.drivers = clone(version.drivers);
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function resetScenario() {
  pushUndoSnapshot();
  if (state.activeScenarioId === "base") {
    state.scenarios.base = makeBaseScenario();
  } else {
    activeScenario().drivers = clone(baseDrivers);
  }
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function bindControls() {
  document.getElementById("scenario-select").addEventListener("change", event => {
    state.activeScenarioId = event.target.value;
    state.selectedVersionId = "";
    updateHistoryControls();
    renderScenarioSelect();
    syncDriverCharts();
    renderAll();
  });
  document.getElementById("compare-scenario-select").addEventListener("change", event => {
    setCompareScenario(event.target.value);
    syncDriverCharts();
    renderAll();
  });
  document.getElementById("undo-change").addEventListener("click", undoChange);
  document.getElementById("redo-change").addEventListener("click", redoChange);
  document.getElementById("version-select").addEventListener("change", event => {
    state.selectedVersionId = event.target.value;
    renderScenarioSelect();
  });
  document.getElementById("restore-version").addEventListener("click", restoreVersion);
  document.getElementById("scenario-dialog-cancel").addEventListener("click", closeScenarioDialog);
  document.getElementById("scenario-dialog-confirm").addEventListener("click", confirmScenarioDialog);
  document.getElementById("scenario-dialog-input").addEventListener("keydown", event => {
    if (event.key === "Enter") {
      event.preventDefault();
      confirmScenarioDialog();
    } else if (event.key === "Escape") {
      closeScenarioDialog();
    }
  });
  document.getElementById("save-as-scenario").addEventListener("click", saveAsScenario);
  document.getElementById("update-scenario").addEventListener("click", updateScenario);
  document.getElementById("reset-scenario").addEventListener("click", resetScenario);

  document.getElementById("tab-chart").addEventListener("click", () => setView("chart"));
  document.getElementById("tab-table").addEventListener("click", () => setView("table"));
  document.addEventListener("keydown", event => {
    const target = event.target;
    const isTextInput = target && ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName);
    if (isTextInput) return;
    const key = event.key.toLowerCase();
    if ((event.metaKey || event.ctrlKey) && key === "z" && !event.shiftKey) {
      event.preventDefault();
      undoChange();
    } else if ((event.metaKey || event.ctrlKey) && (key === "y" || (key === "z" && event.shiftKey))) {
      event.preventDefault();
      redoChange();
    }
  });
}

function setView(view) {
  document.getElementById("tab-chart").classList.toggle("active", view === "chart");
  document.getElementById("tab-table").classList.toggle("active", view === "table");
  document.getElementById("chart-view").classList.toggle("active", view === "chart");
  document.getElementById("table-view").classList.toggle("active", view === "table");
  setTimeout(() => {
    Object.values(state.charts).forEach(chart => chart.resize());
    Object.values(state.outputCharts).forEach(chart => chart.resize());
  }, 0);
}

function init() {
  renderScenarioSelect();
  bindControls();
  initOutputCharts();
  Object.keys(driverMeta).forEach(initDriverChart);
  updateHistoryControls();
  renderAll();
}

init();
