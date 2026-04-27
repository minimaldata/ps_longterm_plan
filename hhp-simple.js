const YEARS = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const FIRST_FORECAST_YEAR = 2026;
const STORAGE_KEY = "ps-hhp-simple-scenarios-v1";
const TOP_DOWN_COLOR = "#6fa76b";
const BOTTOM_UP_COLOR = "#4f7fb8";

const planRevenue = [5546199, 10450032, 16513061, 25376000, 36444600, 46813700, 65766000, 83585600, 128908000];
const NEW_CUSTOMER_COHORT_YEARS = [0, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const NEW_CUSTOMER_BASELINE_YEAR = 2025;
const NEW_CUSTOMER_BASELINE_VALUE = 536412;

const baseDrivers = {
  retention: [0.4469, 0.5037, 0.583, 0.5931, 0.607, 0.6, 0.6, 0.6, 0.6],
  newCustomers: [186245, 259510, 360475, 422060, 536412, 629192, 725961, 829158, 937328],
  profilesReturning: [1.23, 1.24, 1.22, 1.22, 1.21, 1.21, 1.21, 1.21, 1.21],
  profilesNew: [1.25, 1.24, 1.24, 1.24, 1.22, 1.22, 1.22, 1.22, 1.22],
  revReturningProfile: [14.16, 18.32, 20.76, 24.65, 26.21, 26.21, 26.21, 26.21, 26.21],
  revNewProfile: [19.39, 23.82, 24.44, 28.63, 29.77, 29.77, 29.77, 29.77, 29.77],
};

const historicalReturningCustomers = [48172, 118071, 220125, 344335];
const baseNewCustomerCohortCounts = {
  2017: { 0: 90, 2017: 1025 },
  2018: { 0: 1885, 2017: 5190, 2018: 8666 },
  2019: { 0: 3919, 2017: 5127, 2018: 18103, 2019: 11564 },
  2020: { 0: 7115, 2017: 6130, 2018: 22788, 2019: 30675, 2020: 31222 },
  2021: { 0: 12491, 2017: 8575, 2018: 29355, 2019: 34867, 2020: 58371, 2021: 42586 },
  2022: { 0: 7876, 2017: 12069, 2018: 31070, 2019: 34085, 2020: 48613, 2021: 69945, 2022: 55852 },
  2023: { 0: 10464, 2017: 11344, 2018: 30954, 2019: 35809, 2020: 47201, 2021: 66112, 2022: 98870, 2023: 59721 },
  2024: { 0: 10342, 2017: 8431, 2018: 30322, 2019: 34058, 2020: 44758, 2021: 64330, 2022: 88254, 2023: 88714, 2024: 52851 },
  2025: { 0: 15778, 2017: 7945, 2018: 29621, 2019: 33221, 2020: 45526, 2021: 64957, 2022: 89819, 2023: 82657, 2024: 96681, 2025: 70207 },
  2026: { 0: 11044.873, 2017: 7388.7291, 2018: 28732.1081, 2019: 32223.8947, 2020: 44615.3918, 2021: 65606.57, 2022: 89819.18, 2023: 83483.6811, 2024: 88946.8604, 2025: 112331.2, 2026: 65000 },
  2027: { 0: 10713.52681, 2017: 7167.067227, 2018: 27870.14486, 2019: 31257.17786, 2020: 43276.93005, 2021: 64294.4386, 2022: 90717.3718, 2023: 83483.6811, 2024: 89836.329, 2025: 103344.704, 2026: 104000, 2027: 70000 },
  2028: { 0: 10392.12101, 2017: 6952.05521, 2018: 27034.04051, 2019: 30319.46252, 2020: 41978.62214, 2021: 62365.60544, 2022: 88903.02436, 2023: 84318.51791, 2024: 89836.329, 2025: 104378.151, 2026: 95680, 2027: 112000, 2028: 75000 },
  2029: { 0: 10080.35738, 2017: 6743.493554, 2018: 26223.0193, 2019: 29409.87865, 2020: 40719.26348, 2021: 60494.63728, 2022: 86235.93363, 2023: 82632.14755, 2024: 90734.69229, 2025: 104378.151, 2026: 96636.8, 2027: 103040, 2028: 120000, 2029: 80000 },
};

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
  cohortCharts: {},
  outputCharts: {},
  syncingCharts: false,
  syncingCohortCharts: false,
  undoStack: [],
  redoStack: [],
  pendingDialogAction: null,
  cohortEditorMode: "perCohort",
  cohortEditorCohortYear: 2025,
  yearEditorYear: 2026,
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function makeBaseScenario(name = "Finance Base Case") {
  const scenario = {
    id: "base",
    name,
    drivers: clone(baseDrivers),
    newCustomerSource: "topDown",
    newCustomerDrilldown: createDefaultNewCustomerDrilldown(),
  };
  addScenarioVersion(scenario, "Initial");
  return scenario;
}

function loadScenarios() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (parsed && parsed.base && parsed.base.drivers) {
      Object.values(parsed).forEach(normalizeScenario);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(parsed));
      return parsed;
    }
  } catch (error) {
    console.warn("Could not load saved scenarios", error);
  }
  return { base: makeBaseScenario() };
}

function normalizeScenario(scenario) {
  if (!scenario.newCustomerSource) scenario.newCustomerSource = "topDown";
  if (!scenario.newCustomerDrilldown || !scenario.newCustomerDrilldown.counts) {
    scenario.newCustomerDrilldown = createDefaultNewCustomerDrilldown();
  }
  if (!scenario.newCustomerDrilldown.controlPoints) {
    scenario.newCustomerDrilldown.controlPoints = {};
  }
  enforceNewCustomerBaseline(scenario);
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
    newCustomerSource: scenario.newCustomerSource,
    newCustomerDrilldown: clone(scenario.newCustomerDrilldown),
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

function enforceNewCustomerBaseline(scenario) {
  const baselineIndex = YEARS.indexOf(NEW_CUSTOMER_BASELINE_YEAR);
  if (baselineIndex >= 0 && scenario.drivers?.newCustomers) {
    scenario.drivers.newCustomers[baselineIndex] = NEW_CUSTOMER_BASELINE_VALUE;
  }
  if (scenario.newCustomerDrilldown?.counts) {
    scenario.newCustomerDrilldown.counts[String(NEW_CUSTOMER_BASELINE_YEAR)] = clone(baseNewCustomerCohortCounts[NEW_CUSTOMER_BASELINE_YEAR]);
  }
}

function createDefaultNewCustomerDrilldown() {
  return { counts: clone(baseNewCustomerCohortCounts), controlPoints: {} };
}

function newCustomerCohortTotal(counts, year) {
  const row = counts[String(year)] || counts[year] || {};
  return NEW_CUSTOMER_COHORT_YEARS.reduce((sum, cohortYear) => {
    return sum + Number(row[String(cohortYear)] || row[cohortYear] || 0);
  }, 0);
}

function newCustomerCohortValue(counts, year, cohortYear) {
  return Number((counts[String(year)] || {})[String(cohortYear)] || 0);
}

function setNewCustomerCohortValue(counts, year, cohortYear, value) {
  if (!counts[String(year)]) counts[String(year)] = {};
  counts[String(year)][String(cohortYear)] = value;
}

function newCustomerCohortYoy(counts, year, cohortYear) {
  const current = newCustomerCohortValue(counts, year, cohortYear);
  const prior = newCustomerCohortValue(counts, year - 1, cohortYear);
  if (!prior) return null;
  return (current / prior) - 1;
}

function newCustomerCohortApplies(cohortYear, year) {
  return cohortYear === 0 || cohortYear <= year;
}

function cohortLabel(cohortYear) {
  return cohortYear === 0 ? "Legacy / Unknown" : String(cohortYear);
}

function selectedNewCustomerCohortYear() {
  if (!NEW_CUSTOMER_COHORT_YEARS.includes(state.cohortEditorCohortYear)) {
    state.cohortEditorCohortYear = 2025;
  }
  return state.cohortEditorCohortYear;
}

function cohortEditorSliderYears() {
  return NEW_CUSTOMER_COHORT_YEARS.slice().reverse();
}

function selectedNewCustomerYearEditorYear() {
  if (!YEARS.includes(state.yearEditorYear) || !editableYear(state.yearEditorYear)) {
    state.yearEditorYear = FIRST_FORECAST_YEAR;
  }
  return state.yearEditorYear;
}

function yearEditorSliderYears() {
  return YEARS.filter(editableYear);
}

function existingPropertyCohortsForYear(year) {
  return NEW_CUSTOMER_COHORT_YEARS.filter(cohortYear => {
    return newCustomerCohortApplies(cohortYear, year) && cohortYear !== year;
  });
}

function newPropertyCohortValue(counts, year) {
  return newCustomerCohortValue(counts, year, year);
}

function newPropertyCohortYoy(counts, year) {
  const current = newPropertyCohortValue(counts, year);
  const prior = newPropertyCohortValue(counts, year - 1);
  if (!prior) return null;
  return (current / prior) - 1;
}

function napkinControlKey(type, ...parts) {
  return [type, ...parts].join(":");
}

function controlledNapkinPairs(key, pairs) {
  const stored = activeScenario().newCustomerDrilldown.controlPoints?.[key];
  if (!Array.isArray(stored)) return pairs;
  const visibleX = new Set(stored.map(Number));
  return pairs.filter(([x]) => visibleX.has(Number(x)));
}

function rememberNapkinControlPoints(key, data) {
  if (!activeScenario().newCustomerDrilldown.controlPoints) {
    activeScenario().newCustomerDrilldown.controlPoints = {};
  }
  activeScenario().newCustomerDrilldown.controlPoints[key] = (data || [])
    .map(([x]) => Number(x))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
}

function addNapkinControlPoint(key, x) {
  if (!Number.isFinite(Number(x))) return;
  if (!activeScenario().newCustomerDrilldown.controlPoints) {
    activeScenario().newCustomerDrilldown.controlPoints = {};
  }
  const existing = activeScenario().newCustomerDrilldown.controlPoints[key];
  const next = new Set(Array.isArray(existing) ? existing.map(Number) : []);
  next.add(Number(x));
  activeScenario().newCustomerDrilldown.controlPoints[key] = Array.from(next).sort((left, right) => left - right);
}

function changedNapkinXs(beforeData, afterData) {
  const before = new Map((beforeData || []).map(([x, y]) => [Number(x), Number(y)]));
  const changed = [];
  (afterData || []).forEach(([x, y]) => {
    const numericX = Number(x);
    const numericY = Number(y);
    if (!Number.isFinite(numericX) || !Number.isFinite(numericY)) return;
    if (!before.has(numericX) || before.get(numericX) !== numericY) {
      changed.push(numericX);
    }
  });
  return changed;
}

function yearEditorXForCohort(cohorts, cohortYear) {
  return cohorts.indexOf(cohortYear);
}

function yearEditorCohortForX(cohorts, x) {
  const index = Math.round(Number(x));
  if (index < 0 || index >= cohorts.length) return null;
  return cohorts[index];
}

function applyNewCustomerCohortYoy(counts, year, cohortYear, yoy) {
  const preservedFutureYoy = new Map();
  YEARS.forEach(candidateYear => {
    if (candidateYear > year) {
      preservedFutureYoy.set(candidateYear, newCustomerCohortYoy(counts, candidateYear, cohortYear));
    }
  });

  const prior = newCustomerCohortValue(counts, year - 1, cohortYear);
  setNewCustomerCohortValue(counts, year, cohortYear, prior * (1 + yoy));

  YEARS.forEach(candidateYear => {
    const futureYoy = preservedFutureYoy.get(candidateYear);
    const futurePrior = newCustomerCohortValue(counts, candidateYear - 1, cohortYear);
    if (candidateYear > year && futureYoy !== null && futurePrior) {
      setNewCustomerCohortValue(counts, candidateYear, cohortYear, futurePrior * (1 + futureYoy));
    }
  });
}

function bottomUpNewCustomers(scenario) {
  const counts = scenario.newCustomerDrilldown?.counts || {};
  return YEARS.map(year => newCustomerCohortTotal(counts, year));
}

function effectiveNewCustomers(scenario) {
  return scenario.newCustomerSource === "bottomUp"
    ? bottomUpNewCustomers(scenario)
    : scenario.drivers.newCustomers;
}

function driverValues(scenario, key) {
  return key === "newCustomers" ? effectiveNewCustomers(scenario) : scenario.drivers[key];
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
  const newCustomerValues = effectiveNewCustomers(scenario);
  return calculateOutputsWithNewCustomers(scenario, newCustomerValues);
}

function calculateOutputsWithNewCustomers(scenario, newCustomerValues) {
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
    const newCust = newCustomerValues[index];
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
  const lineIsEditable = !(key === "newCustomers" && scenario.newCustomerSource === "bottomUp");
  return {
    name: scenario.name,
    color: meta.color,
    editable: lineIsEditable,
    editDomain: lineIsEditable
      ? {
        moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      }
      : undefined,
    data: chartPairs(driverValues(scenario, key), meta),
  };
}

function makeComparisonLine(key, scenario) {
  const meta = driverMeta[key];
  return {
    name: `Comparison: ${scenario.name}`,
    color: "#98a2b3",
    editable: false,
    data: chartPairs(driverValues(scenario, key), meta),
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
      if (key === "newCustomers" && activeScenario().newCustomerSource === "bottomUp") return;
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
  state.outputCharts.newCustomerDrilldown = echarts.init(document.getElementById("new-customers-drilldown-chart"));
  state.outputCharts.newCustomerTotal = echarts.init(document.getElementById("new-customers-total-chart"));
  window.addEventListener("resize", () => {
    state.outputCharts.revenue.resize();
    state.outputCharts.gap.resize();
    state.outputCharts.newCustomerDrilldown.resize();
    state.outputCharts.newCustomerTotal.resize();
  });
}

function initNewCustomerCohortEditors() {
  const countChart = new NapkinChart(
    "new-customers-cohort-count-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 150000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => Math.round(value).toLocaleString("en-US")),
      },
    },
    "none",
    false
  );
  const yoyChart = new NapkinChart(
    "new-customers-cohort-yoy-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: -50, max: 100, axisLabel: { formatter: value => `${trimNumber(value, 0)}%` } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => `${trimNumber(value, 1)}%`),
      },
    },
    "none",
    false
  );
  const yearChart = new NapkinChart(
    "new-customers-year-existing-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: 0, max: 10, minInterval: 1, axisLabel: { formatter: value => formatYearEditorAxisLabel(value) } },
      yAxis: { type: "value", min: 0, max: 150000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 48, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatYearEditorTooltip(params),
      },
    },
    "none",
    false
  );
  const yearYoyChart = new NapkinChart(
    "new-customers-year-existing-yoy-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: 0, max: 10, minInterval: 1, axisLabel: { formatter: value => formatYearEditorAxisLabel(value) } },
      yAxis: { type: "value", min: -50, max: 100, axisLabel: { formatter: value => `${trimNumber(value, 0)}%` } },
      grid: { left: 12, right: 18, top: 14, bottom: 48, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatYearEditorTooltip(params, value => `${trimNumber(value, 1)}%`),
      },
    },
    "none",
    false
  );
  const newPropertiesCountChart = new NapkinChart(
    "new-customers-new-properties-count-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 150000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => Math.round(value).toLocaleString("en-US")),
      },
    },
    "none",
    false
  );
  const newPropertiesYoyChart = new NapkinChart(
    "new-customers-new-properties-yoy-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: -50, max: 100, axisLabel: { formatter: value => `${trimNumber(value, 0)}%` } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => `${trimNumber(value, 1)}%`),
      },
    },
    "none",
    false
  );

  [countChart, yoyChart, yearChart, yearYoyChart, newPropertiesCountChart, newPropertiesYoyChart].forEach(chart => {
    chart.windowStartX = YEARS[0];
    chart.windowEndX = YEARS[YEARS.length - 1];
    chart.globalMaxX = YEARS[YEARS.length - 1];
    chart._appEditSnapshot = null;
    chart._appEditCommitted = false;
    chart.chart.getZr().on("mousedown", () => {
      if (state.syncingCohortCharts) return;
      chart._appEditSnapshot = snapshotState();
      chart._appEditLineSnapshot = clone(chart.lines?.[0]?.data || []);
      chart._appEditCommitted = false;
    });
    chart.chart.getZr().on("mousemove", () => highlightCellsForNapkinDrag(chart));
    chart.chart.on("mouseover", params => highlightCellsForNapkinPoint(chart, params));
    chart.chart.on("mouseout", clearNewCustomerCellHighlights);
    chart.chart.getZr().on("globalout", clearNewCustomerCellHighlights);
  });

  countChart.onDataChanged = () => handleCohortCountChartChange(countChart);
  yoyChart.onDataChanged = () => handleCohortYoyChartChange(yoyChart);
  yearChart.onDataChanged = () => handleYearExistingChartChange(yearChart);
  yearYoyChart.onDataChanged = () => handleYearExistingYoyChartChange(yearYoyChart);
  newPropertiesCountChart.onDataChanged = () => handleNewPropertiesCountChartChange(newPropertiesCountChart);
  newPropertiesYoyChart.onDataChanged = () => handleNewPropertiesYoyChartChange(newPropertiesYoyChart);
  state.cohortCharts.count = countChart;
  state.cohortCharts.yoy = yoyChart;
  state.cohortCharts.yearExisting = yearChart;
  state.cohortCharts.yearExistingYoy = yearYoyChart;
  state.cohortCharts.newPropertiesCount = newPropertiesCountChart;
  state.cohortCharts.newPropertiesYoy = newPropertiesYoyChart;
  countChart._appHighlightMapper = point => ({ countYear: Number(point[0]), countCohortYear: selectedNewCustomerCohortYear() });
  yoyChart._appHighlightMapper = point => ({ countYear: Number(point[0]), countCohortYear: selectedNewCustomerCohortYear(), yoyYear: Number(point[0]), yoyCohortYear: selectedNewCustomerCohortYear() });
  yearChart._appHighlightMapper = (point, index) => {
    const year = selectedNewCustomerYearEditorYear();
    const cohortYear = existingPropertyCohortsForYear(year)[index];
    return { countYear: year, countCohortYear: cohortYear };
  };
  yearYoyChart._appHighlightMapper = (point, index) => {
    const year = selectedNewCustomerYearEditorYear();
    const cohortYear = existingPropertyCohortsForYear(year)[index];
    return { countYear: year, countCohortYear: cohortYear, yoyYear: year, yoyCohortYear: cohortYear };
  };
  newPropertiesCountChart._appHighlightMapper = point => ({ countYear: Number(point[0]), countCohortYear: Number(point[0]) });
  newPropertiesYoyChart._appHighlightMapper = point => ({ countYear: Number(point[0]), countCohortYear: Number(point[0]) });
}

function highlightCellsForNapkinPoint(chart, params) {
  if (!chart?._appHighlightMapper || !params || params.seriesType !== "line") return;
  const lineIndex = Math.floor(Number(params.seriesIndex) / 2);
  if (lineIndex !== 0) return;
  const pointIndex = Number(params.dataIndex);
  const point = chart.lines?.[0]?.data?.[pointIndex];
  if (!point) return;
  highlightNewCustomerCells(chart._appHighlightMapper(point, pointIndex));
}

function highlightCellsForNapkinDrag(chart) {
  if (!chart?._appHighlightMapper || chart._draggingLineIndex !== 0 || chart._draggingPointIndex === null) return;
  const pointIndex = Number(chart._draggingPointIndex);
  const point = chart.lines?.[0]?.data?.[pointIndex];
  if (!point) return;
  highlightNewCustomerCells(chart._appHighlightMapper(point, pointIndex));
}

function formatCohortNapkinTooltip(params, formatter) {
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
      rows.push(`${item.marker || ""} ${item.seriesName}: ${formatter(Number(rawValue))}`);
    });
  return [tooltipHeader(year), ...rows].join("<br/>");
}

function formatYearEditorAxisLabel(value) {
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const cohortYear = yearEditorCohortForX(cohorts, value);
  return cohortYear === null ? "" : cohortLabel(cohortYear);
}

function formatYearEditorTooltip(params, formatter = value => Math.round(value).toLocaleString("en-US")) {
  const items = Array.isArray(params) ? params : [params];
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const rows = [];
  const seen = new Set();
  items
    .filter(item => item && item.seriesType === "line" && item.seriesIndex % 2 === 1)
    .forEach(item => {
      if (seen.has(item.seriesName)) return;
      seen.add(item.seriesName);
      const data = Array.isArray(item.data) ? item.data : item.value;
      const x = Array.isArray(data) ? data[0] : item.axisValue;
      const y = Array.isArray(data) ? data[1] : item.value;
      const cohortYear = yearEditorCohortForX(cohorts, x);
      if (cohortYear !== null) {
        rows.push(`${item.marker || ""} ${item.seriesName} ${cohortLabel(cohortYear)}: ${formatter(Number(y))}`);
      }
    });
  return [tooltipHeader(year), ...rows].join("<br/>");
}

function renderOutputCharts(outputs) {
  state.outputCharts.revenue.setOption(revenuePathChartOption(outputs), true);

  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
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

function revenuePathChartOption(outputs) {
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

  return {
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
  };
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
    values: driverValues(activeScenario(), key),
    editable: !(key === "newCustomers" && activeScenario().newCustomerSource === "bottomUp"),
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

function renderNewCustomerDrilldown() {
  renderNewCustomerSourceControl();
  renderNewCustomerEditorMode();
  renderNewCustomerCohortEditorControl();
  renderNewCustomerYearEditorControl();
  renderNewCustomerDrilldownCharts();
  syncNewCustomerCohortEditors();
  renderNewCustomerDrilldownTable();
  renderNewCustomerYoyTable();
}

function renderNewCustomerSourceControl() {
  const sourceSelect = document.getElementById("new-customers-source");
  if (sourceSelect) sourceSelect.value = activeScenario().newCustomerSource || "topDown";
}

function renderNewCustomerCohortEditorControl() {
  const slider = document.getElementById("new-customers-cohort-editor-slider");
  const value = document.getElementById("new-customers-cohort-editor-value");
  if (!slider || !value) return;
  const selected = selectedNewCustomerCohortYear();
  const sliderYears = cohortEditorSliderYears();
  slider.max = String(sliderYears.length - 1);
  slider.value = String(Math.max(0, sliderYears.indexOf(selected)));
  value.value = cohortLabel(selected);
}

function renderNewCustomerYearEditorControl() {
  const slider = document.getElementById("new-customers-year-editor-slider");
  const value = document.getElementById("new-customers-year-editor-value");
  if (!slider || !value) return;
  const selected = selectedNewCustomerYearEditorYear();
  const sliderYears = yearEditorSliderYears();
  slider.max = String(sliderYears.length - 1);
  slider.value = String(Math.max(0, sliderYears.indexOf(selected)));
  value.value = String(selected);
}

function renderNewCustomerEditorMode() {
  document.getElementById("new-customers-editor-per-cohort").classList.toggle("active", state.cohortEditorMode === "perCohort");
  document.getElementById("new-customers-editor-per-year").classList.toggle("active", state.cohortEditorMode === "perYear");
  document.getElementById("new-customers-editor-new-properties").classList.toggle("active", state.cohortEditorMode === "newProperties");
  document.getElementById("new-customers-per-cohort-editor").classList.toggle("active", state.cohortEditorMode === "perCohort");
  document.getElementById("new-customers-per-year-editor").classList.toggle("active", state.cohortEditorMode === "perYear");
  document.getElementById("new-customers-new-properties-editor").classList.toggle("active", state.cohortEditorMode === "newProperties");
}

function renderNewCustomerDrilldownCharts() {
  const comparison = compareScenario();
  const comparisonTotals = comparison ? bottomUpNewCustomers(comparison) : null;
  const activeTotals = bottomUpNewCustomers(activeScenario());
  const topDownOutputs = calculateOutputsWithNewCustomers(activeScenario(), activeScenario().drivers.newCustomers);
  const bottomUpOutputs = calculateOutputsWithNewCustomers(activeScenario(), activeTotals);
  state.outputCharts.newCustomerDrilldown.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${item.seriesName}: ${formatCurrency(Number(item.value), 0)}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
    series: [
      { name: "Plan Revenue", type: "line", data: planRevenue, symbolSize: 6, lineStyle: { color: "#111827", width: 2 }, itemStyle: { color: "#111827" } },
      { name: `${activeScenario().name} Top-Down`, type: "line", data: topDownOutputs.revenue, symbolSize: 5, lineStyle: { color: TOP_DOWN_COLOR, width: 2 }, itemStyle: { color: TOP_DOWN_COLOR } },
      { name: `${activeScenario().name} Bottom-Up`, type: "line", data: bottomUpOutputs.revenue, symbolSize: 5, lineStyle: { color: BOTTOM_UP_COLOR, width: 2 }, itemStyle: { color: BOTTOM_UP_COLOR } },
    ],
  }, true);

  const totalSeries = [
    { name: `${activeScenario().name} Top-Down`, type: "line", data: activeScenario().drivers.newCustomers, symbolSize: 5, lineStyle: { color: TOP_DOWN_COLOR, width: 2 }, itemStyle: { color: TOP_DOWN_COLOR } },
    { name: `${activeScenario().name} Bottom-Up`, type: "line", data: activeTotals, symbolSize: 5, z: 3, lineStyle: { color: BOTTOM_UP_COLOR, width: 2 }, itemStyle: { color: BOTTOM_UP_COLOR } },
  ];
  if (comparisonTotals) {
    totalSeries.unshift({ name: `${comparison.name} Bottom-Up`, type: "line", data: comparisonTotals, symbolSize: 5, z: 2, lineStyle: { color: "#98a2b3", width: 3 }, itemStyle: { color: "#98a2b3" } });
  }
  state.outputCharts.newCustomerTotal.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${item.seriesName}: ${Math.round(Number(item.value)).toLocaleString("en-US")}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 58, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: totalSeries,
  }, true);

}

function cohortCountChartLines() {
  const cohortYear = selectedNewCustomerCohortYear();
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("cohortCount", cohortYear);
  const defaultPairs = YEARS
    .filter(year => newCustomerCohortApplies(cohortYear, year))
    .map(year => [year, newCustomerCohortValue(activeCounts, year, cohortYear)]);
  const lines = [{
    name: `${cohortLabel(cohortYear)} Cohort`,
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNapkinPairs(controlKey, defaultPairs),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${comparison.name}`,
      color: "#98a2b3",
      editable: false,
      data: YEARS
        .filter(year => newCustomerCohortApplies(cohortYear, year))
        .map(year => [year, newCustomerCohortValue(comparison.newCustomerDrilldown.counts, year, cohortYear)]),
    });
  }
  return lines;
}

function cohortYoyChartLines() {
  const cohortYear = selectedNewCustomerCohortYear();
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("cohortYoy", cohortYear);
  const defaultPairs = YEARS
    .map(year => [year, newCustomerCohortYoy(activeCounts, year, cohortYear)])
    .filter(([, value]) => value !== null)
    .map(([year, value]) => [year, value * 100]);
  const lines = [{
    name: `${cohortLabel(cohortYear)} YoY`,
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNapkinPairs(controlKey, defaultPairs),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${comparison.name}`,
      color: "#98a2b3",
      editable: false,
      data: YEARS
        .map(year => [year, newCustomerCohortYoy(comparison.newCustomerDrilldown.counts, year, cohortYear)])
        .filter(([, value]) => value !== null)
        .map(([year, value]) => [year, value * 100]),
    });
  }
  return lines;
}

function yearExistingChartLines() {
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("yearExisting", year);
  const defaultPairs = cohorts.map(cohortYear => [
    yearEditorXForCohort(cohorts, cohortYear),
    newCustomerCohortValue(activeCounts, year, cohortYear),
  ]);
  const lines = [{
    name: `${year} Existing`,
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[0, Math.max(0, cohorts.length - 1)]],
      addX: [],
      deleteX: [[0, Math.max(0, cohorts.length - 1)]],
    },
    data: controlledNapkinPairs(controlKey, defaultPairs),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${comparison.name}`,
      color: "#98a2b3",
      editable: false,
      data: cohorts.map(cohortYear => [
        yearEditorXForCohort(cohorts, cohortYear),
        newCustomerCohortValue(comparison.newCustomerDrilldown.counts, year, cohortYear),
      ]),
    });
  }
  return lines;
}

function yearExistingYoyChartLines() {
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("yearExistingYoy", year);
  const defaultPairs = cohorts.map(cohortYear => [
    yearEditorXForCohort(cohorts, cohortYear),
    (newCustomerCohortYoy(activeCounts, year, cohortYear) || 0) * 100,
  ]);
  const lines = [{
    name: `${year} Existing YoY`,
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[0, Math.max(0, cohorts.length - 1)]],
      addX: [],
      deleteX: [[0, Math.max(0, cohorts.length - 1)]],
    },
    data: controlledNapkinPairs(controlKey, defaultPairs),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${comparison.name}`,
      color: "#98a2b3",
      editable: false,
      data: cohorts.map(cohortYear => [
        yearEditorXForCohort(cohorts, cohortYear),
        (newCustomerCohortYoy(comparison.newCustomerDrilldown.counts, year, cohortYear) || 0) * 100,
      ]),
    });
  }
  return lines;
}

function newPropertiesCountChartLines() {
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("newPropertiesCount");
  const defaultPairs = YEARS.map(year => [year, newPropertyCohortValue(activeCounts, year)]);
  const lines = [{
    name: "New Properties",
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNapkinPairs(controlKey, defaultPairs),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${comparison.name}`,
      color: "#98a2b3",
      editable: false,
      data: YEARS.map(year => [year, newPropertyCohortValue(comparison.newCustomerDrilldown.counts, year)]),
    });
  }
  return lines;
}

function newPropertiesYoyChartLines() {
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("newPropertiesYoy");
  const defaultPairs = YEARS
    .map(year => [year, newPropertyCohortYoy(activeCounts, year)])
    .filter(([, value]) => value !== null)
    .map(([year, value]) => [year, value * 100]);
  const lines = [{
    name: "New Properties YoY",
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNapkinPairs(controlKey, defaultPairs),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${comparison.name}`,
      color: "#98a2b3",
      editable: false,
      data: YEARS
        .map(year => [year, newPropertyCohortYoy(comparison.newCustomerDrilldown.counts, year)])
        .filter(([, value]) => value !== null)
        .map(([year, value]) => [year, value * 100]),
    });
  }
  return lines;
}

function syncNewCustomerCohortEditors() {
  state.syncingCohortCharts = true;
  if (state.cohortCharts.count) {
    state.cohortCharts.count.lines = cohortCountChartLines();
    state.cohortCharts.count._refreshChart();
    styleComparisonSeries(state.cohortCharts.count);
  }
  if (state.cohortCharts.yoy) {
    state.cohortCharts.yoy.lines = cohortYoyChartLines();
    state.cohortCharts.yoy._refreshChart();
    styleComparisonSeries(state.cohortCharts.yoy);
  }
  if (state.cohortCharts.yearExisting) {
    const year = selectedNewCustomerYearEditorYear();
    const cohorts = existingPropertyCohortsForYear(year);
    syncYearEditorChart(state.cohortCharts.yearExisting, cohorts, yearExistingChartLines());
  }
  if (state.cohortCharts.yearExistingYoy) {
    const year = selectedNewCustomerYearEditorYear();
    const cohorts = existingPropertyCohortsForYear(year);
    syncYearEditorChart(state.cohortCharts.yearExistingYoy, cohorts, yearExistingYoyChartLines());
  }
  if (state.cohortCharts.newPropertiesCount) {
    state.cohortCharts.newPropertiesCount.lines = newPropertiesCountChartLines();
    state.cohortCharts.newPropertiesCount._refreshChart();
    styleComparisonSeries(state.cohortCharts.newPropertiesCount);
  }
  if (state.cohortCharts.newPropertiesYoy) {
    state.cohortCharts.newPropertiesYoy.lines = newPropertiesYoyChartLines();
    state.cohortCharts.newPropertiesYoy._refreshChart();
    styleComparisonSeries(state.cohortCharts.newPropertiesYoy);
  }
  state.syncingCohortCharts = false;
}

function syncYearEditorChart(chart, cohorts, lines) {
  chart.baseOption.xAxis.min = 0;
  chart.baseOption.xAxis.max = Math.max(0, cohorts.length - 1);
  chart.windowStartX = 0;
  chart.windowEndX = Math.max(0, cohorts.length - 1);
  chart.globalMaxX = Math.max(0, cohorts.length - 1);
  chart.lines = lines;
  chart.resize();
  chart._refreshChart();
  chart.resize();
  styleComparisonSeries(chart);
}

function handleCohortCountChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const cohortYear = selectedNewCustomerCohortYear();
  const changedXs = changedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("cohortCount", cohortYear), chart.lines[0].data);
  changedXs.forEach(year => addNapkinControlPoint(napkinControlKey("cohortYoy", cohortYear), year));
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && newCustomerCohortApplies(cohortYear, year) && value !== null) {
      setNewCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, year, cohortYear, Math.max(0, value));
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function handleCohortYoyChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const cohortYear = selectedNewCustomerCohortYear();
  const changedXs = changedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("cohortYoy", cohortYear), chart.lines[0].data);
  changedXs.forEach(year => addNapkinControlPoint(napkinControlKey("cohortCount", cohortYear), year));
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      applyNewCustomerCohortYoy(activeScenario().newCustomerDrilldown.counts, year, cohortYear, value / 100);
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function interpolateNapkinLineValue(data, x) {
  const points = (data || [])
    .filter(point => Array.isArray(point) && point.length >= 2)
    .map(([pointX, pointY]) => [Number(pointX), Number(pointY)])
    .filter(([pointX, pointY]) => Number.isFinite(pointX) && Number.isFinite(pointY))
    .sort((left, right) => left[0] - right[0]);
  if (!points.length) return null;
  if (x <= points[0][0]) return points[0][1];
  if (x >= points[points.length - 1][0]) return points[points.length - 1][1];
  for (let index = 1; index < points.length; index += 1) {
    const [rightX, rightY] = points[index];
    const [leftX, leftY] = points[index - 1];
    if (x === rightX) return rightY;
    if (x >= leftX && x <= rightX) {
      const position = (x - leftX) / (rightX - leftX);
      return leftY + ((rightY - leftY) * position);
    }
  }
  return null;
}

function handleYearExistingChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const changedXs = changedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("yearExisting", year), chart.lines[0].data);
  changedXs.forEach(x => addNapkinControlPoint(napkinControlKey("yearExistingYoy", year), x));
  cohorts.forEach((cohortYear, index) => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, index);
    if (cohortYear !== null && value !== null) {
      setNewCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, year, cohortYear, Math.max(0, value));
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function handleYearExistingYoyChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const changedXs = changedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("yearExistingYoy", year), chart.lines[0].data);
  changedXs.forEach(x => addNapkinControlPoint(napkinControlKey("yearExisting", year), x));
  cohorts.forEach((cohortYear, index) => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, index);
    if (cohortYear !== null && value !== null) {
      applyNewCustomerCohortYoy(activeScenario().newCustomerDrilldown.counts, year, cohortYear, value / 100);
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function handleNewPropertiesCountChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const changedXs = changedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("newPropertiesCount"), chart.lines[0].data);
  changedXs.forEach(year => addNapkinControlPoint(napkinControlKey("newPropertiesYoy"), year));
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, year, year, Math.max(0, value));
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function handleNewPropertiesYoyChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const changedXs = changedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("newPropertiesYoy"), chart.lines[0].data);
  changedXs.forEach(year => addNapkinControlPoint(napkinControlKey("newPropertiesCount"), year));
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    const prior = newPropertyCohortValue(activeScenario().newCustomerDrilldown.counts, year - 1);
    if (editableYear(year) && prior && value !== null) {
      setNewCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, year, year, prior * (1 + (value / 100)));
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAll();
}

function renderNewCustomerDrilldownTable() {
  const table = document.getElementById("new-customers-drilldown-table");
  const counts = activeScenario().newCustomerDrilldown.counts;
  const rows = NEW_CUSTOMER_COHORT_YEARS.slice().reverse().map(cohortYear => ({
    cohortYear,
    values: YEARS.map(year => {
      if (!newCustomerCohortApplies(cohortYear, year)) return null;
      return Number((counts[String(year)] || {})[String(cohortYear)] || 0);
    }),
  }));
  const totals = YEARS.map(year => newCustomerCohortTotal(counts, year));
  table.innerHTML = `
    <thead>
      <tr>
        <th>Property Cohort</th>
        ${YEARS.map(year => `<th>${year}</th>`).join("")}
      </tr>
    </thead>
    <tbody>
      ${rows.map(row => `
        <tr>
          <td>${row.cohortYear === 0 ? "Legacy / Unknown" : row.cohortYear}</td>
          ${YEARS.map((year, index) => renderNewCustomerCohortCell(row.cohortYear, year, row.values[index])).join("")}
        </tr>
      `).join("")}
      <tr>
        <td><strong>Total New Customers</strong></td>
        ${totals.map(total => `<td class="output">${Math.round(total).toLocaleString("en-US")}</td>`).join("")}
      </tr>
    </tbody>
  `;
  table.querySelectorAll("input[data-cohort-year]").forEach(input => {
    input.addEventListener("change", event => {
      const year = event.target.dataset.year;
      const cohortYear = event.target.dataset.cohortYear;
      const parsed = Number(String(event.target.value).replace(/[,\s]/g, ""));
      if (!Number.isFinite(parsed)) {
        renderNewCustomerDrilldownTable();
        return;
      }
      pushUndoSnapshot();
      setNewCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, Number(year), Number(cohortYear), parsed);
      saveScenarios();
      syncDriverCharts();
      renderAll();
    });
  });
}

function renderNewCustomerCohortCell(cohortYear, year, value) {
  const attrs = `data-count-year="${year}" data-count-cohort-year="${cohortYear}"`;
  if (value === null) return `<td class="output" ${attrs}>-</td>`;
  const formatted = Math.round(value).toLocaleString("en-US");
  if (editableYear(year)) {
    return `
      <td class="input" ${attrs}>
        <input data-year="${year}" data-cohort-year="${cohortYear}" value="${formatted}" aria-label="${cohortYear} cohort ${year} new customers" />
      </td>
    `;
  }
  return `<td class="historical" ${attrs}>${formatted}</td>`;
}

function renderNewCustomerYoyTable() {
  const table = document.getElementById("new-customers-yoy-table");
  const counts = activeScenario().newCustomerDrilldown.counts;
  const rows = NEW_CUSTOMER_COHORT_YEARS.slice().reverse().map(cohortYear => ({
    cohortYear,
    values: YEARS.map(year => newCustomerCohortYoy(counts, year, cohortYear)),
  }));
  table.innerHTML = `
    <thead>
      <tr>
        <th>Cohort YoY % Diff</th>
        ${YEARS.map(year => `<th>${year}</th>`).join("")}
      </tr>
    </thead>
    <tbody>
      ${rows.map(row => `
        <tr>
          <td>${row.cohortYear === 0 ? "Legacy / Unknown" : row.cohortYear}</td>
          ${YEARS.map((year, index) => renderNewCustomerYoyCell(row.cohortYear, year, row.values[index])).join("")}
        </tr>
      `).join("")}
    </tbody>
  `;
  table.querySelectorAll("input[data-yoy-cohort-year]").forEach(input => {
    input.addEventListener("focus", event => {
      highlightNewCustomerCells({
        countYear: event.target.dataset.year,
        countCohortYear: event.target.dataset.yoyCohortYear,
        yoyYear: event.target.dataset.year,
        yoyCohortYear: event.target.dataset.yoyCohortYear,
      });
    });
    input.addEventListener("blur", clearNewCustomerCellHighlights);
    input.addEventListener("change", event => {
      const year = Number(event.target.dataset.year);
      const cohortYear = Number(event.target.dataset.yoyCohortYear);
      const before = snapshotState();
      const updated = updateNewCustomerCohortYoy(year, cohortYear, event.target.value, before);
      if (!updated) renderNewCustomerYoyTable();
    });
  });
}

function renderNewCustomerYoyCell(cohortYear, year, value) {
  const attrs = `data-yoy-year="${year}" data-yoy-cohort-year="${cohortYear}"`;
  if (value === null) return `<td class="output" ${attrs}>-</td>`;
  const formatted = formatValue(value, "percent");
  if (editableYear(year)) {
    return `
      <td class="input" ${attrs}>
        <input data-year="${year}" data-yoy-cohort-year="${cohortYear}" value="${formatted}" aria-label="${cohortYear} cohort ${year} new customer YoY percent diff" />
      </td>
    `;
  }
  return `<td class="historical" ${attrs}>${formatted}</td>`;
}

function highlightNewCustomerCells({ countYear, countCohortYear, yoyYear, yoyCohortYear } = {}) {
  clearNewCustomerCellHighlights();
  if (countYear !== undefined && countCohortYear !== undefined) {
    document
      .querySelectorAll(`[data-count-year="${countYear}"][data-count-cohort-year="${countCohortYear}"]`)
      .forEach(cell => cell.classList.add("linked-highlight"));
  }
  if (yoyYear !== undefined && yoyCohortYear !== undefined) {
    document
      .querySelectorAll(`[data-yoy-year="${yoyYear}"][data-yoy-cohort-year="${yoyCohortYear}"]`)
      .forEach(cell => cell.classList.add("linked-highlight"));
  }
}

function clearNewCustomerCellHighlights() {
  document
    .querySelectorAll(".linked-highlight")
    .forEach(cell => cell.classList.remove("linked-highlight"));
}

function updateNewCustomerCohortYoy(year, cohortYear, rawValue, undoSnapshot = snapshotState()) {
  const parsed = parseInput(rawValue, "percent");
  const prior = newCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, year - 1, cohortYear);
  if (parsed === null || !prior) return false;
  pushUndoSnapshot(undoSnapshot);
  applyNewCustomerCohortYoy(activeScenario().newCustomerDrilldown.counts, year, cohortYear, parsed);
  saveScenarios();
  syncDriverCharts();
  renderAll();
  return true;
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
  renderNewCustomerDrilldown();
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
    newCustomerSource: activeScenario().newCustomerSource,
    newCustomerDrilldown: clone(activeScenario().newCustomerDrilldown),
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
  scenario.newCustomerSource = version.newCustomerSource || "topDown";
  scenario.newCustomerDrilldown = version.newCustomerDrilldown
    ? clone(version.newCustomerDrilldown)
    : createDefaultNewCustomerDrilldown();
  enforceNewCustomerBaseline(scenario);
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
    activeScenario().newCustomerSource = "topDown";
    activeScenario().newCustomerDrilldown = createDefaultNewCustomerDrilldown();
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
  document.getElementById("new-customers-source").addEventListener("change", event => {
    pushUndoSnapshot();
    activeScenario().newCustomerSource = event.target.value;
    saveScenarios();
    syncDriverCharts();
    renderAll();
  });
  document.getElementById("new-customers-cohort-editor-slider").addEventListener("input", event => {
    const sliderYears = cohortEditorSliderYears();
    state.cohortEditorCohortYear = sliderYears[Number(event.target.value)] ?? state.cohortEditorCohortYear;
    renderNewCustomerCohortEditorControl();
    syncNewCustomerCohortEditors();
  });
  document.getElementById("new-customers-year-editor-slider").addEventListener("input", event => {
    const sliderYears = yearEditorSliderYears();
    state.yearEditorYear = sliderYears[Number(event.target.value)] ?? state.yearEditorYear;
    renderNewCustomerYearEditorControl();
    syncNewCustomerCohortEditors();
  });
  document.getElementById("new-customers-editor-per-cohort").addEventListener("click", () => {
    state.cohortEditorMode = "perCohort";
    renderNewCustomerEditorMode();
    setTimeout(() => {
      Object.values(state.cohortCharts).forEach(chart => chart.resize());
      syncNewCustomerCohortEditors();
    }, 0);
  });
  document.getElementById("new-customers-editor-per-year").addEventListener("click", () => {
    state.cohortEditorMode = "perYear";
    renderNewCustomerEditorMode();
    setTimeout(() => {
      Object.values(state.cohortCharts).forEach(chart => chart.resize());
      syncNewCustomerCohortEditors();
    }, 0);
  });
  document.getElementById("new-customers-editor-new-properties").addEventListener("click", () => {
    state.cohortEditorMode = "newProperties";
    renderNewCustomerEditorMode();
    setTimeout(() => {
      Object.values(state.cohortCharts).forEach(chart => chart.resize());
      syncNewCustomerCohortEditors();
    }, 0);
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
  document.getElementById("tab-new-customers").addEventListener("click", () => setView("newCustomers"));
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
  document.getElementById("tab-new-customers").classList.toggle("active", view === "newCustomers");
  document.getElementById("chart-view").classList.toggle("active", view === "chart");
  document.getElementById("table-view").classList.toggle("active", view === "table");
  document.getElementById("new-customers-view").classList.toggle("active", view === "newCustomers");
  setTimeout(() => {
    Object.values(state.charts).forEach(chart => chart.resize());
    Object.values(state.cohortCharts).forEach(chart => chart.resize());
    Object.values(state.outputCharts).forEach(chart => chart.resize());
  }, 0);
}

function init() {
  renderScenarioSelect();
  bindControls();
  initOutputCharts();
  Object.keys(driverMeta).forEach(initDriverChart);
  initNewCustomerCohortEditors();
  updateHistoryControls();
  renderAll();
}

init();
