const YEARS = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const FIRST_FORECAST_YEAR = 2026;
const STORAGE_KEY = "ps-hhp-simple-scenarios-v1";
const TOP_DOWN_COLOR = "#6fa76b";
const BOTTOM_UP_COLOR = "#4f7fb8";

const planRevenue = [5546199, 10450032, 16513061, 25376000, 36444600, 46813700, 65766000, 83585600, 128908000];
const NEW_CUSTOMER_COHORT_YEARS = [0, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const NEW_CUSTOMER_BASELINE_YEAR = 2025;
const NEW_CUSTOMER_BASELINE_VALUE = 536412;
const REV_UNIT_YEARS = [2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const REV_UNIT_LAST_HISTORICAL_YEAR = 2025;
const baseRevUnitNewUnits = {
  2017: 48442,
  2018: 304637,
  2019: 376832,
  2020: 683294,
  2021: 886312,
  2022: 1225709,
  2023: 1133249,
  2024: 1112682,
  2025: 1350383,
  2026: 1350383,
  2027: 1350383,
  2028: 1350383,
  2029: 1350383,
};
const baseRevUnitNewProperties = {
  2017: 107,
  2018: 845,
  2019: 1426,
  2020: 3020,
  2021: 4172,
  2022: 6085,
  2023: 6116,
  2024: 6232,
  2025: 8009,
};
const baseRevUnitRevenue = {
  2017: 24197,
  2018: 327354,
  2019: 814808,
  2020: 2197392,
  2021: 5045249,
  2022: 9935034,
  2023: 15974367,
  2024: 24769574,
  2025: 33349445,
};
const baseUnaffiliatedRevenue = {
  2017: 4678.32,
  2018: 42180.31,
  2019: 78895.59,
  2020: 140791.35,
  2021: 295783.31,
  2022: 420943.96,
  2023: 536146.23,
  2024: 656841.44,
  2025: 878901.02,
};
const baseRevUnitRpu = {
  2017: { 2017: 0.5 },
  2018: { 2017: 2.57, 2018: 0.67 },
  2019: { 2017: 2.58, 2018: 1.4, 2019: 0.7 },
  2020: { 2017: 3.4, 2018: 1.94, 2019: 1.99, 2020: 1.01 },
  2021: { 2017: 5.73, 2018: 3.09, 2019: 2.96, 2020: 2.42, 2021: 1.19 },
  2022: { 2017: 10.45, 2018: 4.65, 2019: 4.15, 2020: 3.3, 2021: 2.9, 2022: 1.32 },
  2023: { 2017: 12.89, 2018: 5.85, 2019: 5.24, 2020: 3.91, 2021: 3.8, 2022: 3.09, 2023: 1.57 },
  2024: { 2017: 14.16, 2018: 7.51, 2019: 6.63, 2020: 4.89, 2021: 5.02, 2022: 4.45, 2023: 3.68, 2024: 1.68 },
  2025: { 2017: 14.4, 2018: 8.18, 2019: 7.23, 2020: 5.45, 2021: 5.66, 2022: 5.24, 2023: 4.67, 2024: 3.89, 2025: 1.95 },
};

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
  newCustomers: { label: "New Customers", format: "integer", chartId: "chart-new-customers", yMax: 2000, scale: 1000, suffix: "000s", color: "#6fa76b" },
  profilesReturning: { label: "Profiles / Returning Cust", format: "decimal", chartId: "chart-profiles-returning", yMax: 2, color: "#7b6fb8" },
  profilesNew: { label: "Profiles / New Cust", format: "decimal", chartId: "chart-profiles-new", yMax: 2, color: "#a96c50" },
  revReturningProfile: { label: "Rev / Returning Cust Profile", format: "currency2", chartId: "chart-rev-returning", yMax: 40, color: "#4f7fb8" },
  revNewProfile: { label: "Rev / New Cust Profile", format: "currency2", chartId: "chart-rev-new", yMax: 50, color: "#4b9b8e" },
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
  currentView: "chart",
  syncingCharts: false,
  syncingCohortCharts: false,
  undoStack: [],
  redoStack: [],
  revUnitCharts: {},
  syncingRevUnitCharts: false,
  pendingDialogAction: null,
  cohortEditorMode: "perCohort",
  cohortEditorCohortYear: 2025,
  yearEditorYear: 2026,
  newCustomerAllCohortsChartType: "line",
  newCustomerAllCohortsHighlights: [],
  newCustomerAllCohortsHighlightKey: "",
  newCustomerUnitDrilldownOpen: false,
  newCustomerNewUnitsDrilldownOpen: false,
  revUnitEditorMode: "cohort",
  revUnitEditorCohortYear: 2025,
  revUnitReferenceCohortYear: 2025,
  revUnitReferenceShifted: false,
  canvasFocusNode: "revenue",
  canvasZoom: 1,
  anonymizedView: false,
};

const originalTextByNode = new WeakMap();

function isAnonymizedView() {
  return state.anonymizedView;
}

function preserveCase(replacement, source) {
  if (source === source.toUpperCase()) return replacement.toUpperCase();
  if (source[0] === source[0].toUpperCase()) {
    return replacement.replace(/\b[a-z]/g, letter => letter.toUpperCase());
  }
  return replacement;
}

function replaceUnitTerm(text) {
  return text.replace(/\b(units|unit)\b/gi, (match, _term, offset, source) => {
    const prefix = source.slice(Math.max(0, offset - 8), offset).toLowerCase();
    if (prefix.endsWith("partner ")) return match;
    return preserveCase(match.toLowerCase() === "units" ? "partner units" : "partner unit", match);
  });
}

function anonymizeText(text) {
  if (!text) return text;
  return replaceUnitTerm(text)
    .replace(/PetScreening \$100M Path/g, "Partner Growth Path")
    .replace(/\bprofiles\b/g, "transactions")
    .replace(/\bProfiles\b/g, "Transactions")
    .replace(/\bprofile\b/g, "transaction")
    .replace(/\bProfile\b/g, "Transaction")
    .replace(/\bproperties\b/g, "partners")
    .replace(/\bProperties\b/g, "Partners")
    .replace(/\bproperty\b/g, "partner")
    .replace(/\bProperty\b/g, "Partner")
    .replace(/\bHHP\b/g, "Core");
}

function displayLabel(text) {
  return isAnonymizedView() ? anonymizeText(String(text)) : String(text);
}

function anonymizedMetricValue() {
  return "Y";
}

function displayScenarioName(scenarioOrName) {
  if (!isAnonymizedView()) return typeof scenarioOrName === "string" ? scenarioOrName : scenarioOrName?.name || "";
  const scenarioId = typeof scenarioOrName === "string" ? null : scenarioOrName?.id;
  const scenarios = Object.values(state.scenarios);
  const index = scenarioId ? scenarios.findIndex(scenario => scenario.id === scenarioId) : -1;
  const letterIndex = index >= 0 ? index : 0;
  return `Scenario ${String.fromCharCode(65 + (letterIndex % 26))}`;
}

const canvasNodeTree = {
  revenue: { parent: null, children: ["payingCustomers", "revenuePerCustomer"] },
  payingCustomers: { parent: "revenue", children: ["newPayingCustomers", "returningPayingCustomers"] },
  revenuePerCustomer: { parent: "revenue", children: ["profilesPerCustomer", "revenuePerProfile"] },
  newPayingCustomers: { parent: "payingCustomers", children: ["existingPropertyNewPayingCustomers", "newPropertyNewPayingCustomers"] },
  returningPayingCustomers: { parent: "payingCustomers", children: ["priorPayingCustomers", "retentionRate"] },
  profilesPerCustomer: { parent: "revenuePerCustomer", children: ["profilesReturningCustomer", "profilesNewCustomer"] },
  revenuePerProfile: { parent: "revenuePerCustomer", children: ["revenueReturningProfile", "revenueNewProfile"] },
  profilesReturningCustomer: { parent: "profilesPerCustomer", children: [] },
  profilesNewCustomer: { parent: "profilesPerCustomer", children: [] },
  revenueReturningProfile: { parent: "revenuePerProfile", children: [] },
  revenueNewProfile: { parent: "revenuePerProfile", children: [] },
  existingPropertyNewPayingCustomers: { parent: "newPayingCustomers", children: [] },
  newPropertyNewPayingCustomers: { parent: "newPayingCustomers", children: ["newUnits", "newCustomersPerUnit"] },
  newUnits: { parent: "newPropertyNewPayingCustomers", children: ["newProperties", "newUnitsPerProperty"] },
  newCustomersPerUnit: { parent: "newPropertyNewPayingCustomers", children: [] },
  newProperties: { parent: "newUnits", children: [] },
  newUnitsPerProperty: { parent: "newUnits", children: [] },
  priorPayingCustomers: { parent: "returningPayingCustomers", children: [] },
  retentionRate: { parent: "returningPayingCustomers", children: [] },
};

const canvasLayoutRows = [
  ["revenue"],
  ["payingCustomers", "revenuePerCustomer"],
  ["newPayingCustomers", "returningPayingCustomers", "profilesPerCustomer", "revenuePerProfile"],
  [
    "existingPropertyNewPayingCustomers",
    "newPropertyNewPayingCustomers",
    "priorPayingCustomers",
    "retentionRate",
    "profilesReturningCustomer",
    "profilesNewCustomer",
    "revenueReturningProfile",
    "revenueNewProfile",
  ],
  ["newUnits", "newCustomersPerUnit"],
  ["newProperties", "newUnitsPerProperty"],
];

const canvasFocusDimensions = {
  main: { width: 660, height: 360 },
  secondary: { width: 430, height: 285 },
  tertiary: { width: 430, height: 285 },
};

const CANVAS_ZOOM_MIN = 0.55;
const CANVAS_ZOOM_MAX = 1;
const CANVAS_ZOOM_STEP = 0.1;

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
    revUnitPlan: createDefaultRevUnitPlan(),
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
  Object.keys(scenario.newCustomerDrilldown.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.newCustomerDrilldown.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.newCustomerDrilldown.controlPoints[key];
    } else {
      scenario.newCustomerDrilldown.controlPoints[key] = controlPoints;
    }
  });
  if (!Array.isArray(scenario.newCustomerDrilldown.lockedCohorts)) {
    scenario.newCustomerDrilldown.lockedCohorts = [];
  }
  if (!scenario.newCustomerDrilldown.unitPlan) {
    scenario.newCustomerDrilldown.unitPlan = createDefaultNewCustomerUnitPlan(scenario.newCustomerDrilldown.counts);
  }
  normalizeNewCustomerUnitPlan(scenario);
  if (!scenario.revUnitPlan || !scenario.revUnitPlan.newUnits) {
    scenario.revUnitPlan = createDefaultRevUnitPlan();
  }
  if (!scenario.revUnitPlan.unaffiliatedRevenue) {
    scenario.revUnitPlan.unaffiliatedRevenue = clone(baseUnaffiliatedRevenue);
  }
  if (!scenario.revUnitPlan.controlPoints) {
    scenario.revUnitPlan.controlPoints = { newUnits: clone(REV_UNIT_YEARS) };
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
    revUnitPlan: clone(scenario.revUnitPlan),
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
  const counts = clone(baseNewCustomerCohortCounts);
  return {
    counts,
    controlPoints: {},
    lockedCohorts: [],
    unitPlan: createDefaultNewCustomerUnitPlan(counts),
  };
}

function createDefaultNewCustomerUnitPlan(counts = baseNewCustomerCohortCounts) {
  const newUnits = clone(baseRevUnitNewUnits);
  const customersPerUnit = {};
  YEARS.forEach(year => {
    const units = Number(newUnits[String(year)] ?? newUnits[year] ?? 0);
    const newCustomers = newCustomerCohortValue(counts, year, year);
    customersPerUnit[String(year)] = units > 0 ? newCustomers / units : 0;
  });
  return {
    newUnits,
    customersPerUnit,
    propertyPlan: createDefaultNewCustomerNewUnitsPropertyPlan(newUnits),
    controlPoints: {
      newUnits: clone(YEARS),
      customersPerUnit: clone(YEARS),
    },
  };
}

function baselineNewPropertiesForUnits(year, units) {
  const historicalProperties = Number(baseRevUnitNewProperties[String(year)] ?? baseRevUnitNewProperties[year]);
  if (Number.isFinite(historicalProperties) && historicalProperties > 0) return historicalProperties;
  const latestHistoricalUnits = Number(baseRevUnitNewUnits[REV_UNIT_LAST_HISTORICAL_YEAR] || 0);
  const latestHistoricalProperties = Number(baseRevUnitNewProperties[REV_UNIT_LAST_HISTORICAL_YEAR] || 0);
  const latestUnitsPerProperty = latestHistoricalProperties > 0 ? latestHistoricalUnits / latestHistoricalProperties : 0;
  return latestUnitsPerProperty > 0 && units > 0 ? Math.max(1, Math.round(units / latestUnitsPerProperty)) : 0;
}

function generatedPropertiesAt250Units(units) {
  return units > 0 ? Math.max(1, Math.round(units / 250)) : 0;
}

function isGeneratedNewUnitsPropertyPlan(propertyPlan, newUnits) {
  if (!propertyPlan?.newProperties || !propertyPlan?.unitsPerNewProperty) return false;
  return YEARS.every(year => {
    const units = Number(newUnits[String(year)] ?? newUnits[year] ?? 0);
    const expectedProperties = generatedPropertiesAt250Units(units);
    const expectedUnitsPerProperty = expectedProperties > 0 ? units / expectedProperties : 0;
    const properties = Number(propertyPlan.newProperties[String(year)] ?? propertyPlan.newProperties[year]);
    const unitsPerProperty = Number(propertyPlan.unitsPerNewProperty[String(year)] ?? propertyPlan.unitsPerNewProperty[year]);
    return Math.abs(properties - expectedProperties) < 0.5
      && Math.abs(unitsPerProperty - expectedUnitsPerProperty) < 0.01;
  });
}

function createDefaultNewCustomerNewUnitsPropertyPlan(newUnits = baseRevUnitNewUnits) {
  const newProperties = {};
  const unitsPerNewProperty = {};
  YEARS.forEach(year => {
    const units = Number(newUnits[String(year)] ?? newUnits[year] ?? 0);
    const properties = baselineNewPropertiesForUnits(year, units);
    newProperties[String(year)] = properties;
    unitsPerNewProperty[String(year)] = properties > 0 ? units / properties : 0;
  });
  return {
    newProperties,
    unitsPerNewProperty,
    controlPoints: {
      newProperties: clone(YEARS),
      unitsPerNewProperty: clone(YEARS),
    },
  };
}

function normalizeNewCustomerUnitPlan(scenario) {
  const plan = scenario.newCustomerDrilldown.unitPlan;
  if (!plan.newUnits) plan.newUnits = clone(baseRevUnitNewUnits);
  if (!plan.customersPerUnit) plan.customersPerUnit = {};
  if (!plan.propertyPlan) plan.propertyPlan = createDefaultNewCustomerNewUnitsPropertyPlan(plan.newUnits);
  if (isGeneratedNewUnitsPropertyPlan(plan.propertyPlan, plan.newUnits)) {
    plan.propertyPlan = createDefaultNewCustomerNewUnitsPropertyPlan(plan.newUnits);
  }
  YEARS.forEach(year => {
    if (!Number.isFinite(Number(plan.newUnits[String(year)] ?? plan.newUnits[year]))) {
      plan.newUnits[String(year)] = baseRevUnitNewUnits[String(year)] ?? baseRevUnitNewUnits[year] ?? 0;
    }
    if (!Number.isFinite(Number(plan.customersPerUnit[String(year)] ?? plan.customersPerUnit[year]))) {
      const units = Number(plan.newUnits[String(year)] ?? plan.newUnits[year] ?? 0);
      const newCustomers = newCustomerCohortValue(scenario.newCustomerDrilldown.counts, year, year);
      plan.customersPerUnit[String(year)] = units > 0 ? newCustomers / units : 0;
    }
  });
  if (!plan.controlPoints) {
    plan.controlPoints = {
      newUnits: clone(YEARS),
      customersPerUnit: clone(YEARS),
    };
  }
  if (!plan.propertyPlan.controlPoints) {
    plan.propertyPlan.controlPoints = {
      newProperties: clone(YEARS),
      unitsPerNewProperty: clone(YEARS),
    };
  }
  if (!plan.propertyPlan.newProperties) plan.propertyPlan.newProperties = {};
  if (!plan.propertyPlan.unitsPerNewProperty) plan.propertyPlan.unitsPerNewProperty = {};
  YEARS.forEach(year => {
    if (!Number.isFinite(Number(plan.propertyPlan.newProperties[String(year)] ?? plan.propertyPlan.newProperties[year]))) {
      const units = Number(plan.newUnits[String(year)] ?? plan.newUnits[year] ?? 0);
      plan.propertyPlan.newProperties[String(year)] = baselineNewPropertiesForUnits(year, units);
    }
    if (!Number.isFinite(Number(plan.propertyPlan.unitsPerNewProperty[String(year)] ?? plan.propertyPlan.unitsPerNewProperty[year]))) {
      const units = Number(plan.newUnits[String(year)] ?? plan.newUnits[year] ?? 0);
      const properties = Number(plan.propertyPlan.newProperties[String(year)] ?? plan.propertyPlan.newProperties[year] ?? 0);
      plan.propertyPlan.unitsPerNewProperty[String(year)] = properties > 0 ? units / properties : 0;
    }
  });
}

function createDefaultRevUnitPlan() {
  return {
    newUnits: clone(baseRevUnitNewUnits),
    revenue: clone(baseRevUnitRevenue),
    revPerUnit: clone(baseRevUnitRpu),
    unaffiliatedRevenue: clone(baseUnaffiliatedRevenue),
    controlPoints: { newUnits: clone(REV_UNIT_YEARS) },
  };
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

function newCustomerCohortStackValue(counts, year, cohortYear) {
  const targetIndex = NEW_CUSTOMER_COHORT_YEARS.indexOf(Number(cohortYear));
  if (targetIndex < 0) return 0;
  return NEW_CUSTOMER_COHORT_YEARS.slice(0, targetIndex + 1).reduce((sum, stackedCohortYear) => {
    if (!newCustomerCohortApplies(stackedCohortYear, year)) return sum;
    return sum + newCustomerCohortValue(counts, year, stackedCohortYear);
  }, 0);
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

function newCustomerCohortColor(cohortYear) {
  const colors = ["#647184", "#4f7fb8", "#6fa76b", "#d39b2a", "#7b6fb8", "#a96c50", "#4b9b8e", "#8e6bb0", "#c47a5a", "#5f9ea0", "#9a8c55", "#5875a4", "#76a46b", "#b66f8f"];
  const index = Math.max(0, NEW_CUSTOMER_COHORT_YEARS.indexOf(cohortYear));
  return colors[index % colors.length];
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

function existingPropertyNewCustomersTotal(counts, year) {
  return existingPropertyCohortsForYear(year).reduce((sum, cohortYear) => {
    return sum + newCustomerCohortValue(counts, year, cohortYear);
  }, 0);
}

function lockedNewCustomerCohorts(scenario = activeScenario()) {
  return new Set((scenario.newCustomerDrilldown.lockedCohorts || []).map(Number));
}

function sameNumberList(left, right) {
  if (left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}

function isNewCustomerCohortLocked(cohortYear, scenario = activeScenario()) {
  return lockedNewCustomerCohorts(scenario).has(Number(cohortYear));
}

function setExistingPropertyNewCustomersTotal(counts, year, targetTotal) {
  const cohorts = existingPropertyCohortsForYear(year);
  if (!cohorts.length) return;
  const unlockedCohorts = cohorts.filter(cohortYear => !isNewCustomerCohortLocked(cohortYear));
  if (!unlockedCohorts.length) return;
  const lockedTotal = cohorts
    .filter(cohortYear => isNewCustomerCohortLocked(cohortYear))
    .reduce((sum, cohortYear) => sum + newCustomerCohortValue(counts, year, cohortYear), 0);
  const unlockedTargetTotal = Math.max(0, targetTotal - lockedTotal);
  const currentUnlockedTotal = unlockedCohorts.reduce((sum, cohortYear) => {
    return sum + newCustomerCohortValue(counts, year, cohortYear);
  }, 0);
  if (currentUnlockedTotal > 0) {
    unlockedCohorts.forEach(cohortYear => {
      const current = newCustomerCohortValue(counts, year, cohortYear);
      setNewCustomerCohortValuePreservingFutureYoy(counts, year, cohortYear, (current / currentUnlockedTotal) * unlockedTargetTotal);
    });
    return;
  }
  const equalValue = unlockedTargetTotal / unlockedCohorts.length;
  unlockedCohorts.forEach(cohortYear => setNewCustomerCohortValuePreservingFutureYoy(counts, year, cohortYear, equalValue));
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
  const visiblePairs = pairs.filter(([x]) => visibleX.has(Number(x)));
  if (visiblePairs.length >= 2) return visiblePairs;
  if (pairs.length <= 2) return pairs;
  return [pairs[0], pairs[pairs.length - 1]];
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

function hasNapkinControlPointSet(key) {
  return Array.isArray(activeScenario().newCustomerDrilldown.controlPoints?.[key]);
}

function addNapkinControlPointIfControlled(key, x) {
  if (hasNapkinControlPointSet(key)) addNapkinControlPoint(key, x);
}

function ensureBridgeExistingPointsForActuals() {
  const scenario = activeScenario();
  const key = napkinControlKey("bridgeExistingProperties");
  const stored = scenario.newCustomerDrilldown.controlPoints?.[key];
  if (!Array.isArray(stored)) return;
  const counts = scenario.newCustomerDrilldown.counts;
  const visibleYears = new Set(stored.map(Number));
  const actualPairs = YEARS.map(year => [year, existingPropertyNewCustomersTotal(counts, year)]);
  const visiblePairs = actualPairs.filter(([year]) => visibleYears.has(Number(year)));
  YEARS.filter(editableYear).forEach(year => {
    if (visibleYears.has(year)) return;
    const actual = existingPropertyNewCustomersTotal(counts, year);
    const interpolated = interpolateNapkinLineValue(visiblePairs, year);
    if (interpolated === null || Math.abs(actual - interpolated) > 0.5) {
      addNapkinControlPoint(key, year);
    }
  });
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

function removedNapkinXs(beforeData, afterData) {
  const after = new Set((afterData || []).map(([x]) => Number(x)));
  return (beforeData || [])
    .map(([x]) => Number(x))
    .filter(x => Number.isFinite(x) && !after.has(x));
}

function affectedNapkinXs(beforeData, afterData) {
  return Array.from(new Set([
    ...changedNapkinXs(beforeData, afterData),
    ...removedNapkinXs(beforeData, afterData),
  ])).sort((left, right) => left - right);
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

function setNewCustomerCohortValuePreservingFutureYoy(counts, year, cohortYear, value) {
  const preservedFutureYoy = new Map();
  YEARS.forEach(candidateYear => {
    if (candidateYear > year && newCustomerCohortApplies(cohortYear, candidateYear)) {
      preservedFutureYoy.set(candidateYear, newCustomerCohortYoy(counts, candidateYear, cohortYear));
    }
  });

  setNewCustomerCohortValue(counts, year, cohortYear, value);

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

function visibleNewCustomerBridgeValue(key, pairs, year, fallback) {
  const value = interpolateNapkinLineValue(controlledNapkinPairs(key, pairs), year);
  return value === null ? fallback : value;
}

function visibleExistingPropertyNewCustomersValue(year) {
  const counts = activeScenario().newCustomerDrilldown.counts;
  const key = napkinControlKey("bridgeExistingProperties");
  const fallback = existingPropertyNewCustomersTotal(counts, year);
  const pairs = YEARS.map(candidateYear => [
    candidateYear,
    existingPropertyNewCustomersTotal(counts, candidateYear),
  ]);
  return visibleNewCustomerBridgeValue(key, pairs, year, fallback);
}

function visibleNewPropertyNewCustomersValue(year) {
  const counts = activeScenario().newCustomerDrilldown.counts;
  const key = napkinControlKey("bridgeNewProperties");
  const fallback = newPropertyCohortValue(counts, year);
  const pairs = YEARS.map(candidateYear => [
    candidateYear,
    newPropertyCohortValue(counts, candidateYear),
  ]);
  return visibleNewCustomerBridgeValue(key, pairs, year, fallback);
}

function visibleBottomUpNewCustomersValue(year) {
  return visibleExistingPropertyNewCustomersValue(year) + visibleNewPropertyNewCustomersValue(year);
}

function visibleBottomUpNewCustomers() {
  return YEARS.map(visibleBottomUpNewCustomersValue);
}

function effectiveNewCustomers(scenario) {
  return scenario.drivers.newCustomers;
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
  syncRevUnitCharts();
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
  if (isAnonymizedView()) return anonymizedMetricValue();
  const sign = value < 0 ? "-" : "";
  return `${sign}$${Math.abs(value).toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  })}`;
}

function formatOptionalCurrency(value, decimals = 0) {
  return Number.isFinite(value) ? formatCurrency(value, decimals) : "-";
}

function formatCompactCurrency(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  const abs = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (abs >= 1000000000) return `${sign}$${trimNumber(abs / 1000000000, 1)}B`;
  if (abs >= 1000000) return `${sign}$${trimNumber(abs / 1000000, 1)}M`;
  if (abs >= 1000) return `${sign}$${trimNumber(abs / 1000, 0)}k`;
  return formatCurrency(value, 0);
}

function formatCompactNumber(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
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
  if (isAnonymizedView()) return anonymizedMetricValue();
  const actualValue = value * (meta.scale || 1);
  if (meta.format === "percent") return `${trimNumber(actualValue * 100, 0)}%`;
  if (meta.format === "currency2") return formatCompactCurrency(actualValue);
  if (meta.format === "integer") return formatCompactNumber(actualValue);
  if (meta.format === "decimal") return trimNumber(actualValue, 2);
  return formatCompactNumber(actualValue);
}

function formatDriverTooltipValue(value, meta) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  const actualValue = value * (meta.scale || 1);
  if (meta.format === "percent") return `${trimNumber(actualValue * 100, 1)}%`;
  if (meta.format === "currency2") return formatCurrency(actualValue, 2);
  if (meta.format === "integer") return Math.round(actualValue).toLocaleString("en-US");
  if (meta.format === "decimal") return actualValue.toFixed(2);
  return String(actualValue);
}

function formatIntegerMetric(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return Math.round(value).toLocaleString("en-US");
}

function formatPercentMetric(value, decimals = 1) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return `${trimNumber(value, decimals)}%`;
}

function formatDecimalMetric(value, decimals = 2) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return trimNumber(value, decimals);
}

function formatAxisYear(value) {
  const year = Number(value);
  return Number.isInteger(year) ? String(year) : "";
}

function tooltipHeader(axisValue) {
  const year = Array.isArray(axisValue) ? axisValue[0] : axisValue;
  return `<strong>${year}</strong>`;
}

function napkinChartByKey(key) {
  return state.cohortCharts[key] || state.revUnitCharts[key] || state.charts[key] || null;
}

function stepYAxisLeadingDigit(value, direction) {
  if (!Number.isFinite(value) || value <= 0) return null;
  const rounded = Math.max(1, Math.round(value));
  const place = Math.pow(10, Math.floor(Math.log10(rounded)));
  const lead = Math.floor(rounded / place);
  if (direction === "up") {
    return lead >= 9 ? 10 * place : (lead + 1) * place;
  }
  if (direction === "down") {
    if (lead > 1) return (lead - 1) * place;
    if (place === 1) return Math.max(1, rounded - 1);
    return 9 * (place / 10);
  }
  return null;
}

function applyNapkinYAxisMax(chart, nextYMax) {
  if (!chart?.baseOption?.yAxis || !Number.isFinite(nextYMax)) return;
  const yMin = Number(chart.baseOption.yAxis.min);
  if (Number.isFinite(yMin) && nextYMax <= yMin) return;
  chart.baseOption.yAxis.max = nextYMax;
  chart.chart.setOption({ yAxis: { max: nextYMax } }, false);
  chart._refreshChart();
}

function handleNapkinYAxisAction(chart, action) {
  const yMax = Number(chart?.baseOption?.yAxis?.max);
  if (!Number.isFinite(yMax)) return;
  if (action === "large-up") {
    applyNapkinYAxisMax(chart, yMax * 10);
  } else if (action === "large-down") {
    applyNapkinYAxisMax(chart, yMax / 10);
  } else if (action === "small-up") {
    applyNapkinYAxisMax(chart, stepYAxisLeadingDigit(yMax, "up"));
  } else if (action === "small-down") {
    applyNapkinYAxisMax(chart, stepYAxisLeadingDigit(yMax, "down"));
  }
}

function bindNapkinYAxisControls() {
  document.querySelectorAll("[data-napkin-y-axis-controls]").forEach(row => {
    row.addEventListener("click", event => {
      const button = event.target.closest("[data-y-axis-action]");
      if (!button || !row.contains(button)) return;
      handleNapkinYAxisAction(napkinChartByKey(row.dataset.napkinYAxisControls), button.dataset.yAxisAction);
    });
  });
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
  const lineIsEditable = true;
  return {
    name: displayScenarioName(scenario),
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
    name: `Comparison: ${displayScenarioName(scenario)}`,
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
            rows.push(`${item.marker || ""} ${item.seriesName || displayLabel(meta.label)}: ${formatDriverTooltipValue(Number(rawValue), meta)}`);
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
    YEARS.forEach((year, index) => {
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      if (value !== null && editableYear(year)) activeScenario().drivers[key][index] = value * (meta.scale || 1);
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

function initEchartById(id) {
  const element = document.getElementById(id);
  return element ? echarts.init(element) : null;
}

function initOutputCharts() {
  state.outputCharts.revenue = initEchartById("revenue-chart");
  state.outputCharts.gap = initEchartById("gap-chart");
  state.outputCharts.newCustomerDrilldown = initEchartById("new-customers-drilldown-chart");
  state.outputCharts.newCustomerTotal = initEchartById("new-customers-total-chart");
  state.outputCharts.newCustomerUnitBridge = initEchartById("new-customers-unit-bridge-chart");
  state.outputCharts.newCustomerUnitDelta = initEchartById("new-customers-unit-delta-chart");
  state.outputCharts.newCustomerNewUnitsBridge = initEchartById("new-customers-new-units-bridge-chart");
  state.outputCharts.newCustomerNewUnitsDelta = initEchartById("new-customers-new-units-delta-chart");
  state.outputCharts.newCustomerAllCohorts = initEchartById("new-customers-all-cohorts-chart");
  state.outputCharts.revUnitRevenueComparison = initEchartById("rev-unit-revenue-comparison-chart");
  state.outputCharts.revUnitAggregateRpu = initEchartById("rev-unit-aggregate-rpu-chart");
  state.outputCharts.treeRevenue = initEchartById("tree-revenue-chart");
  state.outputCharts.treePayingCustomers = initEchartById("tree-paying-customers-chart");
  state.outputCharts.treeNewPayingCustomers = initEchartById("tree-new-paying-customers-chart");
  state.outputCharts.treeExistingPropertyNewPayingCustomers = initEchartById("tree-existing-property-new-paying-customers-chart");
  state.outputCharts.treeNewPropertyNewPayingCustomers = initEchartById("tree-new-property-new-paying-customers-chart");
  state.outputCharts.treeNewUnits = initEchartById("tree-new-units-chart");
  state.outputCharts.treeNewProperties = initEchartById("tree-new-properties-chart");
  state.outputCharts.treeNewUnitsPerProperty = initEchartById("tree-new-units-per-property-chart");
  state.outputCharts.treeNewCustomersPerUnit = initEchartById("tree-new-customers-per-unit-chart");
  state.outputCharts.treeReturningPayingCustomers = initEchartById("tree-returning-paying-customers-chart");
  state.outputCharts.treePriorPayingCustomers = initEchartById("tree-prior-paying-customers-chart");
  state.outputCharts.treeRetentionRate = initEchartById("tree-retention-rate-chart");
  state.outputCharts.treeRevenuePerCustomer = initEchartById("tree-revenue-per-customer-chart");
  state.outputCharts.treeProfilesPerCustomer = initEchartById("tree-profiles-per-customer-chart");
  state.outputCharts.treeRevenuePerProfile = initEchartById("tree-revenue-per-profile-chart");
  state.outputCharts.treeProfilesReturningCustomer = initEchartById("tree-profiles-returning-customer-chart");
  state.outputCharts.treeProfilesNewCustomer = initEchartById("tree-profiles-new-customer-chart");
  state.outputCharts.treeRevenueReturningProfile = initEchartById("tree-revenue-returning-profile-chart");
  state.outputCharts.treeRevenueNewProfile = initEchartById("tree-revenue-new-profile-chart");
  window.addEventListener("resize", () => {
    Object.values(state.outputCharts).forEach(chart => chart?.resize());
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
        formatter: params => formatCohortNapkinTooltip(params, formatIntegerMetric),
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
      yAxis: { type: "value", min: -50, max: 100, axisLabel: { formatter: value => formatPercentMetric(value, 0) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => formatPercentMetric(value, 1)),
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
      yAxis: { type: "value", min: -50, max: 100, axisLabel: { formatter: value => formatPercentMetric(value, 0) } },
      grid: { left: 12, right: 18, top: 14, bottom: 48, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatYearEditorTooltip(params, value => formatPercentMetric(value, 1)),
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
        formatter: params => formatCohortNapkinTooltip(params, formatIntegerMetric),
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
      yAxis: { type: "value", min: -50, max: 100, axisLabel: { formatter: value => formatPercentMetric(value, 0) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => formatPercentMetric(value, 1)),
      },
    },
    "none",
    false
  );
  const bridgeExistingChart = new NapkinChart(
    "new-customers-bridge-existing-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 1000000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, formatIntegerMetric),
      },
    },
    "none",
    false
  );
  const bridgeNewChart = new NapkinChart(
    "new-customers-bridge-new-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 150000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, formatIntegerMetric),
      },
    },
    "none",
    false
  );
  const unitNewUnitsChart = new NapkinChart(
    "new-customers-unit-new-units-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 2000000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, formatIntegerMetric),
      },
    },
    "none",
    false
  );
  const unitCustomersPerUnitChart = new NapkinChart(
    "new-customers-unit-customers-per-unit-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 0.12, axisLabel: { formatter: value => formatDecimalMetric(value, 2) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => formatDecimalMetric(value, 4)),
      },
    },
    "none",
    false
  );
  const newUnitsPropertiesChart = new NapkinChart(
    "new-customers-new-units-properties-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 8000, axisLabel: { formatter: formatCompactNumber } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, formatIntegerMetric),
      },
    },
    "none",
    false
  );
  const newUnitsUnitsPerPropertyChart = new NapkinChart(
    "new-customers-new-units-units-per-property-chart",
    [],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 500, axisLabel: { formatter: value => formatDecimalMetric(value, 0) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => formatDecimalMetric(value, 1)),
      },
    },
    "none",
    false
  );

  [countChart, yoyChart, yearChart, yearYoyChart, newPropertiesCountChart, newPropertiesYoyChart, bridgeExistingChart, bridgeNewChart, unitNewUnitsChart, unitCustomersPerUnitChart, newUnitsPropertiesChart, newUnitsUnitsPerPropertyChart].forEach(chart => {
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
    chart.chart.getZr().on("mouseup", () => {
      if (!chart._appEditCommitted) return;
      setTimeout(() => {
        syncNewCustomerCohortEditors();
        renderNewCustomerDrilldownTable();
        renderNewCustomerYoyTable();
        chart._appEditSnapshot = null;
        chart._appEditLineSnapshot = null;
        chart._appEditCommitted = false;
      }, 0);
    });
    chart.chart.on("mouseover", params => highlightCellsForNapkinPoint(chart, params));
    chart.chart.on("click", params => highlightCellsForNapkinPoint(chart, params));
    chart.chart.on("mouseout", clearNewCustomerCellHighlights);
    chart.chart.getZr().on("globalout", clearNewCustomerCellHighlights);
  });

  countChart.onDataChanged = () => handleCohortCountChartChange(countChart);
  yoyChart.onDataChanged = () => handleCohortYoyChartChange(yoyChart);
  yearChart.onDataChanged = () => handleYearExistingChartChange(yearChart);
  yearYoyChart.onDataChanged = () => handleYearExistingYoyChartChange(yearYoyChart);
  newPropertiesCountChart.onDataChanged = () => handleNewPropertiesCountChartChange(newPropertiesCountChart);
  newPropertiesYoyChart.onDataChanged = () => handleNewPropertiesYoyChartChange(newPropertiesYoyChart);
  bridgeExistingChart.onDataChanged = () => handleBridgeExistingChartChange(bridgeExistingChart);
  bridgeNewChart.onDataChanged = () => handleBridgeNewChartChange(bridgeNewChart);
  unitNewUnitsChart.onDataChanged = () => handleNewCustomerUnitNewUnitsChartChange(unitNewUnitsChart);
  unitCustomersPerUnitChart.onDataChanged = () => handleNewCustomerUnitCustomersPerUnitChartChange(unitCustomersPerUnitChart);
  newUnitsPropertiesChart.onDataChanged = () => handleNewCustomerNewUnitsPropertiesChartChange(newUnitsPropertiesChart);
  newUnitsUnitsPerPropertyChart.onDataChanged = () => handleNewCustomerNewUnitsUnitsPerPropertyChartChange(newUnitsUnitsPerPropertyChart);
  state.cohortCharts.count = countChart;
  state.cohortCharts.yoy = yoyChart;
  state.cohortCharts.yearExisting = yearChart;
  state.cohortCharts.yearExistingYoy = yearYoyChart;
  state.cohortCharts.newPropertiesCount = newPropertiesCountChart;
  state.cohortCharts.newPropertiesYoy = newPropertiesYoyChart;
  state.cohortCharts.bridgeExisting = bridgeExistingChart;
  state.cohortCharts.bridgeNew = bridgeNewChart;
  state.cohortCharts.unitNewUnits = unitNewUnitsChart;
  state.cohortCharts.unitCustomersPerUnit = unitCustomersPerUnitChart;
  state.cohortCharts.newUnitsProperties = newUnitsPropertiesChart;
  state.cohortCharts.newUnitsUnitsPerProperty = newUnitsUnitsPerPropertyChart;
  countChart._appHighlightMapper = (point, index, chart) => ({
    countYears: napkinInfluenceYears(chart, index),
    countCohortYear: selectedNewCustomerCohortYear(),
  });
  yoyChart._appHighlightMapper = (point, index, chart) => ({
    countYears: napkinInfluenceYears(chart, index),
    countCohortYear: selectedNewCustomerCohortYear(),
    yoyYears: napkinInfluenceYears(chart, index),
    yoyCohortYear: selectedNewCustomerCohortYear(),
  });
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
  newPropertiesCountChart._appHighlightMapper = point => {
    const year = Number(point[0]);
    return { countYears: YEARS.filter(item => item >= year), countCohortYear: year };
  };
  newPropertiesYoyChart._appHighlightMapper = point => {
    const year = Number(point[0]);
    return { countYears: YEARS.filter(item => item >= year), countCohortYear: year };
  };
  bridgeExistingChart._appHighlightMapper = point => {
    const year = Number(point[0]);
    return {
      countYear: year,
      countCohortYears: existingPropertyCohortsForYear(year).filter(cohortYear => !isNewCustomerCohortLocked(cohortYear)),
    };
  };
  bridgeNewChart._appHighlightMapper = point => {
    const year = Number(point[0]);
    return { countYears: YEARS.filter(item => item >= year), countCohortYear: year };
  };
}

function highlightCellsForNapkinPoint(chart, params) {
  if (!chart?._appHighlightMapper || !params || params.seriesType !== "line") return;
  const lineIndex = Math.floor(Number(params.seriesIndex) / 2);
  if (lineIndex !== 0) return;
  const pointIndex = Number(params.dataIndex);
  const point = chart.lines?.[0]?.data?.[pointIndex];
  if (!point) return;
  highlightNewCustomerCells(chart._appHighlightMapper(point, pointIndex, chart));
}

function highlightCellsForNapkinDrag(chart) {
  if (!chart?._appHighlightMapper || chart._draggingLineIndex !== 0 || chart._draggingPointIndex === null) return;
  const pointIndex = Number(chart._draggingPointIndex);
  const point = chart.lines?.[0]?.data?.[pointIndex];
  if (!point) return;
  highlightNewCustomerCells(chart._appHighlightMapper(point, pointIndex, chart));
}

function napkinInfluenceYears(chart, pointIndex) {
  const data = chart?.lines?.[0]?.data || [];
  const point = data[pointIndex];
  if (!point) return [];
  const pointYear = Number(point[0]);
  return YEARS.filter(year => editableYear(year) && year >= pointYear);
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
      rows.push(`${item.marker || ""} ${displayLabel(item.seriesName)}: ${isAnonymizedView() ? anonymizedMetricValue() : formatter(Number(rawValue))}`);
    });
  return [tooltipHeader(year), ...rows].join("<br/>");
}

function formatYearEditorAxisLabel(value) {
  const year = selectedNewCustomerYearEditorYear();
  const cohorts = existingPropertyCohortsForYear(year);
  const cohortYear = yearEditorCohortForX(cohorts, value);
  return cohortYear === null ? "" : cohortLabel(cohortYear);
}

function formatYearEditorTooltip(params, formatter = formatIntegerMetric) {
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
        rows.push(`${item.marker || ""} ${displayLabel(item.seriesName)} ${cohortLabel(cohortYear)}: ${isAnonymizedView() ? anonymizedMetricValue() : formatter(Number(y))}`);
      }
    });
  return [tooltipHeader(year), ...rows].join("<br/>");
}

function renderOutputCharts(outputs) {
  state.outputCharts.revenue.setOption(revenuePathChartOption(outputs), true);

  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
  const gapSeries = [{
    name: `${displayScenarioName(activeScenario())} Delta to Plan`,
    type: "bar",
    data: outputs.delta,
    itemStyle: { color: value => value.value < 0 ? "#b42318" : "#4f7f52" },
  }];
  if (comparisonOutputs) {
    gapSeries.push({
      name: `${displayScenarioName(comparison)} Delta to Plan`,
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
          ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(item.value), 0)}`),
        ].join("<br/>");
      },
    },
    grid: { left: 12, right: 18, top: 18, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
    series: gapSeries,
  }, true);
}

function metricTreeSeries(name, data, color, type = "line", extra = {}) {
  return {
    name,
    type,
    data,
    symbolSize: type === "line" ? 4 : undefined,
    lineStyle: type === "line" ? { color, width: 2 } : undefined,
    itemStyle: { color },
    ...extra,
  };
}

const metricTreeScenarioPalette = [
  TOP_DOWN_COLOR,
  BOTTOM_UP_COLOR,
  "#8e6bb0",
  "#c47a5a",
  "#4b9b8e",
  "#9a8c55",
  "#b66f8f",
  "#5875a4",
];

function metricTreeScenarioColor(scenario) {
  if (!scenario) return "#98a2b3";
  const index = Object.values(state.scenarios).findIndex(item => item.id === scenario.id);
  const colorIndex = index >= 0 ? index : 0;
  return metricTreeScenarioPalette[colorIndex % metricTreeScenarioPalette.length];
}

function metricTreeChartOption(series, { format = "number", stacked = false } = {}) {
  const formatter = isAnonymizedView() ? anonymizedMetricValue
    : format === "currency" ? formatCompactCurrency
    : format === "currency2" ? formatCompactCurrency
    : format === "percent" ? value => `${trimNumber(Number(value) * 100, 0)}%`
      : format === "decimal" ? value => trimNumber(Number(value), 2)
        : formatCompactNumber;
  const tooltipFormatter = isAnonymizedView() ? anonymizedMetricValue
    : format === "currency" ? value => formatCurrency(Number(value), 0)
    : format === "currency2" ? value => formatCurrency(Number(value), 2)
    : format === "percent" ? value => `${trimNumber(Number(value) * 100, 1)}%`
      : format === "decimal" ? value => trimNumber(Number(value), 2)
        : value => Math.round(Number(value)).toLocaleString("en-US");
  return {
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${item.seriesName}: ${tooltipFormatter(item.value)}`),
      ].join("<br/>"),
    },
    legend: { top: 0, type: "scroll" },
    grid: { left: 8, right: 10, top: 42, bottom: 24, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String), axisLabel: { fontSize: 10 } },
    yAxis: { type: "value", axisLabel: { formatter, fontSize: 10 } },
    series: stacked
      ? series.map(item => item.type === "bar" ? { ...item, stack: "metricTreeStack" } : item)
      : series,
  };
}

function setMetricTreeChart(key, series, options) {
  state.outputCharts[key]?.setOption(metricTreeChartOption(series, options), true);
}

function profilesPerCustomerValues(outputs) {
  return outputs.paidProfiles.map((value, index) => {
    const customers = outputs.totalCustomers[index];
    return customers > 0 ? value / customers : 0;
  });
}

function revenuePerProfileValues(outputs) {
  return outputs.revenue.map((value, index) => {
    const profiles = outputs.paidProfiles[index];
    return profiles > 0 ? value / profiles : 0;
  });
}

function renderMetricTree(outputs) {
  const scenario = activeScenario();
  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
  const scenarioColor = metricTreeScenarioColor(scenario);
  const comparisonColor = metricTreeScenarioColor(comparison);
  const priorYearPayingCustomers = outputs.totalCustomers.map((value, index) => index === 0 ? null : outputs.totalCustomers[index - 1]);
  const comparisonPriorYearPayingCustomers = comparisonOutputs
    ? comparisonOutputs.totalCustomers.map((value, index) => index === 0 ? null : comparisonOutputs.totalCustomers[index - 1])
    : null;
  const revenuePerCustomer = outputs.revenue.map((value, index) => {
    const customers = outputs.totalCustomers[index];
    return customers > 0 ? value / customers : 0;
  });
  const comparisonRevenuePerCustomer = comparisonOutputs
    ? comparisonOutputs.revenue.map((value, index) => {
      const customers = comparisonOutputs.totalCustomers[index];
      return customers > 0 ? value / customers : 0;
    })
    : null;
  const profilesPerCustomer = profilesPerCustomerValues(outputs);
  const comparisonProfilesPerCustomer = comparisonOutputs ? profilesPerCustomerValues(comparisonOutputs) : null;
  const revenuePerProfile = revenuePerProfileValues(outputs);
  const comparisonRevenuePerProfile = comparisonOutputs ? revenuePerProfileValues(comparisonOutputs) : null;
  const scenarioName = displayScenarioName(scenario);
  const comparisonName = comparison ? displayScenarioName(comparison) : "";

  setMetricTreeChart("treeRevenue", [
    metricTreeSeries(displayLabel(`${scenarioName} HHP`), outputs.revenue, scenarioColor),
    metricTreeSeries("Plan", outputs.planRevenue, "#111827"),
    ...(comparisonOutputs ? [metricTreeSeries(displayLabel(`${comparisonName} HHP`), comparisonOutputs.revenue, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treePayingCustomers", [
    metricTreeSeries(scenarioName, outputs.totalCustomers, scenarioColor),
    ...(comparisonOutputs ? [metricTreeSeries(comparisonName, comparisonOutputs.totalCustomers, comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeNewPayingCustomers", [
    metricTreeSeries(scenarioName, driverValues(scenario, "newCustomers"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, driverValues(comparison, "newCustomers"), comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeExistingPropertyNewPayingCustomers", [
    metricTreeSeries(scenarioName, YEARS.map(visibleExistingPropertyNewCustomersValue), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, YEARS.map(year => existingPropertyNewCustomersTotal(comparison.newCustomerDrilldown.counts, year)), comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeNewPropertyNewPayingCustomers", [
    metricTreeSeries(scenarioName, YEARS.map(visibleNewPropertyNewCustomersValue), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, YEARS.map(year => newPropertyCohortValue(comparison.newCustomerDrilldown.counts, year)), comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeNewUnits", [
    metricTreeSeries(scenarioName, YEARS.map(year => visibleNewCustomerUnitValue("newUnits", newCustomerUnitNewUnitsPairs(scenario), year, newCustomerUnitNewUnitsValue(scenario, year))), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, YEARS.map(year => newCustomerUnitNewUnitsValue(comparison, year)), comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeNewProperties", [
    metricTreeSeries(scenarioName, YEARS.map(year => visibleNewCustomerNewUnitsValue("newProperties", newCustomerNewUnitsNewPropertiesPairs(scenario), year, newCustomerNewUnitsNewPropertiesValue(scenario, year))), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, YEARS.map(year => newCustomerNewUnitsNewPropertiesValue(comparison, year)), comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeNewUnitsPerProperty", [
    metricTreeSeries(scenarioName, YEARS.map(year => visibleNewCustomerNewUnitsValue("unitsPerNewProperty", newCustomerNewUnitsUnitsPerPropertyPairs(scenario), year, newCustomerNewUnitsUnitsPerPropertyValue(scenario, year))), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, YEARS.map(year => newCustomerNewUnitsUnitsPerPropertyValue(comparison, year)), comparisonColor)] : []),
  ], { format: "decimal" });

  setMetricTreeChart("treeNewCustomersPerUnit", [
    metricTreeSeries(scenarioName, YEARS.map(year => visibleNewCustomerUnitValue("customersPerUnit", newCustomerUnitCustomersPerUnitPairs(scenario), year, newCustomerUnitCustomersPerUnitValue(scenario, year))), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, YEARS.map(year => newCustomerUnitCustomersPerUnitValue(comparison, year)), comparisonColor)] : []),
  ], { format: "decimal" });

  setMetricTreeChart("treeReturningPayingCustomers", [
    metricTreeSeries(scenarioName, outputs.returningCustomers, scenarioColor),
    ...(comparisonOutputs ? [metricTreeSeries(comparisonName, comparisonOutputs.returningCustomers, comparisonColor)] : []),
  ]);

  setMetricTreeChart("treePriorPayingCustomers", [
    metricTreeSeries(scenarioName, priorYearPayingCustomers, scenarioColor),
    ...(comparisonPriorYearPayingCustomers ? [metricTreeSeries(comparisonName, comparisonPriorYearPayingCustomers, comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeRetentionRate", [
    metricTreeSeries(scenarioName, driverValues(scenario, "retention"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, driverValues(comparison, "retention"), comparisonColor)] : []),
  ], { format: "percent" });

  setMetricTreeChart("treeRevenuePerCustomer", [
    metricTreeSeries(scenarioName, revenuePerCustomer, scenarioColor),
    ...(comparisonRevenuePerCustomer ? [metricTreeSeries(comparisonName, comparisonRevenuePerCustomer, comparisonColor)] : []),
  ], { format: "currency2" });

  setMetricTreeChart("treeProfilesPerCustomer", [
    metricTreeSeries(scenarioName, profilesPerCustomer, scenarioColor),
    ...(comparisonProfilesPerCustomer ? [metricTreeSeries(comparisonName, comparisonProfilesPerCustomer, comparisonColor)] : []),
  ], { format: "decimal" });

  setMetricTreeChart("treeRevenuePerProfile", [
    metricTreeSeries(scenarioName, revenuePerProfile, scenarioColor),
    ...(comparisonRevenuePerProfile ? [metricTreeSeries(comparisonName, comparisonRevenuePerProfile, comparisonColor)] : []),
  ], { format: "currency2" });

  setMetricTreeChart("treeProfilesReturningCustomer", [
    metricTreeSeries(scenarioName, driverValues(scenario, "profilesReturning"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, driverValues(comparison, "profilesReturning"), comparisonColor)] : []),
  ], { format: "decimal" });

  setMetricTreeChart("treeProfilesNewCustomer", [
    metricTreeSeries(scenarioName, driverValues(scenario, "profilesNew"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, driverValues(comparison, "profilesNew"), comparisonColor)] : []),
  ], { format: "decimal" });

  setMetricTreeChart("treeRevenueReturningProfile", [
    metricTreeSeries(scenarioName, driverValues(scenario, "revReturningProfile"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, driverValues(comparison, "revReturningProfile"), comparisonColor)] : []),
  ], { format: "currency2" });

  setMetricTreeChart("treeRevenueNewProfile", [
    metricTreeSeries(scenarioName, driverValues(scenario, "revNewProfile"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, driverValues(comparison, "revNewProfile"), comparisonColor)] : []),
  ], { format: "currency2" });
}

function secondaryCanvasNodes(focusNode) {
  const focus = canvasNodeTree[focusNode];
  if (!focus) return new Set();
  const secondary = new Set(focus.children || []);
  if (focus.parent) {
    secondary.add(focus.parent);
    (canvasNodeTree[focus.parent]?.children || []).forEach(sibling => {
      if (sibling !== focusNode) secondary.add(sibling);
    });
  }
  return secondary;
}

function canvasNodeFocusRole(nodeId, focusNode, secondary) {
  if (nodeId === "revenue") return "main";
  if (nodeId === focusNode) return "main";
  if (secondary.has(nodeId)) return "secondary";
  return "tertiary";
}

function resizeCanvasCharts() {
  Object.values(state.outputCharts).forEach(chart => chart?.resize());
}

function renderCanvasConnectors(surfaceWidth, surfaceHeight) {
  const svg = document.querySelector(".canvas-connectors");
  if (!svg) return;
  svg.setAttribute("viewBox", `0 0 ${surfaceWidth} ${surfaceHeight}`);
  svg.replaceChildren();
  Object.entries(canvasNodeTree).forEach(([parentId, config]) => {
    const parentNode = document.querySelector(`.canvas-node[data-node-id="${parentId}"]`);
    if (!parentNode) return;
    const startX = parentNode.offsetLeft + parentNode.offsetWidth / 2;
    const startY = parentNode.offsetTop + parentNode.offsetHeight;
    (config.children || []).forEach(childId => {
      const childNode = document.querySelector(`.canvas-node[data-node-id="${childId}"]`);
      if (!childNode) return;
      const endX = childNode.offsetLeft + childNode.offsetWidth / 2;
      const endY = childNode.offsetTop;
      const midpointY = startY + Math.max(70, (endY - startY) / 2);
      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("d", `M ${startX} ${startY} C ${startX} ${midpointY}, ${endX} ${midpointY}, ${endX} ${endY}`);
      svg.appendChild(path);
    });
  });
}

function canvasZoomValue(value) {
  const rounded = Math.round(value * 100) / 100;
  return Math.min(CANVAS_ZOOM_MAX, Math.max(CANVAS_ZOOM_MIN, rounded));
}

function renderCanvasZoomControls() {
  const label = document.getElementById("canvas-zoom-label");
  const zoomOut = document.getElementById("canvas-zoom-out");
  const zoomIn = document.getElementById("canvas-zoom-in");
  if (label) label.textContent = `${Math.round(state.canvasZoom * 100)}%`;
  if (zoomOut) zoomOut.disabled = state.canvasZoom <= CANVAS_ZOOM_MIN;
  if (zoomIn) zoomIn.disabled = state.canvasZoom >= CANVAS_ZOOM_MAX;
}

function applyCanvasZoom(surfaceWidth, surfaceHeight) {
  const surface = document.querySelector(".canvas-surface");
  const viewport = document.querySelector(".canvas-viewport");
  if (!surface || !viewport) return;
  const zoom = state.canvasZoom;
  surface.style.transform = `scale(${zoom})`;
  viewport.style.width = `${surfaceWidth * zoom}px`;
  viewport.style.height = `${surfaceHeight * zoom}px`;
  renderCanvasZoomControls();
}

function layoutCanvasNodes() {
  const surface = document.querySelector(".canvas-surface");
  if (!surface) return;
  const focusNode = canvasNodeTree[state.canvasFocusNode] ? state.canvasFocusNode : "revenue";
  const secondary = secondaryCanvasNodes(focusNode);
  const marginX = 56;
  const marginY = 64;
  const gapX = 58;
  const rowGap = 84;
  const rowLayouts = canvasLayoutRows.map(row => {
    const items = row.map(nodeId => {
      const role = canvasNodeFocusRole(nodeId, focusNode, secondary);
      return { nodeId, role, ...canvasFocusDimensions[role] };
    });
    return {
      items,
      width: items.reduce((total, item) => total + item.width, 0) + Math.max(0, items.length - 1) * gapX,
      height: Math.max(...items.map(item => item.height)),
    };
  });
  const surfaceWidth = Math.max(1500, Math.max(...rowLayouts.map(row => row.width)) + marginX * 2);
  let top = marginY;
  rowLayouts.forEach(row => {
    let left = (surfaceWidth - row.width) / 2;
    row.items.forEach(item => {
      const node = document.querySelector(`.canvas-node[data-node-id="${item.nodeId}"]`);
      if (!node) return;
      node.style.left = `${left}px`;
      node.style.top = `${top}px`;
      left += item.width + gapX;
    });
    top += row.height + rowGap;
  });
  const surfaceHeight = top + marginY - rowGap;
  surface.style.width = `${surfaceWidth}px`;
  surface.style.height = `${surfaceHeight}px`;
  applyCanvasZoom(surfaceWidth, surfaceHeight);
  renderCanvasConnectors(surfaceWidth, surfaceHeight);
}

function refreshCanvasLayoutAndCharts() {
  layoutCanvasNodes();
  resizeCanvasCharts();
}

function renderCanvasFocus() {
  const focusNode = canvasNodeTree[state.canvasFocusNode] ? state.canvasFocusNode : "revenue";
  state.canvasFocusNode = focusNode;
  const secondary = secondaryCanvasNodes(focusNode);
  document.querySelectorAll(".canvas-node[data-node-id]").forEach(node => {
    const nodeId = node.dataset.nodeId;
    const isMain = nodeId === focusNode;
    const hasMainSizing = isMain || nodeId === "revenue";
    const isSecondary = !hasMainSizing && secondary.has(nodeId);
    node.classList.toggle("canvas-node-focus-main", hasMainSizing);
    node.classList.toggle("canvas-node-focus-secondary", isSecondary);
    node.classList.toggle("canvas-node-focus-tertiary", !hasMainSizing && !isSecondary);
    node.setAttribute("aria-current", isMain ? "true" : "false");
  });
  layoutCanvasNodes();
  requestAnimationFrame(refreshCanvasLayoutAndCharts);
  setTimeout(refreshCanvasLayoutAndCharts, 190);
}

function setCanvasFocus(nodeId) {
  if (!canvasNodeTree[nodeId]) return;
  state.canvasFocusNode = nodeId;
  renderCanvasFocus();
}

function setCanvasZoom(value, focalPoint = null) {
  const stage = document.querySelector(".canvas-stage");
  const previousZoom = state.canvasZoom;
  const focal = focalPoint && stage
    ? {
      x: (stage.scrollLeft + focalPoint.x) / previousZoom,
      y: (stage.scrollTop + focalPoint.y) / previousZoom,
      viewportX: focalPoint.x,
      viewportY: focalPoint.y,
    }
    : null;
  const nextZoom = canvasZoomValue(value);
  if (nextZoom === state.canvasZoom) return;
  state.canvasZoom = nextZoom;
  refreshCanvasLayoutAndCharts();
  if (stage && focal) {
    stage.scrollLeft = (focal.x * nextZoom) - focal.viewportX;
    stage.scrollTop = (focal.y * nextZoom) - focal.viewportY;
  }
}

function scrollCanvasToNode(nodeId) {
  const stage = document.querySelector(".canvas-stage");
  const node = document.querySelector(`.canvas-node[data-node-id="${nodeId}"]`);
  if (!stage || !node) return;
  const zoom = state.canvasZoom;
  const nodeCenterX = (node.offsetLeft + node.offsetWidth / 2) * zoom;
  const nodeCenterY = (node.offsetTop + node.offsetHeight / 2) * zoom;
  const maxLeft = Math.max(0, stage.scrollWidth - stage.clientWidth);
  const maxTop = Math.max(0, stage.scrollHeight - stage.clientHeight);
  stage.scrollLeft = Math.min(maxLeft, Math.max(0, nodeCenterX - stage.clientWidth / 2));
  stage.scrollTop = Math.min(maxTop, Math.max(0, nodeCenterY - stage.clientHeight / 2));
}

function centerCanvasOnNode(nodeId) {
  requestAnimationFrame(() => {
    scrollCanvasToNode(nodeId);
    requestAnimationFrame(() => scrollCanvasToNode(nodeId));
  });
  setTimeout(() => scrollCanvasToNode(nodeId), 220);
}

function enterMetricTreeCanvas() {
  state.canvasZoom = CANVAS_ZOOM_MAX;
  state.canvasFocusNode = "revenue";
  renderCanvasFocus();
  centerCanvasOnNode("revenue");
}

function bindCanvasTrackpadZoom() {
  const stage = document.querySelector(".canvas-stage");
  if (!stage) return;
  stage.addEventListener("wheel", event => {
    if (!event.ctrlKey && !event.metaKey) return;
    event.preventDefault();
    const rect = stage.getBoundingClientRect();
    const delta = -event.deltaY * 0.0025;
    setCanvasZoom(state.canvasZoom + delta, {
      x: event.clientX - rect.left,
      y: event.clientY - rect.top,
    });
  }, { passive: false });
}

function revenuePathChartOption(outputs) {
  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
  const scenarioName = displayScenarioName(activeScenario());
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const revenueSeries = [
    { name: `${scenarioName} Existing Customers`, type: "bar", stack: "revenue", data: outputs.revenueExisting, itemStyle: { color: "#4f7fb8" } },
    { name: `${scenarioName} New Customers`, type: "bar", stack: "revenue", data: outputs.revenueNew, itemStyle: { color: "#8ebf86" } },
    { name: "Plan Revenue", type: "line", data: outputs.planRevenue, symbolSize: 6, lineStyle: { color: "#111827", width: 2 }, itemStyle: { color: "#111827" } },
  ];
  if (comparisonOutputs) {
    revenueSeries.push({
      name: `${comparisonName} Revenue`,
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
          .filter(item => item.seriesType === "bar" && item.seriesName.startsWith(scenarioName))
          .reduce((sum, item) => sum + Number(item.value || 0), 0);
        lines.push(`<strong>${scenarioName} Revenue: ${formatCurrency(total, 0)}</strong>`);
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
  document.getElementById("kpi-cagr").textContent = isAnonymizedView() ? anonymizedMetricValue() : `${(cagr * 100).toFixed(1)}%`;
  document.getElementById("kpi-paid-profiles").textContent = formatIntegerMetric(outputs.paidProfiles[last]);
}

function setText(id, value) {
  const element = document.getElementById(id);
  if (element) element.textContent = value;
}

function renderGuidedPlan(outputs) {
  const year = 2025;
  const index = YEARS.indexOf(year);
  if (index < 0) return;
  const scenario = activeScenario();
  const drivers = scenario.drivers;
  const newCustomers = effectiveNewCustomers(scenario)[index];
  const newPaidProfiles = newCustomers * drivers.profilesNew[index];
  const totalRevenue = outputs.revenue[index];
  const revPerPaidProfile = outputs.paidProfiles[index] ? totalRevenue / outputs.paidProfiles[index] : 0;

  setText("guided-existing-revenue", formatCurrency(outputs.revenueExisting[index], 0));
  setText("guided-new-revenue", formatCurrency(outputs.revenueNew[index], 0));
  setText("guided-total-revenue", formatCurrency(totalRevenue, 0));
  setText("guided-retention", formatValue(drivers.retention[index], "percent"));
  setText("guided-returning-customers", Math.round(outputs.returningCustomers[index]).toLocaleString("en-US"));
  setText("guided-existing-profiles-customer", formatValue(drivers.profilesReturning[index], "decimal"));
  setText("guided-existing-rev-profile", formatValue(drivers.revReturningProfile[index], "currency2"));
  setText("guided-new-customers", Math.round(newCustomers).toLocaleString("en-US"));
  setText("guided-new-profiles-customer", formatValue(drivers.profilesNew[index], "decimal"));
  setText("guided-new-rev-profile", formatValue(drivers.revNewProfile[index], "currency2"));
  setText("guided-new-paid-profiles", Math.round(newPaidProfiles).toLocaleString("en-US"));
  setText("guided-plan-revenue", formatCurrency(outputs.planRevenue[index], 0));
  setText("guided-plan-gap", formatCurrency(outputs.delta[index], 0));
  document.getElementById("guided-plan-gap")?.classList.toggle("negative", outputs.delta[index] < 0);
  setText("guided-paid-profiles", Math.round(outputs.paidProfiles[index]).toLocaleString("en-US"));
  setText("guided-rev-paid-profile", formatCurrency(revPerPaidProfile, 2));
  renderGuidedOutline(outputs);
}

function renderGuidedOutline(outputs) {
  const table = document.getElementById("guided-outline-table");
  if (!table) return;
  const outlineYears = YEARS.filter(year => year >= 2025);
  table.innerHTML = `
    <thead>
      <tr>
        <th>Year</th>
        <th>Existing Revenue</th>
        <th>New Revenue</th>
        <th>HHP Revenue</th>
        <th>Plan</th>
        <th>Gap</th>
      </tr>
    </thead>
    <tbody>
      ${outlineYears.map(year => {
        const yearIndex = YEARS.indexOf(year);
        const gap = outputs.delta[yearIndex];
        return `
          <tr>
            <td>${year}</td>
            <td>${formatCurrency(outputs.revenueExisting[yearIndex], 0)}</td>
            <td>${formatCurrency(outputs.revenueNew[yearIndex], 0)}</td>
            <td><strong>${formatCurrency(outputs.revenue[yearIndex], 0)}</strong></td>
            <td>${formatCurrency(outputs.planRevenue[yearIndex], 0)}</td>
            <td class="${gap < 0 ? "negative" : ""}">${formatCurrency(gap, 0)}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;
}

function renderTable(outputs) {
  const table = document.getElementById("driver-table");
  const driverRows = Object.entries(driverMeta).map(([key, meta]) => ({
    key,
    label: meta.label,
    format: meta.format,
    values: driverValues(activeScenario(), key),
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

function renderNewCustomerDrilldown({ syncEditors = true, excludeEditor = null } = {}) {
  renderNewCustomerSourceControl();
  renderNewCustomerEditorMode();
  renderNewCustomerUnitDrilldownToggle();
  renderCohortLockSummary();
  renderCohortLockControls();
  renderNewCustomerCohortEditorControl();
  renderNewCustomerYearEditorControl();
  renderNewCustomerDrilldownCharts();
  renderNewCustomerUnitDrilldown();
  renderNewCustomerAllCohortsChart();
  if (syncEditors) syncNewCustomerCohortEditors({ excludeChart: excludeEditor });
  renderNewCustomerDrilldownTable();
  renderNewCustomerYoyTable();
}

function renderNewCustomerUnitDrilldownToggle() {
  const panel = document.getElementById("new-customers-unit-drilldown");
  const drilldownButton = document.getElementById("toggle-new-customers-unit-drilldown");
  const backButton = document.getElementById("back-new-customers-unit-drilldown");
  const newUnitsPanel = document.getElementById("new-customers-new-units-drilldown");
  const newUnitsDrilldownButton = document.getElementById("toggle-new-customers-new-units-drilldown");
  const newUnitsBackButton = document.getElementById("back-new-customers-new-units-drilldown");
  const view = document.getElementById("new-customers-view");
  if (!panel || !drilldownButton || !backButton || !view || !newUnitsPanel || !newUnitsDrilldownButton || !newUnitsBackButton) return;
  if (!state.newCustomerUnitDrilldownOpen) state.newCustomerNewUnitsDrilldownOpen = false;
  drilldownButton.hidden = state.currentView !== "newCustomers" || state.newCustomerUnitDrilldownOpen;
  backButton.hidden = !state.newCustomerUnitDrilldownOpen || state.newCustomerNewUnitsDrilldownOpen;
  newUnitsDrilldownButton.hidden = !state.newCustomerUnitDrilldownOpen || state.newCustomerNewUnitsDrilldownOpen;
  newUnitsBackButton.hidden = !state.newCustomerNewUnitsDrilldownOpen;
  panel.hidden = !state.newCustomerUnitDrilldownOpen;
  newUnitsPanel.hidden = !state.newCustomerNewUnitsDrilldownOpen;
  view.classList.toggle("new-customers-unit-mode", state.newCustomerUnitDrilldownOpen);
  view.classList.toggle("new-customers-new-units-mode", state.newCustomerNewUnitsDrilldownOpen);
}

function resizeNewCustomerUnitDrilldownCharts() {
  state.outputCharts.newCustomerUnitBridge?.resize();
  state.outputCharts.newCustomerUnitDelta?.resize();
  state.cohortCharts.unitNewUnits?.resize();
  state.cohortCharts.unitCustomersPerUnit?.resize();
}

function resizeNewCustomerNewUnitsDrilldownCharts() {
  state.outputCharts.newCustomerNewUnitsBridge?.resize();
  state.outputCharts.newCustomerNewUnitsDelta?.resize();
  state.cohortCharts.newUnitsProperties?.resize();
  state.cohortCharts.newUnitsUnitsPerProperty?.resize();
}

function renderCohortLockSummary() {
  const summary = document.getElementById("cohort-lock-summary");
  if (!summary) return;
  const locked = Array.from(lockedNewCustomerCohorts()).sort((left, right) => left - right);
  if (!locked.length) {
    summary.textContent = "All cohorts are unlocked. Bridge changes can flow through every cohort.";
    return;
  }
  summary.textContent = `${locked.length} locked: ${locked.map(cohortLabel).join(", ")}`;
}

function renderCohortLockControls() {
  const list = document.getElementById("cohort-lock-list");
  if (!list) return;
  const locked = lockedNewCustomerCohorts();
  list.innerHTML = NEW_CUSTOMER_COHORT_YEARS.map(cohortYear => `
    <label class="cohort-lock-item">
      <input type="checkbox" data-lock-cohort="${cohortYear}" ${locked.has(Number(cohortYear)) ? "checked" : ""} />
      <span>${cohortLabel(cohortYear)}</span>
    </label>
  `).join("");
  list.querySelectorAll("input[data-lock-cohort]").forEach(input => {
    input.addEventListener("change", saveCohortLocks);
  });
}

function saveCohortLocks() {
  const locked = Array.from(document.querySelectorAll("#cohort-lock-list input[data-lock-cohort]:checked"))
    .map(input => Number(input.dataset.lockCohort))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
  const current = Array.from(lockedNewCustomerCohorts()).sort((left, right) => left - right);
  if (sameNumberList(current, locked)) return;
  pushUndoSnapshot();
  activeScenario().newCustomerDrilldown.lockedCohorts = locked;
  saveScenarios();
  renderCohortLockSummary();
  syncNewCustomerCohortEditors();
}

function setCohortLockSelection(predicate) {
  document.querySelectorAll("#cohort-lock-list input[data-lock-cohort]").forEach(input => {
    input.checked = predicate(Number(input.dataset.lockCohort));
  });
  saveCohortLocks();
}

function newCustomerUnitPlan(scenario = activeScenario()) {
  if (!scenario.newCustomerDrilldown.unitPlan) {
    scenario.newCustomerDrilldown.unitPlan = createDefaultNewCustomerUnitPlan(scenario.newCustomerDrilldown.counts);
  }
  return scenario.newCustomerDrilldown.unitPlan;
}

function newCustomerUnitNewUnitsValue(scenario, year) {
  const plan = newCustomerUnitPlan(scenario);
  return Number(plan.newUnits?.[String(year)] ?? plan.newUnits?.[year] ?? 0);
}

function newCustomerUnitCustomersPerUnitValue(scenario, year) {
  const plan = newCustomerUnitPlan(scenario);
  return Number(plan.customersPerUnit?.[String(year)] ?? plan.customersPerUnit?.[year] ?? 0);
}

function newCustomerUnitBottomUpValue(scenario, year) {
  return newCustomerUnitNewUnitsValue(scenario, year) * newCustomerUnitCustomersPerUnitValue(scenario, year);
}

function setNewCustomerUnitNewUnits(year, value, { exact = false } = {}) {
  const plan = newCustomerUnitPlan();
  if (!plan.newUnits) plan.newUnits = {};
  plan.newUnits[String(year)] = exact ? Math.max(0, value) : Math.max(0, Math.round(value));
}

function setNewCustomerUnitCustomersPerUnit(year, value, { exact = false } = {}) {
  const plan = newCustomerUnitPlan();
  if (!plan.customersPerUnit) plan.customersPerUnit = {};
  plan.customersPerUnit[String(year)] = exact ? Math.max(0, value) : Math.max(0, Math.round(value * 10000) / 10000);
}

function newCustomerUnitControlPoints() {
  const plan = newCustomerUnitPlan();
  if (!plan.controlPoints) plan.controlPoints = {};
  return plan.controlPoints;
}

function controlledNewCustomerUnitPairs(key, pairs) {
  const stored = newCustomerUnitControlPoints()[key];
  if (!Array.isArray(stored)) return pairs;
  const visibleX = new Set(stored.map(Number));
  const visiblePairs = pairs.filter(([x]) => visibleX.has(Number(x)));
  if (visiblePairs.length >= 2) return visiblePairs;
  if (pairs.length <= 2) return pairs;
  return [pairs[0], pairs[pairs.length - 1]];
}

function rememberNewCustomerUnitControlPoints(key, data) {
  newCustomerUnitControlPoints()[key] = (data || [])
    .map(([x]) => Number(x))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
}

function ensureControlPointsForYears(controlPoints, key, years) {
  const existing = Array.isArray(controlPoints[key]) ? controlPoints[key] : YEARS;
  controlPoints[key] = Array.from(new Set([
    ...existing.map(Number).filter(Number.isFinite),
    ...years.map(Number).filter(Number.isFinite),
  ])).sort((left, right) => left - right);
}

function newCustomerUnitNewUnitsPairs(scenario = activeScenario()) {
  return YEARS.map(year => [year, newCustomerUnitNewUnitsValue(scenario, year)]);
}

function newCustomerUnitCustomersPerUnitPairs(scenario = activeScenario()) {
  return YEARS.map(year => [year, newCustomerUnitCustomersPerUnitValue(scenario, year)]);
}

function visibleNewCustomerUnitValue(key, pairs, year, fallback) {
  const value = interpolateNapkinLineValue(controlledNewCustomerUnitPairs(key, pairs), year);
  return value === null ? fallback : value;
}

function visibleNewCustomerUnitBottomUpValue(year) {
  const newUnits = visibleNewCustomerUnitValue(
    "newUnits",
    newCustomerUnitNewUnitsPairs(activeScenario()),
    year,
    newCustomerUnitNewUnitsValue(activeScenario(), year)
  );
  const customersPerUnit = visibleNewCustomerUnitValue(
    "customersPerUnit",
    newCustomerUnitCustomersPerUnitPairs(activeScenario()),
    year,
    newCustomerUnitCustomersPerUnitValue(activeScenario(), year)
  );
  return newUnits * customersPerUnit;
}

function newCustomerNewUnitsPropertyPlan(scenario = activeScenario()) {
  const plan = newCustomerUnitPlan(scenario);
  if (!plan.propertyPlan) {
    plan.propertyPlan = createDefaultNewCustomerNewUnitsPropertyPlan(plan.newUnits);
  }
  return plan.propertyPlan;
}

function newCustomerNewUnitsNewPropertiesValue(scenario, year) {
  const plan = newCustomerNewUnitsPropertyPlan(scenario);
  return Number(plan.newProperties?.[String(year)] ?? plan.newProperties?.[year] ?? 0);
}

function newCustomerNewUnitsUnitsPerPropertyValue(scenario, year) {
  const plan = newCustomerNewUnitsPropertyPlan(scenario);
  return Number(plan.unitsPerNewProperty?.[String(year)] ?? plan.unitsPerNewProperty?.[year] ?? 0);
}

function newCustomerNewUnitsBottomUpValue(scenario, year) {
  return newCustomerNewUnitsNewPropertiesValue(scenario, year) * newCustomerNewUnitsUnitsPerPropertyValue(scenario, year);
}

function setNewCustomerNewUnitsNewProperties(year, value, { exact = false } = {}) {
  const plan = newCustomerNewUnitsPropertyPlan();
  if (!plan.newProperties) plan.newProperties = {};
  plan.newProperties[String(year)] = exact ? Math.max(0, value) : Math.max(0, Math.round(value));
}

function setNewCustomerNewUnitsUnitsPerProperty(year, value, { exact = false } = {}) {
  const plan = newCustomerNewUnitsPropertyPlan();
  if (!plan.unitsPerNewProperty) plan.unitsPerNewProperty = {};
  plan.unitsPerNewProperty[String(year)] = exact ? Math.max(0, value) : Math.max(0, Math.round(value * 100) / 100);
}

function newCustomerNewUnitsControlPoints() {
  const plan = newCustomerNewUnitsPropertyPlan();
  if (!plan.controlPoints) plan.controlPoints = {};
  return plan.controlPoints;
}

function controlledNewCustomerNewUnitsPairs(key, pairs) {
  const stored = newCustomerNewUnitsControlPoints()[key];
  if (!Array.isArray(stored)) return pairs;
  const visibleX = new Set(stored.map(Number));
  const visiblePairs = pairs.filter(([x]) => visibleX.has(Number(x)));
  if (visiblePairs.length >= 2) return visiblePairs;
  if (pairs.length <= 2) return pairs;
  return [pairs[0], pairs[pairs.length - 1]];
}

function rememberNewCustomerNewUnitsControlPoints(key, data) {
  newCustomerNewUnitsControlPoints()[key] = (data || [])
    .map(([x]) => Number(x))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
}

function visibleNewCustomerNewUnitsValue(key, pairs, year, fallback) {
  const value = interpolateNapkinLineValue(controlledNewCustomerNewUnitsPairs(key, pairs), year);
  return value === null ? fallback : value;
}

function visibleNewCustomerNewUnitsBottomUpValue(year) {
  const newProperties = visibleNewCustomerNewUnitsValue(
    "newProperties",
    newCustomerNewUnitsNewPropertiesPairs(activeScenario()),
    year,
    newCustomerNewUnitsNewPropertiesValue(activeScenario(), year)
  );
  const unitsPerProperty = visibleNewCustomerNewUnitsValue(
    "unitsPerNewProperty",
    newCustomerNewUnitsUnitsPerPropertyPairs(activeScenario()),
    year,
    newCustomerNewUnitsUnitsPerPropertyValue(activeScenario(), year)
  );
  return newProperties * unitsPerProperty;
}

function newCustomerNewUnitsNewPropertiesPairs(scenario = activeScenario()) {
  return YEARS.map(year => [year, newCustomerNewUnitsNewPropertiesValue(scenario, year)]);
}

function newCustomerNewUnitsUnitsPerPropertyPairs(scenario = activeScenario()) {
  return YEARS.map(year => [year, newCustomerNewUnitsUnitsPerPropertyValue(scenario, year)]);
}

function revUnitNewUnitsValue(scenario, year) {
  return Number(scenario.revUnitPlan?.newUnits?.[String(year)] ?? scenario.revUnitPlan?.newUnits?.[year] ?? 0);
}

function revUnitCumulativeNewUnits(scenario, year) {
  return REV_UNIT_YEARS
    .filter(item => item <= year)
    .reduce((sum, item) => sum + revUnitNewUnitsValue(scenario, item), 0);
}

function revUnitRevenueValue(scenario, year) {
  let total = 0;
  let hasValue = false;
  REV_UNIT_YEARS.forEach(cohortYear => {
    if (cohortYear > year) return;
    const revPerUnit = revUnitRpuValue(scenario, year, cohortYear);
    if (revPerUnit === null) return;
    total += revUnitNewUnitsValue(scenario, cohortYear) * revPerUnit;
    hasValue = true;
  });
  if (hasValue) return total;
  const value = scenario.revUnitPlan?.revenue?.[String(year)] ?? scenario.revUnitPlan?.revenue?.[year];
  return Number.isFinite(Number(value)) ? Number(value) : null;
}

function revUnitRpuValue(scenario, year, cohortYear) {
  const row = scenario.revUnitPlan?.revPerUnit?.[String(year)] ?? scenario.revUnitPlan?.revPerUnit?.[year] ?? {};
  const value = row[String(cohortYear)] ?? row[cohortYear];
  if (Number.isFinite(Number(value))) return Number(value);
  if (cohortYear > year) return null;

  for (let candidateYear = year - 1; candidateYear >= cohortYear; candidateYear -= 1) {
    const candidateRow = scenario.revUnitPlan?.revPerUnit?.[String(candidateYear)] ?? scenario.revUnitPlan?.revPerUnit?.[candidateYear] ?? {};
    const candidate = candidateRow[String(cohortYear)] ?? candidateRow[cohortYear];
    if (Number.isFinite(Number(candidate))) return Number(candidate);
  }

  const latestNewCohortRow = scenario.revUnitPlan?.revPerUnit?.[String(REV_UNIT_LAST_HISTORICAL_YEAR)] ?? {};
  const latestNewCohort = latestNewCohortRow[String(REV_UNIT_LAST_HISTORICAL_YEAR)];
  return Number.isFinite(Number(latestNewCohort)) ? Number(latestNewCohort) : null;
}

function unaffiliatedRevenueValue(scenario, year) {
  const value = scenario.revUnitPlan?.unaffiliatedRevenue?.[String(year)] ?? scenario.revUnitPlan?.unaffiliatedRevenue?.[year];
  if (value && typeof value === "object") {
    return Number.isFinite(Number(value.totalRevenue)) ? Number(value.totalRevenue) : null;
  }
  return Number.isFinite(Number(value)) ? Number(value) : null;
}

function revUnitTotalRevenueValue(scenario, year) {
  const affiliated = revUnitRevenueValue(scenario, year);
  const unaffiliated = unaffiliatedRevenueValue(scenario, year);
  if (affiliated === null && unaffiliated === null) return null;
  return (affiliated || 0) + (unaffiliated || 0);
}

function revUnitAggregateRpuValue(scenario, year) {
  const revenue = revUnitTotalRevenueValue(scenario, year);
  const cumulativeUnits = revUnitCumulativeNewUnits(scenario, year);
  if (revenue === null || !cumulativeUnits) return null;
  return revenue / cumulativeUnits;
}

function setRevUnitNewUnits(year, value) {
  if (!activeScenario().revUnitPlan) activeScenario().revUnitPlan = createDefaultRevUnitPlan();
  if (!activeScenario().revUnitPlan.newUnits) activeScenario().revUnitPlan.newUnits = {};
  activeScenario().revUnitPlan.newUnits[String(year)] = value;
}

function setRevUnitRpu(year, cohortYear, value) {
  if (!activeScenario().revUnitPlan) activeScenario().revUnitPlan = createDefaultRevUnitPlan();
  if (!activeScenario().revUnitPlan.revPerUnit) activeScenario().revUnitPlan.revPerUnit = {};
  if (!activeScenario().revUnitPlan.revPerUnit[String(year)]) {
    activeScenario().revUnitPlan.revPerUnit[String(year)] = {};
  }
  activeScenario().revUnitPlan.revPerUnit[String(year)][String(cohortYear)] = Math.max(0, Math.round(value * 100) / 100);
}

function revUnitNewUnitPairs(scenario) {
  return REV_UNIT_YEARS.map(year => [year, revUnitNewUnitsValue(scenario, year)]);
}

function controlledRevUnitPairs(key, pairs) {
  const stored = activeScenario().revUnitPlan?.controlPoints?.[key];
  if (!Array.isArray(stored)) return pairs;
  const visibleX = new Set(stored.map(Number));
  return pairs.filter(([x]) => visibleX.has(Number(x)));
}

function addRevUnitControlPoint(key, x) {
  if (!Number.isFinite(Number(x))) return;
  if (!activeScenario().revUnitPlan) activeScenario().revUnitPlan = createDefaultRevUnitPlan();
  if (!activeScenario().revUnitPlan.controlPoints) activeScenario().revUnitPlan.controlPoints = {};
  const existing = activeScenario().revUnitPlan.controlPoints[key];
  const next = new Set(Array.isArray(existing) ? existing.map(Number) : []);
  next.add(Number(x));
  activeScenario().revUnitPlan.controlPoints[key] = Array.from(next).sort((left, right) => left - right);
}

function rememberRevUnitControlPoints(key, data) {
  if (!activeScenario().revUnitPlan) activeScenario().revUnitPlan = createDefaultRevUnitPlan();
  if (!activeScenario().revUnitPlan.controlPoints) activeScenario().revUnitPlan.controlPoints = {};
  activeScenario().revUnitPlan.controlPoints[key] = (data || [])
    .map(([x]) => Number(x))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
}

function revUnitForecastYears() {
  return REV_UNIT_YEARS.filter(year => year > REV_UNIT_LAST_HISTORICAL_YEAR);
}

function revUnitEditorItems() {
  if (state.revUnitEditorMode === "allCohorts") return [];
  return REV_UNIT_YEARS;
}

function selectedRevUnitEditorValue() {
  return state.revUnitEditorCohortYear;
}

function setSelectedRevUnitEditorValue(value) {
  state.revUnitEditorCohortYear = value;
}

function revUnitRpuControlKey() {
  if (state.revUnitEditorMode === "allCohorts") return "rpu:allCohorts";
  return `rpu:cohort:${state.revUnitEditorCohortYear}`;
}

function revUnitCohortColor(cohortYear) {
  const colors = ["#4f7fb8", "#6fa76b", "#d39b2a", "#7b6fb8", "#a96c50", "#4b9b8e", "#8e6bb0", "#c47a5a", "#5f9ea0", "#9a8c55", "#5875a4", "#76a46b", "#b66f8f"];
  return colors[Math.max(0, REV_UNIT_YEARS.indexOf(cohortYear)) % colors.length];
}

function revUnitRpuPairs(scenario) {
  const cohortYear = state.revUnitEditorCohortYear;
  return REV_UNIT_YEARS
    .filter(year => year >= cohortYear)
    .map(year => [year, revUnitRpuValue(scenario, year, cohortYear)])
    .filter(([, value]) => value !== null);
}

function revUnitRpuEditRanges() {
  return [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]];
}

function revUnitRpuLine() {
  const editRanges = revUnitRpuEditRanges();
  return {
    name: displayLabel("Rev / Unit"),
    color: BOTTOM_UP_COLOR,
    editDomain: {
      moveX: editRanges,
      addX: editRanges,
      deleteX: editRanges,
    },
    data: controlledRevUnitPairs(revUnitRpuControlKey(), revUnitRpuPairs(activeScenario())),
  };
}

function revUnitReferenceLine() {
  const cohortYear = state.revUnitReferenceCohortYear;
  const editedCohortYear = state.revUnitEditorCohortYear;
  const data = REV_UNIT_YEARS
    .filter(year => year >= cohortYear)
    .map(year => {
      const x = state.revUnitReferenceShifted
        ? editedCohortYear + (year - cohortYear)
        : year;
      return [x, revUnitRpuValue(activeScenario(), year, cohortYear)];
    })
    .filter(([x]) => x >= REV_UNIT_YEARS[0] && x <= REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1])
    .filter(([, value]) => value !== null);
  return {
    name: displayLabel(`Reference ${cohortYear}${state.revUnitReferenceShifted ? " Age-Matched" : ""}`),
    color: "#98a2b3",
    editable: false,
    data,
  };
}

function revUnitAllCohortLines() {
  return REV_UNIT_YEARS.map(cohortYear => {
    const editStart = Math.max(FIRST_FORECAST_YEAR, cohortYear);
    const editRanges = editStart <= REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]
      ? [[editStart, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]]
      : [];
    const pairs = REV_UNIT_YEARS
      .filter(year => year >= cohortYear)
      .map(year => [year, revUnitRpuValue(activeScenario(), year, cohortYear)])
      .filter(([, value]) => value !== null);
    return {
      name: String(cohortYear),
      cohortYear,
      color: revUnitCohortColor(cohortYear),
      editDomain: {
        moveX: editRanges,
        addX: editRanges,
        deleteX: editRanges,
      },
      data: controlledRevUnitPairs(`rpu:allCohorts:${cohortYear}`, pairs),
    };
  });
}

function revUnitAllCohortsChartOption() {
  return {
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params
          .filter(item => item.value !== null && item.value !== undefined)
          .map(item => `${item.marker} Cohort ${item.seriesName}: ${formatCurrency(Number(item.value), 2)}`),
      ].join("<br/>"),
    },
    legend: {
      type: "scroll",
      top: 0,
    },
    grid: { left: 12, right: 18, top: 48, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: REV_UNIT_YEARS.map(String) },
    yAxis: { type: "value", min: 0, axisLabel: { formatter: value => formatCurrency(Number(value), 0) } },
    series: REV_UNIT_YEARS.map(cohortYear => ({
      name: String(cohortYear),
      type: "line",
      data: REV_UNIT_YEARS.map(year => {
        if (year < cohortYear) return null;
        return revUnitRpuValue(activeScenario(), year, cohortYear);
      }),
      connectNulls: false,
      symbolSize: 5,
      lineStyle: { color: revUnitCohortColor(cohortYear), width: 2 },
      itemStyle: { color: revUnitCohortColor(cohortYear) },
      emphasis: {
        focus: "series",
        blurScope: "coordinateSystem",
        lineStyle: { width: 5 },
        itemStyle: { borderColor: "#111827", borderWidth: 2 },
      },
      blur: {
        lineStyle: { opacity: 0.14 },
        itemStyle: { opacity: 0.14 },
      },
    })),
  };
}

function revUnitRpuLines() {
  return [revUnitRpuLine(), revUnitReferenceLine()];
}

function revUnitRpuAxisConfig() {
  return {
    min: REV_UNIT_YEARS[0],
    max: REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1],
    formatter: formatAxisYear,
  };
}

function initRevUnitCharts() {
  const unitsChart = new NapkinChart(
    "rev-unit-units-chart",
    [{
      name: displayLabel("New Units"),
      color: TOP_DOWN_COLOR,
      editDomain: {
        moveX: [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]],
      },
      data: controlledRevUnitPairs("newUnits", revUnitNewUnitPairs(activeScenario())),
    }],
    true,
    {
      animation: false,
      xAxis: {
        type: "value",
        min: REV_UNIT_YEARS[0],
        max: REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1],
        minInterval: 1,
        axisLabel: { formatter: formatAxisYear },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: 2000000,
        axisLabel: { formatter: formatCompactNumber },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const lineItem = items.find(item => item && item.seriesType === "line" && item.seriesIndex % 2 === 1) || items[0];
          const data = Array.isArray(lineItem?.data) ? lineItem.data : lineItem?.value;
          const year = Array.isArray(data) ? data[0] : lineItem?.axisValue;
          const value = Array.isArray(data) ? Number(data[1]) : Number(lineItem?.value);
          return [
            tooltipHeader(year),
            `${lineItem?.marker || ""} ${displayLabel("New Units")}: ${formatIntegerMetric(value)}`,
          ].join("<br/>");
        },
      },
    },
    "none",
    false
  );

  unitsChart.windowStartX = REV_UNIT_YEARS[0];
  unitsChart.windowEndX = REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1];
  unitsChart.globalMaxX = REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1];
  unitsChart._refreshChart();
  unitsChart._appEditSnapshot = null;
  unitsChart._appEditCommitted = false;
  unitsChart.chart.getZr().on("mousedown", () => {
    if (state.syncingRevUnitCharts) return;
    unitsChart._appEditSnapshot = snapshotState();
    unitsChart._appEditLineSnapshot = clone(unitsChart.lines?.[0]?.data || []);
    unitsChart._appEditCommitted = false;
  });
  unitsChart.onDataChanged = () => {
    if (state.syncingRevUnitCharts) return;
    if (unitsChart._appEditSnapshot && !unitsChart._appEditCommitted) {
      pushUndoSnapshot(unitsChart._appEditSnapshot);
      unitsChart._appEditCommitted = true;
    }
    rememberRevUnitControlPoints("newUnits", unitsChart.lines[0].data);
    REV_UNIT_YEARS.forEach(year => {
      if (year <= REV_UNIT_LAST_HISTORICAL_YEAR) return;
      const value = interpolateNapkinLineValue(unitsChart.lines[0].data, year);
      if (value !== null) setRevUnitNewUnits(year, Math.max(0, Math.round(value)));
    });
    saveScenarios();
    renderRevUnitPlan();
  };
  state.revUnitCharts.units = unitsChart;

  const rpuAxis = revUnitRpuAxisConfig();
  const rpuChart = new NapkinChart(
    "rev-unit-rpu-chart",
    revUnitRpuLines(),
    true,
    {
      animation: false,
      xAxis: {
        type: "value",
        min: rpuAxis.min,
        max: rpuAxis.max,
        minInterval: 1,
        axisLabel: { formatter: rpuAxis.formatter },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: 20,
        axisLabel: { formatter: value => formatCurrency(Number(value), 0) },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatCohortNapkinTooltip(params, value => formatCurrency(Number(value), 2)),
      },
    },
    "none",
    false
  );

  rpuChart.windowStartX = rpuAxis.min;
  rpuChart.windowEndX = rpuAxis.max;
  rpuChart.globalMaxX = rpuAxis.max;
  rpuChart._refreshChart();
  rpuChart._appEditSnapshot = null;
  rpuChart._appEditCommitted = false;
  rpuChart.chart.getZr().on("mousedown", () => {
    if (state.syncingRevUnitCharts) return;
    rpuChart._appEditSnapshot = snapshotState();
    rpuChart._appEditLineSnapshot = clone(rpuChart.lines || []);
    rpuChart._appEditCommitted = false;
  });
  rpuChart.onDataChanged = () => handleRevUnitRpuChartChange(rpuChart);
  state.revUnitCharts.rpu = rpuChart;
}

function syncRevUnitCharts() {
  state.syncingRevUnitCharts = true;
  const unitsChart = state.revUnitCharts.units;
  if (unitsChart) {
    unitsChart.lines = [{
      name: displayLabel("New Units"),
      color: TOP_DOWN_COLOR,
      editDomain: {
        moveX: [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, REV_UNIT_YEARS[REV_UNIT_YEARS.length - 1]]],
      },
      data: controlledRevUnitPairs("newUnits", revUnitNewUnitPairs(activeScenario())),
    }];
    unitsChart._refreshChart();
  }
  const rpuChart = state.revUnitCharts.rpu;
  if (rpuChart) {
    if (state.revUnitEditorMode === "allCohorts") {
      rpuChart.setEditMode(false);
      rpuChart.chart.setOption(revUnitAllCohortsChartOption(), true);
      state.syncingRevUnitCharts = false;
      return;
    }
    rpuChart.setEditMode(true);
    const rpuAxis = revUnitRpuAxisConfig();
    rpuChart.baseOption.xAxis.min = rpuAxis.min;
    rpuChart.baseOption.xAxis.max = rpuAxis.max;
    rpuChart.baseOption.xAxis.axisLabel = { formatter: rpuAxis.formatter };
    rpuChart.windowStartX = rpuAxis.min;
    rpuChart.windowEndX = rpuAxis.max;
    rpuChart.globalMaxX = rpuAxis.max;
    rpuChart.lines = revUnitRpuLines();
    rpuChart._refreshChart();
  }
  state.syncingRevUnitCharts = false;
}

function handleRevUnitRpuChartChange(chart) {
  if (state.syncingRevUnitCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const editableLine = chart.lines.find(line => line && line.editable !== false) || chart.lines[0];
  const editableLineIndex = chart.lines.indexOf(editableLine);
  const previousEditableLine = Array.isArray(chart._appEditLineSnapshot)
    ? chart._appEditLineSnapshot[editableLineIndex]
    : null;
  const changedXs = changedNapkinXs(previousEditableLine?.data || [], editableLine.data);
  const key = revUnitRpuControlKey();
  rememberRevUnitControlPoints(key, editableLine.data);

  const cohortYear = state.revUnitEditorCohortYear;
  REV_UNIT_YEARS.filter(year => year >= cohortYear && year > REV_UNIT_LAST_HISTORICAL_YEAR).forEach(year => {
    const value = interpolateNapkinLineValue(editableLine.data, year);
    if (value !== null) setRevUnitRpu(year, cohortYear, value);
  });

  saveScenarios();
  renderRevUnitPlan();
}

function renderRevUnitPlan() {
  renderRevUnitEditorMode();
  renderRevUnitEditorControl();
  renderRevUnitOutputCharts();
  renderRevUnitTable();
  syncRevUnitCharts();
}

function renderRevUnitOutputCharts() {
  const scenario = activeScenario();
  const hhpOutputs = calculateOutputs(scenario);
  if (state.outputCharts.revUnitRevenueComparison) {
    state.outputCharts.revUnitRevenueComparison.setOption({
      animation: false,
      tooltip: {
        trigger: "axis",
        formatter: params => [
          tooltipHeader(params[0]?.axisValue),
          ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(item.value), 0)}`),
        ].join("<br/>"),
      },
      legend: { top: 0 },
      grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
      xAxis: { type: "category", data: YEARS.map(String) },
      yAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
      series: [
        { name: "Plan Revenue", type: "line", data: planRevenue, symbolSize: 6, lineStyle: { color: "#111827", width: 2 }, itemStyle: { color: "#111827" } },
        { name: displayLabel("HHP Model"), type: "line", data: hhpOutputs.revenue, symbolSize: 6, lineStyle: { color: TOP_DOWN_COLOR, width: 2 }, itemStyle: { color: TOP_DOWN_COLOR } },
        { name: displayLabel("Rev / Unit Model"), type: "line", data: YEARS.map(year => revUnitTotalRevenueValue(scenario, year) || 0), symbolSize: 6, lineStyle: { color: BOTTOM_UP_COLOR, width: 2 }, itemStyle: { color: BOTTOM_UP_COLOR } },
      ],
    }, true);
  }

  if (state.outputCharts.revUnitAggregateRpu) {
    state.outputCharts.revUnitAggregateRpu.setOption({
      animation: false,
      tooltip: {
        trigger: "axis",
        formatter: params => [
          tooltipHeader(params[0]?.axisValue),
          ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(item.value), 2)}`),
        ].join("<br/>"),
      },
      grid: { left: 12, right: 18, top: 18, bottom: 36, containLabel: true },
      xAxis: { type: "category", data: REV_UNIT_YEARS.map(String) },
      yAxis: { type: "value", axisLabel: { formatter: value => formatCurrency(Number(value), 0) } },
      series: [{
        name: displayLabel("Aggregate Rev / Cumulative Unit"),
        type: "line",
        data: REV_UNIT_YEARS.map(year => revUnitAggregateRpuValue(scenario, year) || 0),
        areaStyle: { color: "rgba(79, 127, 184, 0.12)" },
        symbolSize: 6,
        lineStyle: { color: BOTTOM_UP_COLOR, width: 2 },
        itemStyle: { color: BOTTOM_UP_COLOR },
      }],
    }, true);
  }
}

function renderRevUnitEditorMode() {
  document.getElementById("rev-unit-editor-cohort")?.classList.toggle("active", state.revUnitEditorMode === "cohort");
  document.getElementById("rev-unit-editor-all-cohorts")?.classList.toggle("active", state.revUnitEditorMode === "allCohorts");
}

function renderRevUnitEditorControl() {
  const slider = document.getElementById("rev-unit-editor-slider");
  const label = document.getElementById("rev-unit-editor-slider-label");
  const value = document.getElementById("rev-unit-editor-slider-value");
  const referenceControl = document.getElementById("rev-unit-reference-control");
  const referenceSlider = document.getElementById("rev-unit-reference-cohort-slider");
  const referenceValue = document.getElementById("rev-unit-reference-cohort-value");
  const referenceShiftControl = document.getElementById("rev-unit-reference-shift-control");
  const referenceShift = document.getElementById("rev-unit-reference-shift");
  if (referenceControl) referenceControl.style.display = state.revUnitEditorMode === "cohort" ? "" : "none";
  if (referenceShiftControl) referenceShiftControl.style.display = state.revUnitEditorMode === "cohort" ? "" : "none";
  if (referenceShift) referenceShift.checked = state.revUnitReferenceShifted;
  if (referenceSlider && referenceValue) {
    referenceSlider.max = String(REV_UNIT_YEARS.length - 1);
    referenceSlider.value = String(Math.max(0, REV_UNIT_YEARS.indexOf(state.revUnitReferenceCohortYear)));
    referenceValue.value = String(state.revUnitReferenceCohortYear);
  }
  if (!slider || !label || !value) return;
  const control = slider.closest(".slider-control");
  if (control) control.style.display = state.revUnitEditorMode === "allCohorts" ? "none" : "";
  if (state.revUnitEditorMode === "allCohorts") return;
  const items = revUnitEditorItems();
  const selected = selectedRevUnitEditorValue();
  slider.max = String(Math.max(0, items.length - 1));
  slider.value = String(Math.max(0, items.indexOf(selected)));
  label.textContent = "Property Cohort";
  value.value = String(selected);
}

function renderRevUnitTable() {
  const table = document.getElementById("rev-unit-table");
  if (!table) return;
  const scenario = activeScenario();
  table.innerHTML = `
    <thead>
      <tr>
        <th class="unit-section" rowspan="2">YEAR</th>
        <th class="unit-section" rowspan="2">NEW UNITS</th>
        <th class="unit-section" rowspan="2">CUMULATIVE_NEW_UNITS</th>
        <th class="unit-section" rowspan="2">TOTAL_REVENUE</th>
        <th class="section-divider" colspan="${REV_UNIT_YEARS.length}">REV_PER_UNIT</th>
        <th class="section-divider unaffiliated-section" rowspan="2">UNAFFILIATED_REVENUE</th>
      </tr>
      <tr>
        ${REV_UNIT_YEARS.map((year, index) => `<th class="${index === 0 ? "section-divider" : ""}">${year}</th>`).join("")}
      </tr>
    </thead>
    <tbody>
      ${REV_UNIT_YEARS.map(year => {
        const revenue = revUnitTotalRevenueValue(scenario, year);
        return `
          <tr>
            <td>${year}</td>
            ${renderRevUnitNewUnitsCell(scenario, year)}
            <td class="output unit-output">${formatValue(revUnitCumulativeNewUnits(scenario, year), "integer")}</td>
            <td class="output revenue-output">${revenue === null ? "-" : formatCurrency(revenue, 0)}</td>
            ${REV_UNIT_YEARS.map((cohortYear, index) => {
              const value = revUnitRpuValue(scenario, year, cohortYear);
              const classes = [
                index === 0 ? "section-divider" : "",
                year === cohortYear ? "input" : "output",
              ].filter(Boolean).join(" ");
              return `<td class="${classes}">${value === null ? "-" : formatCurrency(value, 2)}</td>`;
            }).join("")}
            <td class="section-divider output unaffiliated-output">${formatOptionalCurrency(unaffiliatedRevenueValue(scenario, year))}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  table.querySelectorAll("input[data-rev-unit-year]").forEach(input => {
    input.addEventListener("change", event => {
      const year = Number(event.target.dataset.revUnitYear);
      const parsed = parseInput(event.target.value, "integer");
      if (parsed === null) {
        event.target.value = formatValue(revUnitNewUnitsValue(activeScenario(), year), "integer");
        return;
      }
      const value = Math.max(0, Math.round(parsed));
      if (value === revUnitNewUnitsValue(activeScenario(), year)) return;
      pushUndoSnapshot();
      setRevUnitNewUnits(year, value);
      saveScenarios();
      renderRevUnitPlan();
    });
  });
}

function renderRevUnitNewUnitsCell(scenario, year) {
  const value = revUnitNewUnitsValue(scenario, year);
  if (year > REV_UNIT_LAST_HISTORICAL_YEAR) {
    return `
      <td class="input unit-input">
        <input data-rev-unit-year="${year}" value="${formatValue(value, "integer")}" aria-label="New units ${year}" />
      </td>
    `;
  }
  return `<td class="historical unit-input">${formatValue(value, "integer")}</td>`;
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
  if (!["perCohort", "allCohorts"].includes(state.cohortEditorMode)) {
    state.cohortEditorMode = "perCohort";
  }
  document.getElementById("new-customers-editor-per-cohort")?.classList.toggle("active", state.cohortEditorMode === "perCohort");
  document.getElementById("new-customers-editor-all-cohorts")?.classList.toggle("active", state.cohortEditorMode === "allCohorts");
  document.getElementById("new-customers-per-cohort-editor")?.classList.toggle("active", state.cohortEditorMode === "perCohort");
  document.getElementById("new-customers-per-year-editor")?.classList.toggle("active", state.cohortEditorMode === "perYear");
  document.getElementById("new-customers-new-properties-editor")?.classList.toggle("active", state.cohortEditorMode === "newProperties");
  document.getElementById("new-customers-all-cohorts-editor")?.classList.toggle("active", state.cohortEditorMode === "allCohorts");
}

function renderNewCustomerDrilldownCharts() {
  const comparison = compareScenario();
  const scenarioName = displayScenarioName(activeScenario());
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const comparisonTotals = comparison ? bottomUpNewCustomers(comparison) : null;
  const activeTotals = visibleBottomUpNewCustomers();
  const activeTopDown = activeScenario().drivers.newCustomers;
  const activeDeltaToTopDown = activeTotals.map((value, index) => {
    const delta = value - activeTopDown[index];
    return Math.abs(delta) < 0.000001 ? 0 : delta;
  });
  const comparisonDeltaToTopDown = comparison && comparisonTotals
    ? comparisonTotals.map((value, index) => value - comparison.drivers.newCustomers[index])
    : null;
  state.outputCharts.newCustomerDrilldown.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const year = Number(params[0]?.axisValue);
        const index = YEARS.indexOf(year);
        const topDownMarker = `<span style="display:inline-block;margin-right:4px;border-radius:50%;width:10px;height:10px;background:${TOP_DOWN_COLOR};"></span>`;
        const bottomUpMarker = `<span style="display:inline-block;margin-right:4px;border-radius:50%;width:10px;height:10px;background:${BOTTOM_UP_COLOR};"></span>`;
        return [
          tooltipHeader(year),
          `${topDownMarker} Top-Down: ${formatIntegerMetric(activeTopDown[index] || 0)}`,
          `${bottomUpMarker} Bottom-Up: ${formatIntegerMetric(activeTotals[index] || 0)}`,
          ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(item.value))}`),
        ].join("<br/>");
      },
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: [
      {
        name: `${scenarioName} Delta`,
        type: "bar",
        data: activeDeltaToTopDown,
        itemStyle: { color: value => value.value < 0 ? "#b42318" : "#4f7f52" },
      },
      ...(comparisonDeltaToTopDown ? [{
        name: `${comparisonName} Delta`,
        type: "line",
        data: comparisonDeltaToTopDown,
        symbolSize: 5,
        lineStyle: { color: "#98a2b3", width: 2 },
        itemStyle: { color: "#98a2b3" },
      }] : []),
    ],
  }, true);

  const totalSeries = [
    { name: `${scenarioName} Top-Down`, type: "line", data: activeTopDown, symbolSize: 5, lineStyle: { color: TOP_DOWN_COLOR, width: 2 }, itemStyle: { color: TOP_DOWN_COLOR } },
    { name: `${scenarioName} Bottom-Up`, type: "line", data: activeTotals, symbolSize: 5, z: 3, lineStyle: { color: BOTTOM_UP_COLOR, width: 2 }, itemStyle: { color: BOTTOM_UP_COLOR } },
  ];
  if (comparisonTotals) {
    totalSeries.unshift({ name: `${comparisonName} Bottom-Up`, type: "line", data: comparisonTotals, symbolSize: 5, z: 2, lineStyle: { color: "#98a2b3", width: 3 }, itemStyle: { color: "#98a2b3" } });
  }
  state.outputCharts.newCustomerTotal.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(item.value))}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 58, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: totalSeries,
  }, true);

}

function renderNewCustomerUnitDrilldown() {
  const scenario = activeScenario();
  const comparison = compareScenario();
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const topDown = YEARS.map(year => newPropertyCohortValue(scenario.newCustomerDrilldown.counts, year));
  const bottomUp = YEARS.map(year => visibleNewCustomerUnitBottomUpValue(year));
  const delta = bottomUp.map((value, index) => Math.abs(value - topDown[index]) < 0.000001 ? 0 : value - topDown[index]);
  const comparisonBottomUp = comparison ? YEARS.map(year => newCustomerUnitBottomUpValue(comparison, year)) : null;
  const comparisonDelta = comparison
    ? comparisonBottomUp.map((value, index) => value - newPropertyCohortValue(comparison.newCustomerDrilldown.counts, YEARS[index]))
    : null;

  state.outputCharts.newCustomerUnitBridge?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(item.value))}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 58, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: [
      { name: displayLabel("New Properties Top-Down"), type: "line", data: topDown, symbolSize: 5, lineStyle: { color: TOP_DOWN_COLOR, width: 2 }, itemStyle: { color: TOP_DOWN_COLOR } },
      { name: displayLabel("Unit Build"), type: "line", data: bottomUp, symbolSize: 5, lineStyle: { color: BOTTOM_UP_COLOR, width: 2 }, itemStyle: { color: BOTTOM_UP_COLOR } },
      ...(comparisonBottomUp ? [{
        name: displayLabel(`${comparisonName} Unit Build`),
        type: "line",
        data: comparisonBottomUp,
        symbolSize: 5,
        lineStyle: { color: "#98a2b3", width: 2 },
        itemStyle: { color: "#98a2b3" },
      }] : []),
    ],
  }, true);

  state.outputCharts.newCustomerUnitDelta?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(item.value))}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: [
      {
        name: displayLabel("Unit Build Delta"),
        type: "bar",
        data: delta,
        itemStyle: { color: value => value.value < 0 ? "#b42318" : "#4f7f52" },
      },
      ...(comparisonDelta ? [{
        name: `${comparisonName} Delta`,
        type: "line",
        data: comparisonDelta,
        symbolSize: 5,
        lineStyle: { color: "#98a2b3", width: 2 },
        itemStyle: { color: "#98a2b3" },
      }] : []),
    ],
  }, true);

  renderNewCustomerNewUnitsDrilldown();
}

function renderNewCustomerNewUnitsDrilldown() {
  const scenario = activeScenario();
  const comparison = compareScenario();
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const topDown = YEARS.map(year => newCustomerUnitNewUnitsValue(scenario, year));
  const bottomUp = YEARS.map(year => visibleNewCustomerNewUnitsBottomUpValue(year));
  const delta = bottomUp.map((value, index) => Math.abs(value - topDown[index]) < 0.000001 ? 0 : value - topDown[index]);
  const comparisonTopDown = comparison ? YEARS.map(year => newCustomerUnitNewUnitsValue(comparison, year)) : null;
  const comparisonBottomUp = comparison ? YEARS.map(year => newCustomerNewUnitsBottomUpValue(comparison, year)) : null;
  const comparisonDelta = comparison && comparisonBottomUp
    ? comparisonBottomUp.map((value, index) => value - comparisonTopDown[index])
    : null;

  state.outputCharts.newCustomerNewUnitsBridge?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(item.value))}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 58, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: [
      { name: displayLabel("New Units Top-Down"), type: "line", data: topDown, symbolSize: 5, lineStyle: { color: TOP_DOWN_COLOR, width: 2 }, itemStyle: { color: TOP_DOWN_COLOR } },
      { name: displayLabel("Property Build"), type: "line", data: bottomUp, symbolSize: 5, lineStyle: { color: BOTTOM_UP_COLOR, width: 2 }, itemStyle: { color: BOTTOM_UP_COLOR } },
      ...(comparisonBottomUp ? [{
        name: displayLabel(`${comparisonName} Property Build`),
        type: "line",
        data: comparisonBottomUp,
        symbolSize: 5,
        lineStyle: { color: "#98a2b3", width: 2 },
        itemStyle: { color: "#98a2b3" },
      }] : []),
    ],
  }, true);

  state.outputCharts.newCustomerNewUnitsDelta?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(item.value))}`),
      ].join("<br/>"),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: [
      {
        name: displayLabel("Property Build Delta"),
        type: "bar",
        data: delta,
        itemStyle: { color: value => value.value < 0 ? "#b42318" : "#4f7f52" },
      },
      ...(comparisonDelta ? [{
        name: `${comparisonName} Delta`,
        type: "line",
        data: comparisonDelta,
        symbolSize: 5,
        lineStyle: { color: "#98a2b3", width: 2 },
        itemStyle: { color: "#98a2b3" },
      }] : []),
    ],
  }, true);
}

function renderNewCustomerAllCohortsChart() {
  const chart = state.outputCharts.newCustomerAllCohorts;
  if (!chart) return;
  const counts = activeScenario().newCustomerDrilldown.counts;
  const chartType = state.newCustomerAllCohortsChartType;
  const isArea = chartType === "area";
  const isBar = chartType === "bar";
  const highlightLookup = newCustomerAllCohortsHighlightLookup();
  const cohortNames = NEW_CUSTOMER_COHORT_YEARS.map(cohortLabel);
  const baseSeries = NEW_CUSTOMER_COHORT_YEARS.map(cohortYear => {
    const color = newCustomerCohortColor(cohortYear);
    const highlightedYears = highlightLookup.get(Number(cohortYear)) || new Set();
    return {
      name: cohortLabel(cohortYear),
      type: isBar ? "bar" : "line",
      data: YEARS.map(year => {
        if (!newCustomerCohortApplies(cohortYear, year)) return null;
        const value = newCustomerCohortValue(counts, year, cohortYear);
        if (!isBar || !highlightedYears.has(year)) return value;
        return {
          value,
          itemStyle: {
            color,
            borderColor: "#111827",
            borderWidth: 3,
            shadowColor: "rgba(17, 24, 39, 0.28)",
            shadowBlur: 8,
          },
        };
      }),
      connectNulls: false,
      symbolSize: 5,
      lineStyle: { color, width: isArea ? 1.8 : 2 },
      itemStyle: { color },
      ...(isBar ? { stack: "newCustomersByCohort", barMaxWidth: 48 } : {}),
      ...(isArea ? { stack: "newCustomersByCohort", areaStyle: { color, opacity: 0.58 } } : {}),
      emphasis: {
        focus: "series",
        blurScope: "coordinateSystem",
        lineStyle: { width: 5 },
        itemStyle: { borderColor: "#111827", borderWidth: 2 },
        ...(isArea ? { areaStyle: { opacity: 0.78 } } : {}),
      },
      blur: {
        lineStyle: { opacity: 0.14 },
        itemStyle: { opacity: 0.14 },
        ...(isArea ? { areaStyle: { opacity: 0.12 } } : {}),
      },
    };
  });
  document.getElementById("new-customers-all-cohorts-line")?.classList.toggle("active", chartType === "line");
  document.getElementById("new-customers-all-cohorts-area")?.classList.toggle("active", isArea);
  document.getElementById("new-customers-all-cohorts-bar")?.classList.toggle("active", isBar);
  chart.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params
          .filter(item => item.value !== null && item.value !== undefined)
          .map(item => {
            const value = typeof item.value === "object" && item.value !== null ? item.value.value : item.value;
            return `${item.marker} ${item.seriesName}: ${formatIntegerMetric(Number(value))}`;
          }),
      ].join("<br/>"),
    },
    legend: { type: "scroll", top: 0, data: cohortNames },
    grid: { left: 12, right: 18, top: 48, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactNumber } },
    series: isBar ? baseSeries : [...baseSeries, ...newCustomerAllCohortsHighlightSeries(counts, { stacked: isArea })],
  }, true);
}

function newCustomerAllCohortsHighlightLookup() {
  const lookup = new Map();
  (state.newCustomerAllCohortsHighlights || []).forEach(highlight => {
    const cohortYear = Number(highlight.cohortYear);
    if (!Number.isFinite(cohortYear)) return;
    if (!lookup.has(cohortYear)) lookup.set(cohortYear, new Set());
    const years = lookup.get(cohortYear);
    (highlight.years || []).forEach(year => {
      const numericYear = Number(year);
      if (Number.isFinite(numericYear)) years.add(numericYear);
    });
  });
  return lookup;
}

function newCustomerAllCohortsHighlightSeries(counts, { stacked = false } = {}) {
  return (state.newCustomerAllCohortsHighlights || []).map((highlight, index) => {
    const cohortYear = Number(highlight.cohortYear);
    const years = new Set((highlight.years || []).map(Number));
    const color = newCustomerCohortColor(cohortYear);
    return {
      name: `Affected ${index + 1}`,
      type: "line",
      data: YEARS.map(year => {
        if (!years.has(year) || !newCustomerCohortApplies(cohortYear, year)) return null;
        return stacked
          ? newCustomerCohortStackValue(counts, year, cohortYear)
          : newCustomerCohortValue(counts, year, cohortYear);
      }),
      connectNulls: false,
      silent: true,
      z: 20,
      showSymbol: true,
      symbolSize: 8,
      lineStyle: { color, width: 6, opacity: 0.95 },
      itemStyle: { color, borderColor: "#111827", borderWidth: 1 },
      tooltip: { show: false },
    };
  });
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
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
    name: displayLabel("New Properties"),
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
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
    name: displayLabel("New Properties YoY"),
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
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

function bridgeExistingChartLines() {
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("bridgeExistingProperties");
  const defaultPairs = YEARS.map(year => [year, existingPropertyNewCustomersTotal(activeCounts, year)]);
  const lines = [{
    name: displayLabel("Existing Properties"),
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: YEARS.map(year => [year, existingPropertyNewCustomersTotal(comparison.newCustomerDrilldown.counts, year)]),
    });
  }
  return lines;
}

function bridgeNewChartLines() {
  const activeCounts = activeScenario().newCustomerDrilldown.counts;
  const controlKey = napkinControlKey("bridgeNewProperties");
  const defaultPairs = YEARS.map(year => [year, newPropertyCohortValue(activeCounts, year)]);
  const lines = [{
    name: displayLabel("New Properties"),
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
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: YEARS.map(year => [year, newPropertyCohortValue(comparison.newCustomerDrilldown.counts, year)]),
    });
  }
  return lines;
}

function newCustomerUnitNewUnitsChartLines() {
  const lines = [{
    name: displayLabel("New Units"),
    color: TOP_DOWN_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNewCustomerUnitPairs("newUnits", newCustomerUnitNewUnitsPairs(activeScenario())),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: newCustomerUnitNewUnitsPairs(comparison),
    });
  }
  return lines;
}

function newCustomerUnitCustomersPerUnitChartLines() {
  const lines = [{
    name: displayLabel("New Customers / Unit"),
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNewCustomerUnitPairs("customersPerUnit", newCustomerUnitCustomersPerUnitPairs(activeScenario())),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: newCustomerUnitCustomersPerUnitPairs(comparison),
    });
  }
  return lines;
}

function newCustomerNewUnitsPropertiesChartLines() {
  const lines = [{
    name: displayLabel("New Properties"),
    color: TOP_DOWN_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNewCustomerNewUnitsPairs("newProperties", newCustomerNewUnitsNewPropertiesPairs(activeScenario())),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: newCustomerNewUnitsNewPropertiesPairs(comparison),
    });
  }
  return lines;
}

function newCustomerNewUnitsUnitsPerPropertyChartLines() {
  const lines = [{
    name: displayLabel("Units / Property"),
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: controlledNewCustomerNewUnitsPairs("unitsPerNewProperty", newCustomerNewUnitsUnitsPerPropertyPairs(activeScenario())),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: newCustomerNewUnitsUnitsPerPropertyPairs(comparison),
    });
  }
  return lines;
}

function setNewPropertyTopDownToUnitBuild() {
  pushUndoSnapshot();
  const scenario = activeScenario();
  const years = YEARS.filter(editableYear);
  years.forEach(year => {
    setNewCustomerCohortValuePreservingFutureYoy(
      scenario.newCustomerDrilldown.counts,
      year,
      year,
      visibleNewCustomerUnitBottomUpValue(year)
    );
  });
  ensureControlPointsForYears(
    scenario.newCustomerDrilldown.controlPoints,
    napkinControlKey("bridgeNewProperties"),
    years
  );
  ensureControlPointsForYears(
    scenario.newCustomerDrilldown.controlPoints,
    napkinControlKey("newPropertiesCount"),
    years
  );
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange();
}

function setTotalNewCustomersTopDownToBottomUp() {
  pushUndoSnapshot();
  activeScenario().newCustomerSource = "topDown";
  YEARS.forEach((year, index) => {
    if (editableYear(year)) {
      activeScenario().drivers.newCustomers[index] = visibleBottomUpNewCustomersValue(year);
    }
  });
  saveScenarios();
  syncDriverCharts();
  syncRevUnitCharts();
  renderAfterNewCustomerChartChange();
}

function setNewUnitsTopDownToPropertyBuild() {
  pushUndoSnapshot();
  const years = YEARS.filter(editableYear);
  years.forEach(year => {
    setNewCustomerUnitNewUnits(year, visibleNewCustomerNewUnitsBottomUpValue(year), { exact: true });
  });
  ensureControlPointsForYears(newCustomerUnitControlPoints(), "newUnits", years);
  saveScenarios();
  renderAfterNewCustomerChartChange();
}

function matchNewCustomerUnitNewUnitsToTopDown() {
  pushUndoSnapshot();
  const counts = activeScenario().newCustomerDrilldown.counts;
  const years = YEARS.filter(editableYear);
  years.forEach(year => {
    const targetNewCustomers = newPropertyCohortValue(counts, year);
    const customersPerUnit = visibleNewCustomerUnitValue(
      "customersPerUnit",
      newCustomerUnitCustomersPerUnitPairs(activeScenario()),
      year,
      newCustomerUnitCustomersPerUnitValue(activeScenario(), year)
    );
    setNewCustomerUnitNewUnits(year, customersPerUnit > 0 ? targetNewCustomers / customersPerUnit : 0, { exact: true });
  });
  ensureControlPointsForYears(newCustomerUnitControlPoints(), "newUnits", years);
  saveScenarios();
  renderAfterNewCustomerChartChange();
}

function matchNewCustomerUnitCustomersPerUnitToTopDown() {
  pushUndoSnapshot();
  const counts = activeScenario().newCustomerDrilldown.counts;
  const years = YEARS.filter(editableYear);
  years.forEach(year => {
    const targetNewCustomers = newPropertyCohortValue(counts, year);
    const newUnits = visibleNewCustomerUnitValue(
      "newUnits",
      newCustomerUnitNewUnitsPairs(activeScenario()),
      year,
      newCustomerUnitNewUnitsValue(activeScenario(), year)
    );
    setNewCustomerUnitCustomersPerUnit(year, newUnits > 0 ? targetNewCustomers / newUnits : 0, { exact: true });
  });
  ensureControlPointsForYears(newCustomerUnitControlPoints(), "customersPerUnit", years);
  saveScenarios();
  renderAfterNewCustomerChartChange();
}

function matchNewCustomerNewUnitsPropertiesToTopDown() {
  pushUndoSnapshot();
  const years = YEARS.filter(editableYear);
  years.forEach(year => {
    const targetNewUnits = visibleNewCustomerUnitValue(
      "newUnits",
      newCustomerUnitNewUnitsPairs(activeScenario()),
      year,
      newCustomerUnitNewUnitsValue(activeScenario(), year)
    );
    const unitsPerProperty = visibleNewCustomerNewUnitsValue(
      "unitsPerNewProperty",
      newCustomerNewUnitsUnitsPerPropertyPairs(activeScenario()),
      year,
      newCustomerNewUnitsUnitsPerPropertyValue(activeScenario(), year)
    );
    setNewCustomerNewUnitsNewProperties(year, unitsPerProperty > 0 ? targetNewUnits / unitsPerProperty : 0, { exact: true });
  });
  ensureControlPointsForYears(newCustomerNewUnitsControlPoints(), "newProperties", years);
  saveScenarios();
  renderAfterNewCustomerChartChange();
}

function matchNewCustomerNewUnitsUnitsPerPropertyToTopDown() {
  pushUndoSnapshot();
  const years = YEARS.filter(editableYear);
  years.forEach(year => {
    const targetNewUnits = visibleNewCustomerUnitValue(
      "newUnits",
      newCustomerUnitNewUnitsPairs(activeScenario()),
      year,
      newCustomerUnitNewUnitsValue(activeScenario(), year)
    );
    const newProperties = visibleNewCustomerNewUnitsValue(
      "newProperties",
      newCustomerNewUnitsNewPropertiesPairs(activeScenario()),
      year,
      newCustomerNewUnitsNewPropertiesValue(activeScenario(), year)
    );
    setNewCustomerNewUnitsUnitsPerProperty(year, newProperties > 0 ? targetNewUnits / newProperties : 0, { exact: true });
  });
  ensureControlPointsForYears(newCustomerNewUnitsControlPoints(), "unitsPerNewProperty", years);
  saveScenarios();
  renderAfterNewCustomerChartChange();
}

function syncNewCustomerCohortEditors({ excludeChart = null } = {}) {
  state.syncingCohortCharts = true;
  if (state.cohortCharts.count && state.cohortCharts.count !== excludeChart) {
    state.cohortCharts.count.lines = safeNapkinLines(cohortCountChartLines());
    state.cohortCharts.count._refreshChart();
    styleComparisonSeries(state.cohortCharts.count);
  }
  if (state.cohortCharts.yoy && state.cohortCharts.yoy !== excludeChart) {
    state.cohortCharts.yoy.lines = safeNapkinLines(cohortYoyChartLines());
    state.cohortCharts.yoy._refreshChart();
    styleComparisonSeries(state.cohortCharts.yoy);
  }
  if (state.cohortCharts.yearExisting && state.cohortCharts.yearExisting !== excludeChart) {
    const year = selectedNewCustomerYearEditorYear();
    const cohorts = existingPropertyCohortsForYear(year);
    syncYearEditorChart(state.cohortCharts.yearExisting, cohorts, yearExistingChartLines());
  }
  if (state.cohortCharts.yearExistingYoy && state.cohortCharts.yearExistingYoy !== excludeChart) {
    const year = selectedNewCustomerYearEditorYear();
    const cohorts = existingPropertyCohortsForYear(year);
    syncYearEditorChart(state.cohortCharts.yearExistingYoy, cohorts, yearExistingYoyChartLines());
  }
  if (state.cohortCharts.newPropertiesCount && state.cohortCharts.newPropertiesCount !== excludeChart) {
    state.cohortCharts.newPropertiesCount.lines = safeNapkinLines(newPropertiesCountChartLines());
    state.cohortCharts.newPropertiesCount._refreshChart();
    styleComparisonSeries(state.cohortCharts.newPropertiesCount);
  }
  if (state.cohortCharts.newPropertiesYoy && state.cohortCharts.newPropertiesYoy !== excludeChart) {
    state.cohortCharts.newPropertiesYoy.lines = safeNapkinLines(newPropertiesYoyChartLines());
    state.cohortCharts.newPropertiesYoy._refreshChart();
    styleComparisonSeries(state.cohortCharts.newPropertiesYoy);
  }
  if (state.cohortCharts.bridgeExisting && state.cohortCharts.bridgeExisting !== excludeChart) {
    state.cohortCharts.bridgeExisting.lines = safeNapkinLines(bridgeExistingChartLines());
    state.cohortCharts.bridgeExisting._refreshChart();
    styleComparisonSeries(state.cohortCharts.bridgeExisting);
  }
  if (state.cohortCharts.bridgeNew && state.cohortCharts.bridgeNew !== excludeChart) {
    state.cohortCharts.bridgeNew.lines = safeNapkinLines(bridgeNewChartLines());
    state.cohortCharts.bridgeNew._refreshChart();
    styleComparisonSeries(state.cohortCharts.bridgeNew);
  }
  if (state.cohortCharts.unitNewUnits && state.cohortCharts.unitNewUnits !== excludeChart) {
    state.cohortCharts.unitNewUnits.lines = safeNapkinLines(newCustomerUnitNewUnitsChartLines());
    state.cohortCharts.unitNewUnits._refreshChart();
    styleComparisonSeries(state.cohortCharts.unitNewUnits);
  }
  if (state.cohortCharts.unitCustomersPerUnit && state.cohortCharts.unitCustomersPerUnit !== excludeChart) {
    state.cohortCharts.unitCustomersPerUnit.lines = safeNapkinLines(newCustomerUnitCustomersPerUnitChartLines());
    state.cohortCharts.unitCustomersPerUnit._refreshChart();
    styleComparisonSeries(state.cohortCharts.unitCustomersPerUnit);
  }
  if (state.cohortCharts.newUnitsProperties && state.cohortCharts.newUnitsProperties !== excludeChart) {
    state.cohortCharts.newUnitsProperties.lines = safeNapkinLines(newCustomerNewUnitsPropertiesChartLines());
    state.cohortCharts.newUnitsProperties._refreshChart();
    styleComparisonSeries(state.cohortCharts.newUnitsProperties);
  }
  if (state.cohortCharts.newUnitsUnitsPerProperty && state.cohortCharts.newUnitsUnitsPerProperty !== excludeChart) {
    state.cohortCharts.newUnitsUnitsPerProperty.lines = safeNapkinLines(newCustomerNewUnitsUnitsPerPropertyChartLines());
    state.cohortCharts.newUnitsUnitsPerProperty._refreshChart();
    styleComparisonSeries(state.cohortCharts.newUnitsUnitsPerProperty);
  }
  state.syncingCohortCharts = false;
}

function safeNapkinLines(lines) {
  return (lines || []).filter(line => Array.isArray(line.data) && line.data.length > 0);
}

function renderAfterNewCustomerChartChange(chart) {
  renderAll({ excludeNewCustomerChart: chart?._isDragging ? chart : null });
}

function syncYearEditorChart(chart, cohorts, lines) {
  chart.baseOption.xAxis.min = 0;
  chart.baseOption.xAxis.max = Math.max(0, cohorts.length - 1);
  chart.windowStartX = 0;
  chart.windowEndX = Math.max(0, cohorts.length - 1);
  chart.globalMaxX = Math.max(0, cohorts.length - 1);
  chart.lines = safeNapkinLines(lines);
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
  changedXs.forEach(year => {
    addNapkinControlPointIfControlled(napkinControlKey("cohortYoy", cohortYear), year);
    addNapkinControlPointIfControlled(napkinControlKey("cohortYoy", cohortYear), year + 1);
  });
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && newCustomerCohortApplies(cohortYear, year) && value !== null) {
      setNewCustomerCohortValue(activeScenario().newCustomerDrilldown.counts, year, cohortYear, Math.max(0, value));
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange(chart);
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
  changedXs.forEach(year => {
    YEARS
      .filter(item => editableYear(item) && item >= year)
      .forEach(item => addNapkinControlPointIfControlled(napkinControlKey("cohortCount", cohortYear), item));
  });
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      applyNewCustomerCohortYoy(activeScenario().newCustomerDrilldown.counts, year, cohortYear, value / 100);
    }
  });
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange(chart);
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
  renderAfterNewCustomerChartChange(chart);
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
  renderAfterNewCustomerChartChange(chart);
}

function handleNewPropertiesCountChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const affectedYears = affectedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("newPropertiesCount"), chart.lines[0].data);
  affectedYears.forEach(year => addNapkinControlPoint(napkinControlKey("newPropertiesYoy"), year));
  affectedYears.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerCohortValuePreservingFutureYoy(activeScenario().newCustomerDrilldown.counts, year, year, Math.max(0, value));
    }
  });
  ensureBridgeExistingPointsForActuals();
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange(chart);
}

function handleNewPropertiesYoyChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const affectedYears = affectedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("newPropertiesYoy"), chart.lines[0].data);
  affectedYears.forEach(year => addNapkinControlPoint(napkinControlKey("newPropertiesCount"), year));
  affectedYears.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    const prior = newPropertyCohortValue(activeScenario().newCustomerDrilldown.counts, year - 1);
    if (editableYear(year) && prior && value !== null) {
      setNewCustomerCohortValuePreservingFutureYoy(activeScenario().newCustomerDrilldown.counts, year, year, prior * (1 + (value / 100)));
    }
  });
  ensureBridgeExistingPointsForActuals();
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange(chart);
}

function handleBridgeExistingChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const affectedYears = affectedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("bridgeExistingProperties"), chart.lines[0].data);
  affectedYears.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setExistingPropertyNewCustomersTotal(activeScenario().newCustomerDrilldown.counts, year, Math.max(0, value));
    }
  });
  ensureBridgeExistingPointsForActuals();
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange(chart);
}

function handleBridgeNewChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  const affectedYears = affectedNapkinXs(chart._appEditLineSnapshot, chart.lines[0].data);
  rememberNapkinControlPoints(napkinControlKey("bridgeNewProperties"), chart.lines[0].data);
  affectedYears.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerCohortValuePreservingFutureYoy(activeScenario().newCustomerDrilldown.counts, year, year, Math.max(0, value));
    }
  });
  ensureBridgeExistingPointsForActuals();
  saveScenarios();
  syncDriverCharts();
  renderAfterNewCustomerChartChange(chart);
}

function handleNewCustomerUnitNewUnitsChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  rememberNewCustomerUnitControlPoints("newUnits", chart.lines[0].data);
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerUnitNewUnits(year, value);
    }
  });
  saveScenarios();
  renderAfterNewCustomerChartChange(chart);
}

function handleNewCustomerUnitCustomersPerUnitChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  rememberNewCustomerUnitControlPoints("customersPerUnit", chart.lines[0].data);
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerUnitCustomersPerUnit(year, value);
    }
  });
  saveScenarios();
  renderAfterNewCustomerChartChange(chart);
}

function handleNewCustomerNewUnitsPropertiesChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  rememberNewCustomerNewUnitsControlPoints("newProperties", chart.lines[0].data);
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerNewUnitsNewProperties(year, value);
    }
  });
  saveScenarios();
  renderAfterNewCustomerChartChange(chart);
}

function handleNewCustomerNewUnitsUnitsPerPropertyChartChange(chart) {
  if (state.syncingCohortCharts) return;
  if (chart._appEditSnapshot && !chart._appEditCommitted) {
    pushUndoSnapshot(chart._appEditSnapshot);
    chart._appEditCommitted = true;
  }
  rememberNewCustomerNewUnitsControlPoints("unitsPerNewProperty", chart.lines[0].data);
  YEARS.forEach(year => {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (editableYear(year) && value !== null) {
      setNewCustomerNewUnitsUnitsPerProperty(year, value);
    }
  });
  saveScenarios();
  renderAfterNewCustomerChartChange(chart);
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

function numberList(value) {
  const values = Array.isArray(value) ? value : [value];
  return values.map(Number).filter(Number.isFinite);
}

function allCohortsHighlightYears(years, cohortYear) {
  const numericYears = numberList(years);
  if (!numericYears.length) return [];
  const uniqueYears = Array.from(new Set(numericYears)).sort((left, right) => left - right);
  const firstYear = uniqueYears[0];
  return YEARS.filter(item => item >= firstYear && newCustomerCohortApplies(cohortYear, item));
}

function setNewCustomerAllCohortsHighlightsFromPayload({ countYear, countYears, countCohortYear, countCohortYears } = {}) {
  const years = numberList(countYears !== undefined ? countYears : countYear);
  const cohortYears = numberList(countCohortYears !== undefined ? countCohortYears : countCohortYear);
  const highlights = cohortYears
    .map(cohortYear => ({ cohortYear, years: allCohortsHighlightYears(years, cohortYear) }))
    .filter(highlight => highlight.years.length);
  const key = JSON.stringify(highlights);
  if (key === state.newCustomerAllCohortsHighlightKey) return;
  state.newCustomerAllCohortsHighlightKey = key;
  state.newCustomerAllCohortsHighlights = highlights;
  renderNewCustomerAllCohortsChart();
}

function clearNewCustomerAllCohortsHighlights() {
  if (!state.newCustomerAllCohortsHighlightKey) return;
  state.newCustomerAllCohortsHighlightKey = "";
  state.newCustomerAllCohortsHighlights = [];
  renderNewCustomerAllCohortsChart();
}

function highlightNewCustomerCells({ countYear, countYears, countCohortYear, countCohortYears, yoyYear, yoyYears, yoyCohortYear, yoyCohortYears } = {}) {
  clearNewCustomerTableCellHighlights();
  setNewCustomerAllCohortsHighlightsFromPayload({ countYear, countYears, countCohortYear, countCohortYears });
  const normalizedCountYears = numberList(countYears !== undefined ? countYears : countYear);
  const normalizedCountCohortYears = numberList(countCohortYears !== undefined ? countCohortYears : countCohortYear);
  normalizedCountYears.forEach(year => {
    normalizedCountCohortYears.forEach(cohortYear => {
      document
        .querySelectorAll(`[data-count-year="${year}"][data-count-cohort-year="${cohortYear}"]`)
        .forEach(cell => cell.classList.add("linked-highlight"));
    });
  });
  const normalizedYoyYears = numberList(yoyYears !== undefined ? yoyYears : yoyYear);
  const normalizedYoyCohortYears = numberList(yoyCohortYears !== undefined ? yoyCohortYears : yoyCohortYear);
  normalizedYoyYears.forEach(year => {
    normalizedYoyCohortYears.forEach(cohortYear => {
      document
        .querySelectorAll(`[data-yoy-year="${year}"][data-yoy-cohort-year="${cohortYear}"]`)
        .forEach(cell => cell.classList.add("linked-highlight"));
    });
  });
}

function clearNewCustomerTableCellHighlights() {
  document
    .querySelectorAll(".linked-highlight")
    .forEach(cell => cell.classList.remove("linked-highlight"));
}

function clearNewCustomerCellHighlights() {
  clearNewCustomerTableCellHighlights();
  clearNewCustomerAllCohortsHighlights();
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
    .map(scenario => `<option value="${scenario.id}">${displayScenarioName(scenario)}</option>`)
    .join("");
  select.value = state.activeScenarioId;
  compareSelect.innerHTML = [
    `<option value="">None</option>`,
    ...scenarios.map(scenario => `<option value="${scenario.id}">${displayScenarioName(scenario)}</option>`),
  ].join("");
  if (state.compareScenarioId && !state.scenarios[state.compareScenarioId]) {
    setCompareScenario("");
  }
  compareSelect.value = state.compareScenarioId;

  const versions = activeScenario().versions || [];
  versionSelect.innerHTML = versions.length
    ? versions.slice().reverse().map((version, index) => `<option value="${version.id}">${formatVersionLabel(version, index)}</option>`).join("")
    : `<option value="">No versions</option>`;
  if (!state.selectedVersionId || !versions.some(version => version.id === state.selectedVersionId)) {
    state.selectedVersionId = versions.length ? versions[versions.length - 1].id : "";
  }
  versionSelect.value = state.selectedVersionId;
  restoreButton.disabled = !state.selectedVersionId;
}

function formatVersionLabel(version, reverseIndex = 0) {
  if (isAnonymizedView()) return `Version ${reverseIndex + 1}`;
  const date = new Date(version.createdAt);
  const timestamp = Number.isNaN(date.getTime())
    ? version.createdAt
    : date.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  return `${timestamp} - ${version.label}`;
}

function isMetricTextNodeValue(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  const normalizedYear = trimmed.replace(/,/g, "");
  if (/^20\d{2}$/.test(normalizedYear)) return false;
  return /^-?\$?[\d,]+(\.\d+)?%?$/.test(trimmed)
    || /^-?\$?[\d,]+(\.\d+)?[kKmMbB]$/.test(trimmed)
    || /^-\$?[\d,]+(\.\d+)?$/.test(trimmed);
}

function anonymizeTextNodeValue(text) {
  const replaced = anonymizeText(text);
  if (!isMetricTextNodeValue(replaced)) return replaced;
  const leading = replaced.match(/^\s*/)?.[0] || "";
  const trailing = replaced.match(/\s*$/)?.[0] || "";
  return `${leading}${anonymizedMetricValue()}${trailing}`;
}

function applyAnonymizedView() {
  const checkbox = document.getElementById("anonymized-view-toggle");
  if (checkbox) checkbox.checked = isAnonymizedView();
  document.body.classList.toggle("anonymized-view", isAnonymizedView());
  document.title = displayLabel("PetScreening $100M Path");

  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent || ["SCRIPT", "STYLE", "NOSCRIPT", "TEXTAREA"].includes(parent.tagName)) {
        return NodeFilter.FILTER_REJECT;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  nodes.forEach(node => {
    if (!originalTextByNode.has(node)) originalTextByNode.set(node, node.nodeValue);
    const original = originalTextByNode.get(node);
    node.nodeValue = isAnonymizedView() ? anonymizeTextNodeValue(original) : original;
  });
}

function renderAll({ syncNewCustomerEditors = true, excludeNewCustomerChart = null } = {}) {
  const outputs = calculateOutputs(activeScenario());
  renderKpis(outputs);
  renderGuidedPlan(outputs);
  renderOutputCharts(outputs);
  renderMetricTree(outputs);
  renderCanvasFocus();
  renderTable(outputs);
  renderNewCustomerDrilldown({ syncEditors: syncNewCustomerEditors, excludeEditor: excludeNewCustomerChart });
  renderRevUnitPlan();
  applyAnonymizedView();
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
    revUnitPlan: clone(activeScenario().revUnitPlan),
  };
  addScenarioVersion(state.scenarios[id], "Save As");
  state.activeScenarioId = id;
  state.selectedVersionId = state.scenarios[id].versions[state.scenarios[id].versions.length - 1].id;
  setCompareScenario(previousActiveId);
  saveScenarios();
  renderScenarioSelect();
  syncDriverCharts();
  syncRevUnitCharts();
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
  scenario.revUnitPlan = version.revUnitPlan
    ? clone(version.revUnitPlan)
    : createDefaultRevUnitPlan();
  enforceNewCustomerBaseline(scenario);
  saveScenarios();
  syncDriverCharts();
  syncRevUnitCharts();
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
    activeScenario().revUnitPlan = createDefaultRevUnitPlan();
  }
  saveScenarios();
  syncDriverCharts();
  syncRevUnitCharts();
  renderAll();
}

function focusElementById(id) {
  if (!id) return;
  setTimeout(() => {
    document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "center" });
  }, 0);
}

function openMetricTreeTarget(action, focusChart) {
  if (action === "newCustomers") {
    state.newCustomerUnitDrilldownOpen = false;
    state.newCustomerNewUnitsDrilldownOpen = false;
    setView("newCustomers");
    focusElementById("new-customers-total-chart");
    return;
  }
  if (action === "unitDrilldown") {
    state.newCustomerUnitDrilldownOpen = true;
    state.newCustomerNewUnitsDrilldownOpen = false;
    setView("newCustomers");
    renderNewCustomerUnitDrilldownToggle();
    focusElementById("new-customers-unit-drilldown");
    return;
  }
  if (action === "newUnitsDrilldown") {
    state.newCustomerUnitDrilldownOpen = true;
    state.newCustomerNewUnitsDrilldownOpen = true;
    setView("newCustomers");
    renderNewCustomerUnitDrilldownToggle();
    focusElementById("new-customers-new-units-drilldown");
    return;
  }
  if (action === "revUnit") {
    setView("revPerUnit");
    focusElementById("rev-unit-revenue-comparison-chart");
    return;
  }
  setView("chart");
  focusElementById(focusChart || "revenue-chart");
}

function bindControls() {
  bindNapkinYAxisControls();
  document.getElementById("anonymized-view-toggle")?.addEventListener("change", event => {
    state.anonymizedView = event.target.checked;
    syncDriverCharts();
    syncRevUnitCharts();
    renderAll();
    renderScenarioSelect();
    applyAnonymizedView();
  });
  document.getElementById("scenario-select").addEventListener("change", event => {
    state.activeScenarioId = event.target.value;
    state.selectedVersionId = "";
    updateHistoryControls();
    renderScenarioSelect();
    syncDriverCharts();
    syncRevUnitCharts();
    renderAll();
  });
  document.getElementById("compare-scenario-select").addEventListener("change", event => {
    setCompareScenario(event.target.value);
    syncDriverCharts();
    syncRevUnitCharts();
    renderAll();
  });
  document.getElementById("new-customers-source")?.addEventListener("change", event => {
    pushUndoSnapshot();
    activeScenario().newCustomerSource = event.target.value;
    saveScenarios();
    syncDriverCharts();
    syncRevUnitCharts();
    renderAll();
  });
  document.getElementById("rev-unit-editor-cohort").addEventListener("click", () => {
    state.revUnitEditorMode = "cohort";
    renderRevUnitPlan();
  });
  document.getElementById("rev-unit-editor-all-cohorts").addEventListener("click", () => {
    state.revUnitEditorMode = "allCohorts";
    renderRevUnitPlan();
  });
  document.getElementById("rev-unit-editor-slider").addEventListener("input", event => {
    const items = revUnitEditorItems();
    setSelectedRevUnitEditorValue(items[Number(event.target.value)] ?? selectedRevUnitEditorValue());
    renderRevUnitEditorControl();
    syncRevUnitCharts();
  });
  document.getElementById("rev-unit-reference-cohort-slider").addEventListener("input", event => {
    state.revUnitReferenceCohortYear = REV_UNIT_YEARS[Number(event.target.value)] ?? state.revUnitReferenceCohortYear;
    renderRevUnitEditorControl();
    syncRevUnitCharts();
  });
  document.getElementById("rev-unit-reference-shift").addEventListener("change", event => {
    state.revUnitReferenceShifted = event.target.checked;
    renderRevUnitEditorControl();
    syncRevUnitCharts();
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
  document.getElementById("new-customers-editor-all-cohorts").addEventListener("click", () => {
    state.cohortEditorMode = "allCohorts";
    renderNewCustomerEditorMode();
    renderNewCustomerAllCohortsChart();
    setTimeout(() => {
      state.outputCharts.newCustomerAllCohorts?.resize();
    }, 0);
  });
  document.getElementById("new-customers-all-cohorts-line").addEventListener("click", () => {
    state.newCustomerAllCohortsChartType = "line";
    renderNewCustomerAllCohortsChart();
  });
  document.getElementById("new-customers-all-cohorts-area").addEventListener("click", () => {
    state.newCustomerAllCohortsChartType = "area";
    renderNewCustomerAllCohortsChart();
  });
  document.getElementById("new-customers-all-cohorts-bar").addEventListener("click", () => {
    state.newCustomerAllCohortsChartType = "bar";
    renderNewCustomerAllCohortsChart();
  });
  document.getElementById("toggle-new-customers-unit-drilldown").addEventListener("click", () => {
    state.newCustomerUnitDrilldownOpen = true;
    state.newCustomerNewUnitsDrilldownOpen = false;
    renderNewCustomerUnitDrilldownToggle();
    setTimeout(() => {
      resizeNewCustomerUnitDrilldownCharts();
      syncNewCustomerCohortEditors();
    }, 0);
  });
  document.getElementById("back-new-customers-unit-drilldown").addEventListener("click", () => {
    state.newCustomerUnitDrilldownOpen = false;
    state.newCustomerNewUnitsDrilldownOpen = false;
    renderNewCustomerUnitDrilldownToggle();
    setTimeout(() => {
      syncNewCustomerCohortEditors();
    }, 0);
  });
  document.getElementById("toggle-new-customers-new-units-drilldown").addEventListener("click", () => {
    state.newCustomerNewUnitsDrilldownOpen = true;
    renderNewCustomerUnitDrilldownToggle();
    setTimeout(() => {
      resizeNewCustomerNewUnitsDrilldownCharts();
      syncNewCustomerCohortEditors();
    }, 0);
  });
  document.getElementById("back-new-customers-new-units-drilldown").addEventListener("click", () => {
    state.newCustomerNewUnitsDrilldownOpen = false;
    renderNewCustomerUnitDrilldownToggle();
    setTimeout(() => {
      resizeNewCustomerUnitDrilldownCharts();
      syncNewCustomerCohortEditors();
    }, 0);
  });
  document.getElementById("match-new-customer-unit-new-units").addEventListener("click", matchNewCustomerUnitNewUnitsToTopDown);
  document.getElementById("match-new-customer-unit-customers-per-unit").addEventListener("click", matchNewCustomerUnitCustomersPerUnitToTopDown);
  document.getElementById("match-new-customer-new-units-properties").addEventListener("click", matchNewCustomerNewUnitsPropertiesToTopDown);
  document.getElementById("match-new-customer-new-units-units-per-property").addEventListener("click", matchNewCustomerNewUnitsUnitsPerPropertyToTopDown);
  document.getElementById("set-new-customers-total-top-down").addEventListener("click", setTotalNewCustomersTopDownToBottomUp);
  document.getElementById("set-new-customer-unit-top-down").addEventListener("click", setNewPropertyTopDownToUnitBuild);
  document.getElementById("set-new-customer-new-units-top-down").addEventListener("click", setNewUnitsTopDownToPropertyBuild);
  document.getElementById("unlock-all-cohorts").addEventListener("click", () => setCohortLockSelection(() => false));
  document.getElementById("lock-all-cohorts").addEventListener("click", () => setCohortLockSelection(() => true));
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

  document.getElementById("tab-tree").addEventListener("click", () => setView("tree"));
  document.getElementById("tab-guided").addEventListener("click", () => setView("guided"));
  document.getElementById("tab-chart").addEventListener("click", () => setView("chart"));
  document.getElementById("tab-table").addEventListener("click", () => setView("table"));
  document.getElementById("tab-rev-per-unit").addEventListener("click", () => setView("revPerUnit"));
  document.getElementById("canvas-zoom-out").addEventListener("click", () => {
    setCanvasZoom(state.canvasZoom - CANVAS_ZOOM_STEP);
  });
  document.getElementById("canvas-zoom-in").addEventListener("click", () => {
    setCanvasZoom(state.canvasZoom + CANVAS_ZOOM_STEP);
  });
  bindCanvasTrackpadZoom();
  document.getElementById("drilldown-new-customers-from-chart").addEventListener("click", () => {
    state.newCustomerUnitDrilldownOpen = false;
    state.newCustomerNewUnitsDrilldownOpen = false;
    setView("newCustomers");
  });
  document.querySelectorAll(".guided-jump").forEach(button => {
    button.addEventListener("click", event => {
      const targetView = event.currentTarget.dataset.targetView;
      const focusChart = event.currentTarget.dataset.focusChart;
      if (targetView) setView(targetView);
      if (focusChart) {
        focusElementById(focusChart);
      }
    });
  });
  document.querySelectorAll(".tree-jump").forEach(button => {
    button.addEventListener("click", event => {
      event.stopPropagation();
      openMetricTreeTarget(event.currentTarget.dataset.treeAction, event.currentTarget.dataset.focusChart);
    });
  });
  document.querySelectorAll(".canvas-node[data-node-id]").forEach(node => {
    node.addEventListener("click", event => {
      if (event.target.closest("button")) return;
      setCanvasFocus(event.currentTarget.dataset.nodeId);
    });
    node.addEventListener("keydown", event => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      setCanvasFocus(event.currentTarget.dataset.nodeId);
    });
  });
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
  state.currentView = view;
  document.getElementById("tab-tree").classList.toggle("active", view === "tree");
  document.getElementById("tab-guided").classList.toggle("active", view === "guided");
  document.getElementById("tab-chart").classList.toggle("active", view === "chart");
  document.getElementById("tab-table").classList.toggle("active", view === "table");
  document.getElementById("tab-rev-per-unit").classList.toggle("active", view === "revPerUnit");
  document.getElementById("tree-view").classList.toggle("active", view === "tree");
  document.getElementById("guided-view").classList.toggle("active", view === "guided");
  document.getElementById("chart-view").classList.toggle("active", view === "chart");
  document.getElementById("table-view").classList.toggle("active", view === "table");
  document.getElementById("new-customers-view").classList.toggle("active", view === "newCustomers");
  document.getElementById("rev-per-unit-view").classList.toggle("active", view === "revPerUnit");
  renderNewCustomerUnitDrilldownToggle();
  if (view === "tree") {
    enterMetricTreeCanvas();
  }
  setTimeout(() => {
    Object.values(state.charts).forEach(chart => chart.resize());
    Object.values(state.cohortCharts).forEach(chart => chart.resize());
    Object.values(state.outputCharts).forEach(chart => chart.resize());
    Object.values(state.revUnitCharts).forEach(chart => chart.resize());
  }, 0);
}

function init() {
  renderScenarioSelect();
  bindControls();
  initOutputCharts();
  Object.keys(driverMeta).forEach(initDriverChart);
  initNewCustomerCohortEditors();
  initRevUnitCharts();
  updateHistoryControls();
  renderAll();
}

init();
