const YEARS = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029];
const FIRST_FORECAST_YEAR = 2026;
const FIRST_COST_EDIT_YEAR = 2025;
const STORAGE_KEY = "ps-hhp-simple-scenarios-v1";
const TOP_DOWN_COLOR = "#6fa76b";
const BOTTOM_UP_COLOR = "#4f7fb8";
const INITIATIVE_TEAMS = [
  "Sales",
  "Customer Success",
  "Customer Operations",
  "Customer Experience",
  "Marketing",
  "Product",
  "Growth",
  "External Partners",
];
const LEAF_DEFENSE_KEYS = [
  "retentionRate",
  "profilesReturningCustomer",
  "profilesNewCustomer",
  "revenueReturningProfile",
  "revenueNewProfile",
  "existingPropertyNewPayingCustomers",
  "growthNewPayingCustomers",
  "growthRevPerNewPayingCustomer",
  "strProperties",
  "strRevPerProperty",
  "otherRevenue",
  "nonLaborCost",
  "laborFte",
  "laborCostPerFte",
  "newCustomersPerUnit",
  "newProperties",
  "newUnitsPerProperty",
  "priorPayingCustomers",
];

const planRevenue = [5546199, 11835538, 19529444, 30709500, 42235617, 66845700, 95697000, 123957600, 180563000];
const growthRevenue = [0, 1385506, 2636508, 5052000, 7000000, null, null, null, null];
const strRevenue = [0, 0, 0, 0, 0, null, null, null, null];
const otherRevenue = [null, null, 1678634, 848400, 1000000, null, null, null, null];
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

function baseGrowthRevenueTopDownValues() {
  let latestValue = 0;
  return YEARS.map((year, index) => {
    const value = growthRevenue[index];
    if (value !== null && value !== undefined) {
      latestValue = Number(value);
      return latestValue;
    }
    return year >= FIRST_FORECAST_YEAR ? latestValue : null;
  });
}

function baseGrowthRevPerNewPayingCustomerValues() {
  const historicalRatios = growthRevenue
    .map((value, index) => {
      const customers = Number(baseDrivers.newCustomers[index] || 0);
      return value !== null && value !== undefined && customers > 0 ? Number(value) / customers : null;
    })
    .filter(value => Number.isFinite(value));
  const fallback = historicalRatios.length ? historicalRatios[historicalRatios.length - 1] : 0;
  return YEARS.map((_year, index) => {
    const customers = Number(baseDrivers.newCustomers[index] || 0);
    const revenue = growthRevenue[index];
    if (revenue !== null && revenue !== undefined && customers > 0) return Number(revenue) / customers;
    return fallback;
  });
}

function createDefaultGrowthRevenuePlan() {
  return {
    topDown: baseGrowthRevenueTopDownValues(),
    revPerNewPayingCustomer: baseGrowthRevPerNewPayingCustomerValues(),
    controlPoints: {},
  };
}

function baseStrRevenueTopDownValues() {
  let latestValue = 0;
  return YEARS.map((year, index) => {
    const value = strRevenue[index];
    if (value !== null && value !== undefined) {
      latestValue = Number(value);
      return latestValue;
    }
    return year >= FIRST_FORECAST_YEAR ? latestValue : null;
  });
}

function baseStrPropertiesValues() {
  return YEARS.map(() => 0);
}

function baseStrRevPerPropertyValues() {
  return YEARS.map(() => 0);
}

function createDefaultStrRevenuePlan() {
  return {
    topDown: baseStrRevenueTopDownValues(),
    numStrProperties: baseStrPropertiesValues(),
    revPerStrProperty: baseStrRevPerPropertyValues(),
    controlPoints: {},
  };
}

function baseOtherRevenueTopDownValues() {
  let latestValue = 0;
  return YEARS.map((year, index) => {
    const value = otherRevenue[index];
    if (value !== null && value !== undefined) {
      latestValue = Number(value);
      return latestValue;
    }
    return year >= FIRST_FORECAST_YEAR ? latestValue : null;
  });
}

function createDefaultOtherRevenuePlan() {
  return {
    topDown: baseOtherRevenueTopDownValues(),
    controlPoints: {},
  };
}

function createDefaultTotalRevenuePlan() {
  return {
    topDown: clone(planRevenue),
    controlPoints: {},
  };
}

function baseLtrRevenueTopDownValues() {
  const growth = baseGrowthRevenueTopDownValues();
  const str = baseStrRevenueTopDownValues();
  const other = baseOtherRevenueTopDownValues();
  return YEARS.map((_year, index) => {
    const value = Number(planRevenue[index] || 0)
      - Number(growth[index] || 0)
      - Number(str[index] || 0)
      - Number(other[index] || 0);
    return Math.max(0, value);
  });
}

function createDefaultLtrRevenuePlan() {
  return {
    topDown: baseLtrRevenueTopDownValues(),
    edited: true,
    controlPoints: {},
  };
}

const baseCosts = {
  labor: [2770938, 5297866, 9316198, 13007538, 19229862.5, 24093112.5, 28568700, 32424900, 36646200],
  nonLabor: [2583412, 6966031, 12436889, 13792362, 15739100, 19684400, 22264400, 24120000, 27974500],
  laborDepartments: {
    admin: [2770938, 5297866, 1661527, 2489400, 2986400, 3441906.25, 3803200, 3998400, 4158400],
    product: [0, 0, 1704263, 1225100, 2273200, 3032506.25, 3776800, 4299900, 5021200],
    engineering: [0, 0, 0, 0, 0, 0, 0, 0, 0],
    growth: [0, 0, 656920, 629000, 1521800, 2251900, 3011300, 3801000, 4622300],
    sales: [0, 0, 1090265, 2658900, 4578162.5, 6111400, 7310300, 8400200, 9533800],
    marketing: [0, 0, 685404, 1188038, 1714200, 2159400, 2614300, 2920300, 3272100],
    clientSuccess: [0, 0, 1203867, 1544700, 2227500, 2622500, 3012600, 3375500, 3795800],
    customerExperience: [0, 0, 556065, 827400, 945700, 1118600, 1298300, 1485200, 1679600],
    assistanceAnimal: [0, 0, 1282492, 1773800, 2124900, 2462600, 2814000, 3179300, 3559300],
    fidoTabby: [0, 0, 475395, 671200, 858000, 892300, 927900, 965100, 1003700],
  },
  nonLaborCategories: {
    staffingCosts: [44261, 119347, 213077, 236300, 283600, 340300, 391300, 450100, 495200],
    marketingCosts: [673707, 1816616, 3243317, 3596800, 3885200, 6444300, 7552100, 7705500, 9605000],
    outsideMarketing: [1124, 3030, 5410, 6000, 6600, 6900, 7200, 7200, 7200],
    conferences: [91144, 245764, 438778, 486600, 557500, 613300, 674600, 742100, 816300],
    consultants: [422678, 1139728, 2034828, 2256600, 3135500, 3725400, 4394000, 5196300, 6161000],
    legalProfessionalServices: [56211, 151570, 270607, 300100, 315100, 330900, 347500, 364900, 383100],
    insurance: [16801, 45304, 80885, 89700, 112100, 145700, 196700, 265500, 358400],
    softwareDevelopment: [724618, 1953892, 3488405, 3868600, 4061700, 4264400, 4477400, 4701000, 4935700],
    softwareSubscriptions: [244680, 659766, 1177921, 1306300, 1567600, 1881200, 2163300, 2487900, 2861200],
    rent: [64659, 174348, 311275, 345200, 345200, 345200, 345200, 345200, 345200],
    travelMealsEntertainment: [144418, 389416, 695248, 771022, 886800, 975500, 1073200, 1180500, 1298700],
    allOther: [99111, 267250, 477138, 529140, 582200, 611300, 641900, 673800, 707500],
  },
};

const costMeta = {
  total: { label: "Total Cost", chartId: "cost-total-chart", yMax: 75, color: "#2f6f73" },
  labor: { label: "Labor", chartId: "cost-labor-chart", yMax: 45, color: "#4f7fb8" },
  nonLabor: { label: "Non-Labor", chartId: "cost-non-labor-chart", yMax: 35, color: "#6fa76b" },
};
const COST_LOCK_KEYS = ["total", "labor", "nonLabor"];
const LABOR_DEPARTMENT_KEYS = [
  "admin",
  "product",
  "engineering",
  "growth",
  "sales",
  "marketing",
  "clientSuccess",
  "customerExperience",
  "assistanceAnimal",
  "fidoTabby",
];
const laborDepartmentMeta = {
  admin: { label: "Admin", chartId: "cost-labor-admin-chart", yMax: 7, color: "#4f7fb8" },
  product: { label: "Product", chartId: "cost-labor-product-chart", yMax: 8, color: "#5c8fb7" },
  engineering: { label: "Engineering", chartId: "cost-labor-engineering-chart", yMax: 1, color: "#7d90b2" },
  growth: { label: "Growth", chartId: "cost-labor-growth-chart", yMax: 6, color: "#4b9b8e" },
  sales: { label: "Sales", chartId: "cost-labor-sales-chart", yMax: 12, color: "#6fa76b" },
  marketing: { label: "Marketing", chartId: "cost-labor-marketing-chart", yMax: 5, color: "#d39b2a" },
  clientSuccess: { label: "Client Success", chartId: "cost-labor-client-success-chart", yMax: 5, color: "#7b6fb8" },
  customerExperience: { label: "Customer Experience", chartId: "cost-labor-customer-experience-chart", yMax: 3, color: "#a96c50" },
  assistanceAnimal: { label: "Assistance Animal", chartId: "cost-labor-assistance-animal-chart", yMax: 5, color: "#73866d" },
  fidoTabby: { label: "FidoTabby", chartId: "cost-labor-fido-tabby-chart", yMax: 2, color: "#5d748c" },
};
const laborCostPerFte2025 = {
  admin: 101500,
  product: 119000,
  engineering: 133000,
  growth: 108500,
  sales: 105000,
  marketing: 101500,
  clientSuccess: 84000,
  customerExperience: 59500,
  assistanceAnimal: 56000,
  fidoTabby: 98000,
};
const NON_LABOR_CATEGORY_KEYS = [
  "staffingCosts",
  "marketingCosts",
  "outsideMarketing",
  "conferences",
  "consultants",
  "legalProfessionalServices",
  "insurance",
  "softwareDevelopment",
  "softwareSubscriptions",
  "rent",
  "travelMealsEntertainment",
  "allOther",
];
const nonLaborCategoryMeta = {
  staffingCosts: { label: "Staffing Costs", yMax: 1, color: "#4f7fb8" },
  marketingCosts: { label: "Marketing Costs", yMax: 12, color: "#6fa76b" },
  outsideMarketing: { label: "Outside Marketing", yMax: 1, color: "#8ab5a3" },
  conferences: { label: "Conferences", yMax: 2, color: "#d39b2a" },
  consultants: { label: "Consultants", yMax: 8, color: "#7b6fb8" },
  legalProfessionalServices: { label: "Legal & Professional Services", yMax: 1, color: "#a96c50" },
  insurance: { label: "Insurance", yMax: 1, color: "#73866d" },
  softwareDevelopment: { label: "Software Development", yMax: 7, color: "#4b9b8e" },
  softwareSubscriptions: { label: "Software Subscriptions", yMax: 4, color: "#5c8fb7" },
  rent: { label: "Rent", yMax: 1, color: "#b5787c" },
  travelMealsEntertainment: { label: "Travel, Meals & Entertainment", yMax: 2, color: "#5d748c" },
  allOther: { label: "All Other", yMax: 1, color: "#98a2b3" },
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
  costCharts: {},
  growthCharts: {},
  strCharts: {},
  otherRevenueCharts: {},
  revenueDrilldownCharts: {},
  profitCharts: {},
  outputCharts: {},
  currentView: "chart",
  syncingCharts: false,
  syncingCostCharts: false,
  syncingGrowthCharts: false,
  syncingStrCharts: false,
  syncingOtherRevenueCharts: false,
  syncingRevenueDrilldownCharts: false,
  syncingProfitCharts: false,
  syncingCohortCharts: false,
  undoStack: [],
  redoStack: [],
  revUnitCharts: {},
  defenseCharts: {},
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
  costDrilldown: "",
  selectedLaborAllocationKeys: [],
  laborAllocationTopLocked: false,
  laborAllocationOthersLocked: false,
  selectedLaborFteKeys: [],
  laborFteTopLocked: false,
  laborFteLockedDriver: "",
  selectedNonLaborAllocationKeys: [],
  nonLaborAllocationTopLocked: false,
  nonLaborAllocationOthersLocked: false,
  selectedNonLaborCategoryKey: "conferences",
  selectedNonLaborCategoryDepartmentKeys: [],
  nonLaborCategoryTopLocked: false,
  nonLaborCategoryOthersLocked: false,
  costChartModes: {
    laborDrilldownTopDown: "value",
    laborDepartmentMix: "value",
    laborAllocationControl: "value",
    laborFteSelectedCost: "value",
    laborFteCount: "value",
    laborFteCostPerFte: "value",
    nonLaborDrilldownTopDown: "value",
    nonLaborCategoryMix: "value",
    nonLaborAllocationControl: "value",
    nonLaborCategorySelectedSpend: "value",
    nonLaborCategoryDepartmentControl: "value",
  },
  costChartLineModes: {
    laborFteSelectedCost: "grouped",
    laborFteCount: "grouped",
    laborFteCostPerFte: "grouped",
    nonLaborCategoryDepartmentControl: "grouped",
  },
  retentionImpactChartType: "line",
  retentionImpactCumulative: false,
  activeDefenseKey: "retentionRate",
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
  profit: { parent: null, children: ["revenue", "cost"] },
  revenue: { parent: "profit", children: ["ltrRevenue", "strRevenue", "growthRevenue", "otherRevenue"] },
  ltrRevenue: { parent: "revenue", children: ["payingCustomers", "revenuePerCustomer"] },
  growthRevenue: { parent: "revenue", children: ["growthNewPayingCustomers", "growthRevPerNewPayingCustomer"] },
  growthNewPayingCustomers: { parent: "growthRevenue", children: [] },
  growthRevPerNewPayingCustomer: { parent: "growthRevenue", children: [] },
  strRevenue: { parent: "revenue", children: ["strProperties", "strRevPerProperty"] },
  strProperties: { parent: "strRevenue", children: [] },
  strRevPerProperty: { parent: "strRevenue", children: [] },
  otherRevenue: { parent: "revenue", children: [] },
  cost: { parent: "profit", children: ["laborCost", "nonLaborCost"] },
  laborCost: { parent: "cost", children: ["laborFte", "laborCostPerFte"] },
  nonLaborCost: { parent: "cost", children: [] },
  laborFte: { parent: "laborCost", children: [] },
  laborCostPerFte: { parent: "laborCost", children: [] },
  payingCustomers: { parent: "ltrRevenue", children: ["newPayingCustomers", "returningPayingCustomers"] },
  revenuePerCustomer: { parent: "ltrRevenue", children: ["profilesPerCustomer", "revenuePerProfile"] },
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
  ["profit"],
  ["revenue", "cost"],
  ["ltrRevenue", "strRevenue", "growthRevenue", "otherRevenue", "laborCost", "nonLaborCost"],
  ["payingCustomers", "revenuePerCustomer", "strProperties", "strRevPerProperty", "growthNewPayingCustomers", "growthRevPerNewPayingCustomer", "laborFte", "laborCostPerFte"],
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

function normalizeRetentionInitiative(initiative) {
  if (typeof initiative === "string") {
    return { name: initiative, teams: [] };
  }
  if (!initiative || typeof initiative !== "object") {
    return { name: "", teams: [] };
  }
  const teams = Array.isArray(initiative.teams)
    ? initiative.teams.filter(team => INITIATIVE_TEAMS.includes(team))
    : [];
  return {
    name: String(initiative.name || "").trim(),
    teams: Array.from(new Set(teams)),
  };
}

function createDefaultDefenses() {
  return LEAF_DEFENSE_KEYS.reduce((defenses, key) => {
    defenses[key] = {
      initiatives: [],
      pctTotalLines: [],
    };
    return defenses;
  }, {});
}

function defaultRetentionPctTotalLines(bands) {
  const chartYears = [FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]];
  if (bands.length < 2) return [];
  return bands.slice(0, -1).map((_band, index) => {
    const cumulative = ((index + 1) / bands.length) * 100;
    return chartYears.map(year => [year, cumulative]);
  });
}

function laborDepartmentTotalForYear(laborDepartments, index) {
  return LABOR_DEPARTMENT_KEYS.reduce((sum, key) => {
    return sum + Number(laborDepartments?.[key]?.[index] || 0);
  }, 0);
}

function nonLaborCategoryTotalForYear(nonLaborCategories, index) {
  return NON_LABOR_CATEGORY_KEYS.reduce((sum, key) => {
    return sum + Number(nonLaborCategories?.[key]?.[index] || 0);
  }, 0);
}

function scaledBaseLaborDepartments(targetLaborValues) {
  const departments = {};
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    departments[key] = YEARS.map((_year, index) => {
      const baseValue = Number(baseCosts.laborDepartments[key][index] || 0);
      const baseTotal = Number(baseCosts.labor[index] || 0);
      const targetTotal = Number(targetLaborValues?.[index] ?? baseTotal);
      if (!baseTotal) return baseValue;
      return baseValue * (targetTotal / baseTotal);
    });
  });
  return departments;
}

function baseLaborCostPerFteValue(key, index) {
  const baseValue = Number(laborCostPerFte2025[key] || 100000);
  return baseValue * Math.pow(1.03, YEARS[index] - 2025);
}

function createBaseLaborCostPerFte() {
  const values = {};
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    values[key] = YEARS.map((_year, index) => baseLaborCostPerFteValue(key, index));
  });
  return values;
}

function createBaseLaborFte(laborDepartments = baseCosts.laborDepartments, laborCostPerFte = createBaseLaborCostPerFte()) {
  const values = {};
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    values[key] = YEARS.map((_year, index) => {
      const costPerFte = Number(laborCostPerFte?.[key]?.[index] || baseLaborCostPerFteValue(key, index));
      const laborCost = Number(laborDepartments?.[key]?.[index] || 0);
      return costPerFte > 0 ? laborCost / costPerFte : 0;
    });
  });
  return values;
}

function scaledBaseNonLaborCategories(targetNonLaborValues) {
  const categories = {};
  NON_LABOR_CATEGORY_KEYS.forEach(key => {
    categories[key] = YEARS.map((_year, index) => {
      const baseValue = Number(baseCosts.nonLaborCategories[key][index] || 0);
      const baseTotal = Number(baseCosts.nonLabor[index] || 0);
      const targetTotal = Number(targetNonLaborValues?.[index] ?? baseTotal);
      if (!baseTotal) return baseValue;
      return baseValue * (targetTotal / baseTotal);
    });
  });
  return categories;
}

function baseLaborDepartmentShareForYear(departmentKey, index) {
  const referenceIndex = Math.max(0, YEARS.indexOf(2025));
  const baseTotal = Number(baseCosts.labor[referenceIndex] || 0);
  if (baseTotal > 0) return Number(baseCosts.laborDepartments[departmentKey]?.[referenceIndex] || 0) / baseTotal;
  return 1 / LABOR_DEPARTMENT_KEYS.length;
}

function createBaseNonLaborCategoryDepartments(nonLaborCategories = baseCosts.nonLaborCategories) {
  const categoryDepartments = {};
  NON_LABOR_CATEGORY_KEYS.forEach(categoryKey => {
    categoryDepartments[categoryKey] = {};
    LABOR_DEPARTMENT_KEYS.forEach(departmentKey => {
      categoryDepartments[categoryKey][departmentKey] = YEARS.map((_year, index) => {
        const categoryValue = Number(nonLaborCategories?.[categoryKey]?.[index] ?? baseCosts.nonLaborCategories[categoryKey][index] ?? 0);
        return categoryValue * baseLaborDepartmentShareForYear(departmentKey, index);
      });
    });
  });
  return categoryDepartments;
}

function syncLaborTotalsFromDepartments(scenario) {
  scenario.costs.labor = YEARS.map((_year, index) => {
    return laborDepartmentTotalForYear(scenario.costs.laborDepartments, index);
  });
}

function normalizeCostPlan(scenario) {
  if (!scenario.costs || typeof scenario.costs !== "object") {
    scenario.costs = clone(baseCosts);
  }
  ["labor", "nonLabor"].forEach(key => {
    if (!Array.isArray(scenario.costs[key])) {
      scenario.costs[key] = clone(baseCosts[key]);
    }
    scenario.costs[key] = YEARS.map((_year, index) => {
      const value = Number(scenario.costs[key][index]);
      return Number.isFinite(value) ? value : baseCosts[key][index];
    });
  });
  if (!scenario.costs.laborDepartments || typeof scenario.costs.laborDepartments !== "object") {
    scenario.costs.laborDepartments = scaledBaseLaborDepartments(scenario.costs.labor);
  }
  if (!scenario.costs.laborCostPerFte || typeof scenario.costs.laborCostPerFte !== "object") {
    scenario.costs.laborCostPerFte = createBaseLaborCostPerFte();
  }
  if (!scenario.costs.laborFte || typeof scenario.costs.laborFte !== "object") {
    scenario.costs.laborFte = createBaseLaborFte(scenario.costs.laborDepartments, scenario.costs.laborCostPerFte);
  }
  if (!scenario.costs.nonLaborCategories || typeof scenario.costs.nonLaborCategories !== "object") {
    scenario.costs.nonLaborCategories = scaledBaseNonLaborCategories(scenario.costs.nonLabor);
  }
  if (!scenario.costs.nonLaborCategoryDepartments || typeof scenario.costs.nonLaborCategoryDepartments !== "object") {
    scenario.costs.nonLaborCategoryDepartments = createBaseNonLaborCategoryDepartments(scenario.costs.nonLaborCategories);
  }
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    if (!Array.isArray(scenario.costs.laborDepartments[key])) {
      scenario.costs.laborDepartments[key] = clone(baseCosts.laborDepartments[key]);
    }
    scenario.costs.laborDepartments[key] = YEARS.map((_year, index) => {
      const value = Number(scenario.costs.laborDepartments[key][index]);
      return Number.isFinite(value) ? value : baseCosts.laborDepartments[key][index];
    });
  });
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    if (!Array.isArray(scenario.costs.laborCostPerFte[key])) {
      scenario.costs.laborCostPerFte[key] = YEARS.map((_year, index) => baseLaborCostPerFteValue(key, index));
    }
    scenario.costs.laborCostPerFte[key] = YEARS.map((_year, index) => {
      const value = Number(scenario.costs.laborCostPerFte[key][index]);
      return Number.isFinite(value) && value > 0 ? value : baseLaborCostPerFteValue(key, index);
    });
    if (!Array.isArray(scenario.costs.laborFte[key])) {
      scenario.costs.laborFte[key] = createBaseLaborFte(scenario.costs.laborDepartments, scenario.costs.laborCostPerFte)[key];
    }
    scenario.costs.laborFte[key] = YEARS.map((_year, index) => {
      const value = Number(scenario.costs.laborFte[key][index]);
      return Number.isFinite(value) && value >= 0 ? value : 0;
    });
  });
  NON_LABOR_CATEGORY_KEYS.forEach(key => {
    if (!Array.isArray(scenario.costs.nonLaborCategories[key])) {
      scenario.costs.nonLaborCategories[key] = clone(baseCosts.nonLaborCategories[key]);
    }
    scenario.costs.nonLaborCategories[key] = YEARS.map((_year, index) => {
      const value = Number(scenario.costs.nonLaborCategories[key][index]);
      return Number.isFinite(value) ? value : baseCosts.nonLaborCategories[key][index];
    });
  });
  NON_LABOR_CATEGORY_KEYS.forEach(categoryKey => {
    if (!scenario.costs.nonLaborCategoryDepartments[categoryKey] || typeof scenario.costs.nonLaborCategoryDepartments[categoryKey] !== "object") {
      scenario.costs.nonLaborCategoryDepartments[categoryKey] = createBaseNonLaborCategoryDepartments(scenario.costs.nonLaborCategories)[categoryKey];
    }
    LABOR_DEPARTMENT_KEYS.forEach(departmentKey => {
      if (!Array.isArray(scenario.costs.nonLaborCategoryDepartments[categoryKey][departmentKey])) {
        scenario.costs.nonLaborCategoryDepartments[categoryKey][departmentKey] = createBaseNonLaborCategoryDepartments(scenario.costs.nonLaborCategories)[categoryKey][departmentKey];
      }
      scenario.costs.nonLaborCategoryDepartments[categoryKey][departmentKey] = YEARS.map((_year, index) => {
        const value = Number(scenario.costs.nonLaborCategoryDepartments[categoryKey][departmentKey][index]);
        return Number.isFinite(value) && value >= 0 ? value : 0;
      });
    });
  });
  if (!scenario.costControlPoints || typeof scenario.costControlPoints !== "object") {
    scenario.costControlPoints = {};
  }
  if (!scenario.costPctTotalLines || typeof scenario.costPctTotalLines !== "object") {
    scenario.costPctTotalLines = {};
  }
  if (!scenario.costLocks || typeof scenario.costLocks !== "object") {
    scenario.costLocks = {};
  }
  COST_LOCK_KEYS.forEach(key => {
    scenario.costLocks[key] = Boolean(scenario.costLocks[key]);
  });
  if (scenario.costLocks.labor && scenario.costLocks.nonLabor) {
    scenario.costLocks.nonLabor = false;
  }
  Object.keys(scenario.costControlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.costControlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.costControlPoints[key];
    } else {
      scenario.costControlPoints[key] = controlPoints;
    }
  });
}

function normalizeRevenuePaths(scenario) {
  if (!scenario.revenuePaths || typeof scenario.revenuePaths !== "object") {
    scenario.revenuePaths = {};
  }
  if (!scenario.revenuePaths.total || typeof scenario.revenuePaths.total !== "object") {
    scenario.revenuePaths.total = createDefaultTotalRevenuePlan();
  }
  if (!scenario.revenuePaths.ltr || typeof scenario.revenuePaths.ltr !== "object") {
    scenario.revenuePaths.ltr = createDefaultLtrRevenuePlan();
  }
  if (!scenario.revenuePaths.growth || typeof scenario.revenuePaths.growth !== "object") {
    scenario.revenuePaths.growth = createDefaultGrowthRevenuePlan();
  }
  if (!scenario.revenuePaths.str || typeof scenario.revenuePaths.str !== "object") {
    scenario.revenuePaths.str = createDefaultStrRevenuePlan();
  }
  if (!scenario.revenuePaths.other || typeof scenario.revenuePaths.other !== "object") {
    scenario.revenuePaths.other = createDefaultOtherRevenuePlan();
  }
  const totalDefaults = createDefaultTotalRevenuePlan();
  if (!Array.isArray(scenario.revenuePaths.total.topDown)) {
    scenario.revenuePaths.total.topDown = clone(totalDefaults.topDown);
  }
  scenario.revenuePaths.total.topDown = YEARS.map((_year, index) => {
    const rawValue = scenario.revenuePaths.total.topDown[index];
    const value = Number(rawValue);
    return Number.isFinite(value) ? Math.max(0, value) : totalDefaults.topDown[index];
  });
  if (!scenario.revenuePaths.total.controlPoints || typeof scenario.revenuePaths.total.controlPoints !== "object") {
    scenario.revenuePaths.total.controlPoints = {};
  }
  Object.keys(scenario.revenuePaths.total.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.revenuePaths.total.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.revenuePaths.total.controlPoints[key];
    } else {
      scenario.revenuePaths.total.controlPoints[key] = controlPoints;
    }
  });
  const ltrDefaults = createDefaultLtrRevenuePlan();
  if (!Array.isArray(scenario.revenuePaths.ltr.topDown)) {
    scenario.revenuePaths.ltr.topDown = clone(ltrDefaults.topDown);
  }
  scenario.revenuePaths.ltr.topDown = YEARS.map((_year, index) => {
    const rawValue = scenario.revenuePaths.ltr.topDown[index];
    if (rawValue === null || rawValue === undefined || rawValue === "") return ltrDefaults.topDown[index];
    const value = Number(rawValue);
    return Number.isFinite(value) ? Math.max(0, value) : ltrDefaults.topDown[index];
  });
  scenario.revenuePaths.ltr.edited = true;
  if (!scenario.revenuePaths.ltr.controlPoints || typeof scenario.revenuePaths.ltr.controlPoints !== "object") {
    scenario.revenuePaths.ltr.controlPoints = {};
  }
  Object.keys(scenario.revenuePaths.ltr.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.revenuePaths.ltr.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.revenuePaths.ltr.controlPoints[key];
    } else {
      scenario.revenuePaths.ltr.controlPoints[key] = controlPoints;
    }
  });
  const defaults = createDefaultGrowthRevenuePlan();
  ["topDown", "revPerNewPayingCustomer"].forEach(key => {
    if (!Array.isArray(scenario.revenuePaths.growth[key])) {
      scenario.revenuePaths.growth[key] = clone(defaults[key]);
    }
    scenario.revenuePaths.growth[key] = YEARS.map((_year, index) => {
      const rawValue = scenario.revenuePaths.growth[key][index];
      if (rawValue === null || rawValue === undefined || rawValue === "") {
        return key === "topDown" ? null : defaults[key][index];
      }
      const value = Number(rawValue);
      return Number.isFinite(value) ? Math.max(0, value) : defaults[key][index];
    });
  });
  if (!scenario.revenuePaths.growth.controlPoints || typeof scenario.revenuePaths.growth.controlPoints !== "object") {
    scenario.revenuePaths.growth.controlPoints = {};
  }
  Object.keys(scenario.revenuePaths.growth.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.revenuePaths.growth.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.revenuePaths.growth.controlPoints[key];
    } else {
      scenario.revenuePaths.growth.controlPoints[key] = controlPoints;
    }
  });
  const strDefaults = createDefaultStrRevenuePlan();
  ["topDown", "numStrProperties", "revPerStrProperty"].forEach(key => {
    if (!Array.isArray(scenario.revenuePaths.str[key])) {
      scenario.revenuePaths.str[key] = clone(strDefaults[key]);
    }
    scenario.revenuePaths.str[key] = YEARS.map((_year, index) => {
      const rawValue = scenario.revenuePaths.str[key][index];
      if (rawValue === null || rawValue === undefined || rawValue === "") {
        return key === "topDown" ? null : strDefaults[key][index];
      }
      const value = Number(rawValue);
      return Number.isFinite(value) ? Math.max(0, value) : strDefaults[key][index];
    });
  });
  if (!scenario.revenuePaths.str.controlPoints || typeof scenario.revenuePaths.str.controlPoints !== "object") {
    scenario.revenuePaths.str.controlPoints = {};
  }
  Object.keys(scenario.revenuePaths.str.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.revenuePaths.str.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.revenuePaths.str.controlPoints[key];
    } else {
      scenario.revenuePaths.str.controlPoints[key] = controlPoints;
    }
  });
  const otherDefaults = createDefaultOtherRevenuePlan();
  if (!Array.isArray(scenario.revenuePaths.other.topDown)) {
    scenario.revenuePaths.other.topDown = clone(otherDefaults.topDown);
  }
  scenario.revenuePaths.other.topDown = YEARS.map((_year, index) => {
    const rawValue = scenario.revenuePaths.other.topDown[index];
    if (rawValue === null || rawValue === undefined || rawValue === "") return null;
    const value = Number(rawValue);
    return Number.isFinite(value) ? Math.max(0, value) : otherDefaults.topDown[index];
  });
  if (!scenario.revenuePaths.other.controlPoints || typeof scenario.revenuePaths.other.controlPoints !== "object") {
    scenario.revenuePaths.other.controlPoints = {};
  }
  Object.keys(scenario.revenuePaths.other.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.revenuePaths.other.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.revenuePaths.other.controlPoints[key];
    } else {
      scenario.revenuePaths.other.controlPoints[key] = controlPoints;
    }
  });
}

function createDefaultProfitPlan() {
  return {
    profit: YEARS.map(() => null),
    revenue: YEARS.map(() => null),
    profitEdited: false,
    revenueEdited: false,
    locks: {
      profit: false,
      revenue: false,
      cost: false,
    },
    controlPoints: {},
  };
}

function normalizeProfitPlan(scenario) {
  if (!scenario.profitPlan || typeof scenario.profitPlan !== "object") {
    scenario.profitPlan = createDefaultProfitPlan();
  }
  const defaults = createDefaultProfitPlan();
  ["profit", "revenue"].forEach(key => {
    if (!Array.isArray(scenario.profitPlan[key])) {
      scenario.profitPlan[key] = clone(defaults[key]);
    }
    scenario.profitPlan[key] = YEARS.map((_year, index) => {
      const rawValue = scenario.profitPlan[key][index];
      if (rawValue === null || rawValue === undefined || rawValue === "") return null;
      const value = Number(rawValue);
      return Number.isFinite(value) ? value : null;
    });
  });
  scenario.profitPlan.profitEdited = Boolean(scenario.profitPlan.profitEdited);
  scenario.profitPlan.revenueEdited = Boolean(scenario.profitPlan.revenueEdited);
  if (!scenario.profitPlan.locks || typeof scenario.profitPlan.locks !== "object") {
    scenario.profitPlan.locks = clone(defaults.locks);
  }
  ["profit", "revenue", "cost"].forEach(key => {
    scenario.profitPlan.locks[key] = Boolean(scenario.profitPlan.locks[key]);
  });
  if (scenario.profitPlan.locks.revenue && scenario.profitPlan.locks.cost) {
    scenario.profitPlan.locks.cost = false;
  }
  if (!scenario.profitPlan.controlPoints || typeof scenario.profitPlan.controlPoints !== "object") {
    scenario.profitPlan.controlPoints = {};
  }
  Object.keys(scenario.profitPlan.controlPoints).forEach(key => {
    const controlPoints = Array.from(new Set((scenario.profitPlan.controlPoints[key] || [])
      .map(Number)
      .filter(Number.isFinite)))
      .sort((left, right) => left - right);
    if (controlPoints.length < 2) {
      delete scenario.profitPlan.controlPoints[key];
    } else {
      scenario.profitPlan.controlPoints[key] = controlPoints;
    }
  });
}

function makeBaseScenario(name = "Finance Base Case") {
  const scenario = {
    id: "base",
    name,
    drivers: clone(baseDrivers),
    costs: clone(baseCosts),
    costControlPoints: {},
    costPctTotalLines: {},
    costLocks: {},
    newCustomerSource: "topDown",
    newCustomerDrilldown: createDefaultNewCustomerDrilldown(),
    revUnitPlan: createDefaultRevUnitPlan(),
    revenuePaths: {
      total: createDefaultTotalRevenuePlan(),
      ltr: createDefaultLtrRevenuePlan(),
      growth: createDefaultGrowthRevenuePlan(),
      str: createDefaultStrRevenuePlan(),
      other: createDefaultOtherRevenuePlan(),
    },
    profitPlan: createDefaultProfitPlan(),
    defenses: createDefaultDefenses(),
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
  normalizeCostPlan(scenario);
  normalizeRevenuePaths(scenario);
  normalizeProfitPlan(scenario);
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
  if (!scenario.defenses) {
    scenario.defenses = createDefaultDefenses();
  }
  if (scenario.defenses.retention && !scenario.defenses.retentionRate) {
    scenario.defenses.retentionRate = clone(scenario.defenses.retention);
  }
  LEAF_DEFENSE_KEYS.forEach(key => {
    if (!scenario.defenses[key]) {
      scenario.defenses[key] = { initiatives: [], pctTotalLines: [] };
    }
    if (!Array.isArray(scenario.defenses[key].initiatives)) {
      scenario.defenses[key].initiatives = [];
    }
    scenario.defenses[key].initiatives = scenario.defenses[key].initiatives
      .map(normalizeRetentionInitiative)
      .filter(initiative => initiative.name);
    if (!Array.isArray(scenario.defenses[key].pctTotalLines)) {
      scenario.defenses[key].pctTotalLines = [];
    }
  });
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
    costs: clone(scenario.costs || baseCosts),
    costControlPoints: clone(scenario.costControlPoints || {}),
    costPctTotalLines: clone(scenario.costPctTotalLines || {}),
    costLocks: clone(scenario.costLocks || {}),
    newCustomerSource: scenario.newCustomerSource,
    newCustomerDrilldown: clone(scenario.newCustomerDrilldown),
    revUnitPlan: clone(scenario.revUnitPlan),
    revenuePaths: clone(scenario.revenuePaths || { growth: createDefaultGrowthRevenuePlan(), str: createDefaultStrRevenuePlan(), other: createDefaultOtherRevenuePlan() }),
    profitPlan: clone(scenario.profitPlan || createDefaultProfitPlan()),
    defenses: clone(scenario.defenses || createDefaultDefenses()),
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

function costValues(scenario, key) {
  normalizeCostPlan(scenario);
  if (key === "total") {
    return YEARS.map((_year, index) => scenario.costs.labor[index] + scenario.costs.nonLabor[index]);
  }
  return scenario.costs[key] || clone(baseCosts[key]);
}

function laborDepartmentValues(scenario, key) {
  normalizeCostPlan(scenario);
  return scenario.costs.laborDepartments[key] || clone(baseCosts.laborDepartments[key]);
}

function laborBottomUpValues(scenario) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => laborDepartmentTotalForYear(scenario.costs.laborDepartments, index));
}

function selectedLaborAllocationKeys() {
  return (state.selectedLaborAllocationKeys || []).filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
}

function laborDepartmentGroupTotalForYear(laborDepartments, keys, index) {
  return keys.reduce((sum, key) => sum + Number(laborDepartments?.[key]?.[index] || 0), 0);
}

function laborAllocationControlValues(scenario) {
  normalizeCostPlan(scenario);
  const keys = selectedLaborAllocationKeys();
  return YEARS.map((_year, index) => laborDepartmentGroupTotalForYear(scenario.costs.laborDepartments, keys, index));
}

function selectedLaborFteKeys() {
  return (state.selectedLaborFteKeys || []).filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
}

function activeLaborFteKeys() {
  const selected = selectedLaborFteKeys();
  return selected.length ? selected : LABOR_DEPARTMENT_KEYS;
}

function laborFteFilterLabel() {
  const selected = selectedLaborFteKeys();
  if (!selected.length) return "All Departments";
  if (selected.length === 1) return displayLabel(laborDepartmentMeta[selected[0]].label);
  return `${selected.length} Departments`;
}

function laborFteCountValues(scenario, keys = activeLaborFteKeys()) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => {
    return keys.reduce((sum, key) => sum + Number(scenario.costs.laborFte?.[key]?.[index] || 0), 0);
  });
}

function laborFteSelectedCostValues(scenario, keys = activeLaborFteKeys()) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => laborDepartmentGroupTotalForYear(scenario.costs.laborDepartments, keys, index));
}

function laborFteBuildCostValues(scenario, keys = activeLaborFteKeys()) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => {
    return keys.reduce((sum, key) => {
      const fte = Number(scenario.costs.laborFte?.[key]?.[index] || 0);
      const costPerFte = Number(scenario.costs.laborCostPerFte?.[key]?.[index] || 0);
      return sum + (fte * costPerFte);
    }, 0);
  });
}

function laborFteCostPerFteValues(scenario, keys = activeLaborFteKeys()) {
  const fteValues = laborFteCountValues(scenario, keys);
  const costValues = laborFteBuildCostValues(scenario, keys);
  return YEARS.map((_year, index) => {
    const fte = Number(fteValues[index] || 0);
    return fte > 0 ? costValues[index] / fte : 0;
  });
}

function distributeLaborFteAcrossKeys(scenario, index, keys, targetFte) {
  normalizeCostPlan(scenario);
  if (!keys.length) return;
  const actualValue = Math.max(0, Number(targetFte) || 0);
  const currentTotal = keys.reduce((sum, key) => sum + Number(scenario.costs.laborFte[key][index] || 0), 0);
  const baseFte = createBaseLaborFte(baseCosts.laborDepartments, createBaseLaborCostPerFte());
  const baseTotal = keys.reduce((sum, key) => sum + Number(baseFte[key][index] || 0), 0);
  keys.forEach(key => {
    const currentValue = Number(scenario.costs.laborFte[key][index] || 0);
    const baseValue = Number(baseFte[key][index] || 0);
    const share = currentTotal > 0
      ? currentValue / currentTotal
      : baseTotal > 0 ? baseValue / baseTotal : 1 / keys.length;
    scenario.costs.laborFte[key][index] = actualValue * share;
  });
}

function scaleLaborCostPerFteAcrossKeys(scenario, index, keys, targetCostPerFte) {
  normalizeCostPlan(scenario);
  if (!keys.length) return;
  const actualValue = Math.max(0, Number(targetCostPerFte) || 0);
  const currentBlended = laborFteCostPerFteValues(scenario, keys)[index];
  const scale = currentBlended > 0 ? actualValue / currentBlended : 1;
  keys.forEach(key => {
    scenario.costs.laborCostPerFte[key][index] = Math.max(0, Number(scenario.costs.laborCostPerFte[key][index] || 0) * scale);
  });
}

function scaleLaborFteAcrossKeys(scenario, index, keys, scale) {
  normalizeCostPlan(scenario);
  if (!keys.length || !Number.isFinite(scale)) return;
  keys.forEach(key => {
    scenario.costs.laborFte[key][index] = Math.max(0, Number(scenario.costs.laborFte[key][index] || 0) * scale);
  });
}

function scaleLaborCostPerFteValuesAcrossKeys(scenario, index, keys, scale) {
  normalizeCostPlan(scenario);
  if (!keys.length || !Number.isFinite(scale)) return;
  keys.forEach(key => {
    scenario.costs.laborCostPerFte[key][index] = Math.max(0, Number(scenario.costs.laborCostPerFte[key][index] || 0) * scale);
  });
}

function adjustLaborFteToBuildCost(scenario, index, keys, targetCost) {
  const currentBuild = laborFteBuildCostValues(scenario, keys)[index];
  if (currentBuild > 0) {
    scaleLaborFteAcrossKeys(scenario, index, keys, Math.max(0, Number(targetCost) || 0) / currentBuild);
    return;
  }
  const averageCostPerFte = keys.reduce((sum, key) => sum + Number(scenario.costs.laborCostPerFte[key][index] || baseLaborCostPerFteValue(key, index)), 0) / keys.length;
  distributeLaborFteAcrossKeys(scenario, index, keys, averageCostPerFte > 0 ? Math.max(0, Number(targetCost) || 0) / averageCostPerFte : 0);
}

function adjustLaborCostPerFteToBuildCost(scenario, index, keys, targetCost) {
  const currentBuild = laborFteBuildCostValues(scenario, keys)[index];
  if (currentBuild <= 0) return;
  scaleLaborCostPerFteValuesAcrossKeys(scenario, index, keys, Math.max(0, Number(targetCost) || 0) / currentBuild);
}

function laborFteTopMatchesBottomUp(scenario = activeScenario()) {
  const keys = activeLaborFteKeys();
  const selectedCostValues = laborFteSelectedCostValues(scenario, keys);
  const buildCostValues = laborFteBuildCostValues(scenario, keys);
  return YEARS.every((year, index) => {
    if (!editableCostYear(year)) return true;
    return Math.abs(Number(selectedCostValues[index] || 0) - Number(buildCostValues[index] || 0)) < 1;
  });
}

function setLaborFteBuildCostForYear(scenario, index, keys, targetCost) {
  normalizeCostPlan(scenario);
  const actualValue = Math.max(0, Number(targetCost) || 0);
  const currentBuild = laborFteBuildCostValues(scenario, keys)[index];
  if (state.laborFteLockedDriver === "fte") {
    adjustLaborCostPerFteToBuildCost(scenario, index, keys, actualValue);
  } else if (state.laborFteLockedDriver === "costPerFte") {
    adjustLaborFteToBuildCost(scenario, index, keys, actualValue);
  } else if (currentBuild > 0) {
    const sharedScale = Math.sqrt(actualValue / currentBuild);
    scaleLaborFteAcrossKeys(scenario, index, keys, sharedScale);
    scaleLaborCostPerFteValuesAcrossKeys(scenario, index, keys, sharedScale);
  } else {
    adjustLaborFteToBuildCost(scenario, index, keys, actualValue);
  }
  addCostControlPointIfControlled("laborFteCount", YEARS[index]);
  addCostControlPointIfControlled("laborFteCostPerFte", YEARS[index]);
  addCostControlPointIfControlled("laborFteBuildCost", YEARS[index]);
}

function laborFteDependentEditBlocked(chartKey) {
  return state.laborFteTopLocked && (
    (chartKey === "laborFteCount" && state.laborFteLockedDriver === "costPerFte")
    || (chartKey === "laborFteCostPerFte" && state.laborFteLockedDriver === "fte")
  );
}

function rejectLaborFteDependentEdit(chart) {
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  syncCostCharts();
  renderCostDrilldownView();
}

function setLaborDepartmentAllocationToFteBuildForYear(scenario, index, keys, targetCost = null) {
  normalizeCostPlan(scenario);
  const targetValue = Number.isFinite(Number(targetCost))
    ? Math.max(0, Number(targetCost))
    : laborFteBuildCostValues(scenario, keys)[index];
  const buildValues = keys.map(key => {
    return Number(scenario.costs.laborFte[key][index] || 0) * Number(scenario.costs.laborCostPerFte[key][index] || 0);
  });
  const buildTotal = buildValues.reduce((sum, value) => sum + value, 0);
  keys.forEach(key => {
    const buildValue = Number(scenario.costs.laborFte[key][index] || 0) * Number(scenario.costs.laborCostPerFte[key][index] || 0);
    const share = buildTotal > 0 ? buildValue / buildTotal : 1 / keys.length;
    scenario.costs.laborDepartments[key][index] = targetValue * share;
    addCostControlPointIfControlled(laborDepartmentChartKey(key), YEARS[index]);
  });
  syncLaborTotalsFromDepartments(scenario);
  addCostControlPointIfControlled("laborBottomUp", YEARS[index]);
  addCostControlPointIfControlled("laborAllocationControl", YEARS[index]);
  addCostControlPointIfControlled("labor", YEARS[index]);
  addCostControlPointIfControlled("total", YEARS[index]);
}

function nonLaborCategoryValues(scenario, key) {
  normalizeCostPlan(scenario);
  return scenario.costs.nonLaborCategories[key] || clone(baseCosts.nonLaborCategories[key]);
}

function nonLaborBottomUpValues(scenario) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => nonLaborCategoryTotalForYear(scenario.costs.nonLaborCategories, index));
}

function selectedNonLaborAllocationKeys() {
  return (state.selectedNonLaborAllocationKeys || []).filter(key => NON_LABOR_CATEGORY_KEYS.includes(key));
}

function selectedNonLaborCategoryKey() {
  return NON_LABOR_CATEGORY_KEYS.includes(state.selectedNonLaborCategoryKey)
    ? state.selectedNonLaborCategoryKey
    : "conferences";
}

function selectedNonLaborCategoryDepartmentKeys() {
  return (state.selectedNonLaborCategoryDepartmentKeys || []).filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
}

function nonLaborCategoryGroupTotalForYear(nonLaborCategories, keys, index) {
  return keys.reduce((sum, key) => sum + Number(nonLaborCategories?.[key]?.[index] || 0), 0);
}

function nonLaborAllocationControlValues(scenario) {
  normalizeCostPlan(scenario);
  const keys = selectedNonLaborAllocationKeys();
  return YEARS.map((_year, index) => nonLaborCategoryGroupTotalForYear(scenario.costs.nonLaborCategories, keys, index));
}

function nonLaborCategoryDepartmentValues(scenario, categoryKey, departmentKey) {
  normalizeCostPlan(scenario);
  return scenario.costs.nonLaborCategoryDepartments?.[categoryKey]?.[departmentKey] || YEARS.map(() => 0);
}

function nonLaborCategoryDepartmentBottomUpValues(scenario, categoryKey = selectedNonLaborCategoryKey()) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => {
    return LABOR_DEPARTMENT_KEYS.reduce((sum, departmentKey) => {
      return sum + Number(scenario.costs.nonLaborCategoryDepartments?.[categoryKey]?.[departmentKey]?.[index] || 0);
    }, 0);
  });
}

function nonLaborCategoryDepartmentGroupTotalForYear(categoryDepartments, categoryKey, departmentKeys, index) {
  return departmentKeys.reduce((sum, departmentKey) => {
    return sum + Number(categoryDepartments?.[categoryKey]?.[departmentKey]?.[index] || 0);
  }, 0);
}

function nonLaborCategoryDepartmentControlValues(scenario) {
  normalizeCostPlan(scenario);
  const categoryKey = selectedNonLaborCategoryKey();
  const keys = selectedNonLaborCategoryDepartmentKeys();
  return YEARS.map((_year, index) => {
    return nonLaborCategoryDepartmentGroupTotalForYear(scenario.costs.nonLaborCategoryDepartments, categoryKey, keys, index);
  });
}

function controlledCostPairs(scenario, key, pairs) {
  normalizeCostPlan(scenario);
  const stored = scenario.costControlPoints?.[key];
  if (!Array.isArray(stored)) return pairs;
  const visibleX = new Set([
    ...historicalCostYears(),
    ...stored.map(Number),
  ]);
  const visiblePairs = pairs.filter(([x]) => visibleX.has(Number(x)));
  if (visiblePairs.length >= 2) return visiblePairs;
  if (pairs.length <= 2) return pairs;
  return [pairs[0], pairs[pairs.length - 1]];
}

function rememberCostControlPoints(key, data) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  scenario.costControlPoints[key] = (data || [])
    .map(([x]) => Number(x))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
}

function rememberCostControlYears(key, years) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  scenario.costControlPoints[key] = Array.from(new Set((years || [])
    .map(Number)
    .filter(Number.isFinite)))
    .sort((left, right) => left - right);
}

function editableCostYears() {
  return YEARS.filter(editableCostYear);
}

function historicalCostYears() {
  return YEARS.filter(year => !editableCostYear(year));
}

function addCostControlPointIfControlled(key, x) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (!Array.isArray(scenario.costControlPoints[key])) return;
  const next = new Set(scenario.costControlPoints[key].map(Number));
  next.add(Number(x));
  scenario.costControlPoints[key] = Array.from(next)
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
}

function laborTopDownMatchesBottomUp(scenario = activeScenario()) {
  normalizeCostPlan(scenario);
  return YEARS.every((year, index) => {
    if (!editableCostYear(year)) return true;
    const topDown = Number(scenario.costs.labor[index] || 0);
    const bottomUp = laborDepartmentTotalForYear(scenario.costs.laborDepartments, index);
    return Math.abs(topDown - bottomUp) < 1;
  });
}

function setLaborDepartmentTotalForYear(scenario, index, targetTotal) {
  normalizeCostPlan(scenario);
  setLaborDepartmentBottomUpTotalForYear(scenario, index, targetTotal);
  scenario.costs.labor[index] = Math.max(0, Number(targetTotal) || 0);
}

function setLaborDepartmentBottomUpTotalForYear(scenario, index, targetTotal, { lockOthers = false, selectedKeys = [] } = {}) {
  normalizeCostPlan(scenario);
  const actualValue = Math.max(0, Number(targetTotal) || 0);
  const selected = selectedKeys.filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
  if (lockOthers && selected.length) {
    const unselected = LABOR_DEPARTMENT_KEYS.filter(key => !selected.includes(key));
    const unselectedTotal = laborDepartmentGroupTotalForYear(scenario.costs.laborDepartments, unselected, index);
    distributeLaborTotalAcrossKeys(scenario, index, selected, Math.max(0, actualValue - unselectedTotal));
    addCostControlPointIfControlled("laborBottomUp", YEARS[index]);
    addCostControlPointIfControlled("laborAllocationControl", YEARS[index]);
    return;
  }
  const currentTotal = laborDepartmentTotalForYear(scenario.costs.laborDepartments, index);
  const baseTotal = Number(baseCosts.labor[index] || 0);
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    const currentValue = Number(scenario.costs.laborDepartments[key][index] || 0);
    const baseValue = Number(baseCosts.laborDepartments[key][index] || 0);
    const share = currentTotal > 0
      ? currentValue / currentTotal
      : baseTotal > 0 ? baseValue / baseTotal : 1 / LABOR_DEPARTMENT_KEYS.length;
    scenario.costs.laborDepartments[key][index] = actualValue * share;
    addCostControlPointIfControlled(`laborDepartment:${key}`, YEARS[index]);
  });
  addCostControlPointIfControlled("laborBottomUp", YEARS[index]);
}

function setLaborDepartmentValue(scenario, departmentKey, year, value) {
  const index = YEARS.indexOf(year);
  if (index < 0 || !editableCostYear(year) || !Number.isFinite(value)) return;
  normalizeCostPlan(scenario);
  scenario.costs.laborDepartments[departmentKey][index] = Math.max(0, value * 1000000);
  reconcileLaborDepartmentsToTopDown(scenario, index);
}

function reconcileLaborDepartmentsToTopDown(scenario, index) {
  normalizeCostPlan(scenario);
  const targetTotal = Math.max(0, Number(scenario.costs.labor[index] || 0));
  const currentTotal = laborDepartmentTotalForYear(scenario.costs.laborDepartments, index);
  const baseTotal = Number(baseCosts.labor[index] || 0);
  if (targetTotal === 0) {
    LABOR_DEPARTMENT_KEYS.forEach(key => {
      scenario.costs.laborDepartments[key][index] = 0;
    });
    return;
  }
  if (currentTotal > 0) {
    const scale = targetTotal / currentTotal;
    LABOR_DEPARTMENT_KEYS.forEach(key => {
      scenario.costs.laborDepartments[key][index] = Math.max(0, Number(scenario.costs.laborDepartments[key][index] || 0) * scale);
    });
    return;
  }
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    const baseValue = Number(baseCosts.laborDepartments[key][index] || 0);
    const share = baseTotal > 0 ? baseValue / baseTotal : 1 / LABOR_DEPARTMENT_KEYS.length;
    scenario.costs.laborDepartments[key][index] = targetTotal * share;
  });
}

function setLaborDepartmentValuesForYear(scenario, index, valuesByDepartment) {
  normalizeCostPlan(scenario);
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    const value = Number(valuesByDepartment?.[key]);
    scenario.costs.laborDepartments[key][index] = Number.isFinite(value) ? Math.max(0, value) : 0;
    addCostControlPointIfControlled(laborDepartmentChartKey(key), YEARS[index]);
  });
  addCostControlPointIfControlled("laborBottomUp", YEARS[index]);
}

function distributeLaborTotalAcrossKeys(scenario, index, keys, targetTotal) {
  if (!keys.length) return;
  const actualValue = Math.max(0, Number(targetTotal) || 0);
  const currentTotal = laborDepartmentGroupTotalForYear(scenario.costs.laborDepartments, keys, index);
  const baseTotal = laborDepartmentGroupTotalForYear(baseCosts.laborDepartments, keys, index);
  keys.forEach(key => {
    const currentValue = Number(scenario.costs.laborDepartments[key][index] || 0);
    const baseValue = Number(baseCosts.laborDepartments[key][index] || 0);
    const share = currentTotal > 0
      ? currentValue / currentTotal
      : baseTotal > 0 ? baseValue / baseTotal : 1 / keys.length;
    scenario.costs.laborDepartments[key][index] = actualValue * share;
    addCostControlPointIfControlled(laborDepartmentChartKey(key), YEARS[index]);
  });
}

function setLaborAllocationControlValue(scenario, index, selectedKeys, targetTotal, { lockTop = false } = {}) {
  normalizeCostPlan(scenario);
  const selected = selectedKeys.filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
  if (!selected.length) return;
  const unselected = LABOR_DEPARTMENT_KEYS.filter(key => !selected.includes(key));
  const existingBottomUpTotal = laborDepartmentTotalForYear(scenario.costs.laborDepartments, index);
  let selectedTarget = Math.max(0, Number(targetTotal) || 0);
  if (lockTop) {
    selectedTarget = Math.min(selectedTarget, existingBottomUpTotal);
  }
  distributeLaborTotalAcrossKeys(scenario, index, selected, selectedTarget);
  if (lockTop) {
    distributeLaborTotalAcrossKeys(scenario, index, unselected, Math.max(0, existingBottomUpTotal - selectedTarget));
  }
  addCostControlPointIfControlled("laborBottomUp", YEARS[index]);
  addCostControlPointIfControlled("laborAllocationControl", YEARS[index]);
}

function nonLaborTopDownMatchesBottomUp(scenario = activeScenario()) {
  normalizeCostPlan(scenario);
  return YEARS.every((year, index) => {
    if (!editableCostYear(year)) return true;
    const topDown = Number(scenario.costs.nonLabor[index] || 0);
    const bottomUp = nonLaborCategoryTotalForYear(scenario.costs.nonLaborCategories, index);
    return Math.abs(topDown - bottomUp) < 1;
  });
}

function setNonLaborCategoryTotalForYear(scenario, index, targetTotal) {
  normalizeCostPlan(scenario);
  setNonLaborCategoryBottomUpTotalForYear(scenario, index, targetTotal);
  scenario.costs.nonLabor[index] = Math.max(0, Number(targetTotal) || 0);
}

function distributeNonLaborTotalAcrossKeys(scenario, index, keys, targetTotal) {
  if (!keys.length) return;
  const actualValue = Math.max(0, Number(targetTotal) || 0);
  const currentTotal = nonLaborCategoryGroupTotalForYear(scenario.costs.nonLaborCategories, keys, index);
  const baseTotal = nonLaborCategoryGroupTotalForYear(baseCosts.nonLaborCategories, keys, index);
  keys.forEach(key => {
    const currentValue = Number(scenario.costs.nonLaborCategories[key][index] || 0);
    const baseValue = Number(baseCosts.nonLaborCategories[key][index] || 0);
    const share = currentTotal > 0
      ? currentValue / currentTotal
      : baseTotal > 0 ? baseValue / baseTotal : 1 / keys.length;
    scenario.costs.nonLaborCategories[key][index] = actualValue * share;
    addCostControlPointIfControlled(nonLaborCategoryChartKey(key), YEARS[index]);
  });
}

function setNonLaborCategoryBottomUpTotalForYear(scenario, index, targetTotal, { lockOthers = false, selectedKeys = [] } = {}) {
  normalizeCostPlan(scenario);
  const actualValue = Math.max(0, Number(targetTotal) || 0);
  const selected = selectedKeys.filter(key => NON_LABOR_CATEGORY_KEYS.includes(key));
  if (lockOthers && selected.length) {
    const unselected = NON_LABOR_CATEGORY_KEYS.filter(key => !selected.includes(key));
    const unselectedTotal = nonLaborCategoryGroupTotalForYear(scenario.costs.nonLaborCategories, unselected, index);
    distributeNonLaborTotalAcrossKeys(scenario, index, selected, Math.max(0, actualValue - unselectedTotal));
    addCostControlPointIfControlled("nonLaborBottomUp", YEARS[index]);
    addCostControlPointIfControlled("nonLaborAllocationControl", YEARS[index]);
    return;
  }
  const currentTotal = nonLaborCategoryTotalForYear(scenario.costs.nonLaborCategories, index);
  const baseTotal = Number(baseCosts.nonLabor[index] || 0);
  NON_LABOR_CATEGORY_KEYS.forEach(key => {
    const currentValue = Number(scenario.costs.nonLaborCategories[key][index] || 0);
    const baseValue = Number(baseCosts.nonLaborCategories[key][index] || 0);
    const share = currentTotal > 0
      ? currentValue / currentTotal
      : baseTotal > 0 ? baseValue / baseTotal : 1 / NON_LABOR_CATEGORY_KEYS.length;
    scenario.costs.nonLaborCategories[key][index] = actualValue * share;
    addCostControlPointIfControlled(nonLaborCategoryChartKey(key), YEARS[index]);
  });
  addCostControlPointIfControlled("nonLaborBottomUp", YEARS[index]);
}

function setNonLaborAllocationControlValue(scenario, index, selectedKeys, targetTotal, { lockTop = false } = {}) {
  normalizeCostPlan(scenario);
  const selected = selectedKeys.filter(key => NON_LABOR_CATEGORY_KEYS.includes(key));
  if (!selected.length) return;
  const unselected = NON_LABOR_CATEGORY_KEYS.filter(key => !selected.includes(key));
  const existingBottomUpTotal = nonLaborCategoryTotalForYear(scenario.costs.nonLaborCategories, index);
  let selectedTarget = Math.max(0, Number(targetTotal) || 0);
  if (lockTop) {
    selectedTarget = Math.min(selectedTarget, existingBottomUpTotal);
  }
  distributeNonLaborTotalAcrossKeys(scenario, index, selected, selectedTarget);
  if (lockTop) {
    distributeNonLaborTotalAcrossKeys(scenario, index, unselected, Math.max(0, existingBottomUpTotal - selectedTarget));
  }
  addCostControlPointIfControlled("nonLaborBottomUp", YEARS[index]);
  addCostControlPointIfControlled("nonLaborAllocationControl", YEARS[index]);
}

function nonLaborCategoryDepartmentChartKey(categoryKey, departmentKey) {
  return `nonLaborCategoryDepartment:${categoryKey}:${departmentKey}`;
}

function distributeNonLaborCategoryDepartmentAcrossKeys(scenario, index, categoryKey, departmentKeys, targetTotal) {
  normalizeCostPlan(scenario);
  const keys = departmentKeys.filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
  if (!keys.length || !NON_LABOR_CATEGORY_KEYS.includes(categoryKey)) return;
  const actualValue = Math.max(0, Number(targetTotal) || 0);
  const currentTotal = nonLaborCategoryDepartmentGroupTotalForYear(scenario.costs.nonLaborCategoryDepartments, categoryKey, keys, index);
  const baseDepartments = createBaseNonLaborCategoryDepartments(baseCosts.nonLaborCategories);
  const baseTotal = nonLaborCategoryDepartmentGroupTotalForYear(baseDepartments, categoryKey, keys, index);
  keys.forEach(key => {
    const currentValue = Number(scenario.costs.nonLaborCategoryDepartments[categoryKey][key][index] || 0);
    const baseValue = Number(baseDepartments[categoryKey]?.[key]?.[index] || 0);
    const share = currentTotal > 0
      ? currentValue / currentTotal
      : baseTotal > 0 ? baseValue / baseTotal : 1 / keys.length;
    scenario.costs.nonLaborCategoryDepartments[categoryKey][key][index] = actualValue * share;
    addCostControlPointIfControlled(nonLaborCategoryDepartmentChartKey(categoryKey, key), YEARS[index]);
  });
}

function setNonLaborCategoryDepartmentBottomUpTotalForYear(scenario, index, categoryKey, targetTotal, { lockOthers = false, selectedKeys = [] } = {}) {
  normalizeCostPlan(scenario);
  if (!NON_LABOR_CATEGORY_KEYS.includes(categoryKey)) return;
  const selected = selectedKeys.filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
  const actualValue = Math.max(0, Number(targetTotal) || 0);
  if (lockOthers && selected.length) {
    const unselected = LABOR_DEPARTMENT_KEYS.filter(key => !selected.includes(key));
    const unselectedTotal = nonLaborCategoryDepartmentGroupTotalForYear(scenario.costs.nonLaborCategoryDepartments, categoryKey, unselected, index);
    distributeNonLaborCategoryDepartmentAcrossKeys(scenario, index, categoryKey, selected, Math.max(0, actualValue - unselectedTotal));
  } else {
    distributeNonLaborCategoryDepartmentAcrossKeys(scenario, index, categoryKey, LABOR_DEPARTMENT_KEYS, actualValue);
  }
  addCostControlPointIfControlled(`nonLaborCategoryDepartmentBottomUp:${categoryKey}`, YEARS[index]);
  addCostControlPointIfControlled(`nonLaborCategoryDepartmentControl:${categoryKey}`, YEARS[index]);
}

function setNonLaborCategoryDepartmentControlValue(scenario, index, categoryKey, selectedKeys, targetTotal, { lockTop = false } = {}) {
  normalizeCostPlan(scenario);
  const selected = selectedKeys.filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
  if (!selected.length || !NON_LABOR_CATEGORY_KEYS.includes(categoryKey)) return;
  const unselected = LABOR_DEPARTMENT_KEYS.filter(key => !selected.includes(key));
  const existingBottomUpTotal = nonLaborCategoryDepartmentBottomUpValues(scenario, categoryKey)[index];
  let selectedTarget = Math.max(0, Number(targetTotal) || 0);
  if (lockTop) selectedTarget = Math.min(selectedTarget, existingBottomUpTotal);
  distributeNonLaborCategoryDepartmentAcrossKeys(scenario, index, categoryKey, selected, selectedTarget);
  if (lockTop) {
    distributeNonLaborCategoryDepartmentAcrossKeys(scenario, index, categoryKey, unselected, Math.max(0, existingBottomUpTotal - selectedTarget));
  }
  addCostControlPointIfControlled(`nonLaborCategoryDepartmentBottomUp:${categoryKey}`, YEARS[index]);
  addCostControlPointIfControlled(`nonLaborCategoryDepartmentControl:${categoryKey}`, YEARS[index]);
}

function nonLaborCategoryTopMatchesBottomUp(scenario = activeScenario(), categoryKey = selectedNonLaborCategoryKey()) {
  normalizeCostPlan(scenario);
  const topDown = nonLaborCategoryValues(scenario, categoryKey);
  const bottomUp = nonLaborCategoryDepartmentBottomUpValues(scenario, categoryKey);
  return YEARS.every((year, index) => {
    if (!editableCostYear(year)) return true;
    return Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0)) < 1;
  });
}

function costZeroFallbackBase(values, index) {
  let bestDistance = Infinity;
  let bestValue = 0;
  (values || []).forEach((value, valueIndex) => {
    const numericValue = Number(value);
    if (!Number.isFinite(numericValue) || numericValue <= 0) return;
    const distance = Math.abs(valueIndex - index);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestValue = numericValue;
    }
  });
  return bestValue > 0 ? Math.max(1, bestValue / 2) : 100000;
}

function costYoYDenominator(values, index) {
  const previous = Number(values?.[index - 1] || 0);
  const current = Number(values?.[index] || 0);
  if (previous > 0) return previous;
  if (current > 0) return current / 2;
  return costZeroFallbackBase(values, index);
}

function costYoYPercentFromValues(values, index) {
  if (index === 0) return 0;
  const previous = Number(values[index - 1] || 0);
  const current = Number(values[index] || 0);
  if (previous <= 0 && current <= 0) return 0;
  const denominator = costYoYDenominator(values, index);
  return denominator > 0 ? ((current / denominator) - 1) * 100 : 0;
}

function valueFromEditedCostPoint(previousValue, editedValue, chartKey, currentValue = 0, values = [], index = 0) {
  const numericValue = Number(editedValue);
  if (!Number.isFinite(numericValue)) return null;
  if (!isCostChartYoY(chartKey)) return numericValue * 1000000;
  const previous = Math.max(0, Number(previousValue) || 0);
  const current = Math.max(0, Number(currentValue) || 0);
  if (previous <= 0 && current <= 0 && Math.abs(numericValue) < 0.000001) return 0;
  const denominator = previous > 0
    ? previous
    : current > 0 ? current / 2 : costZeroFallbackBase(values, index);
  return Math.max(0, denominator * (1 + numericValue / 100));
}

function isCostLocked(key, scenario = activeScenario()) {
  normalizeCostPlan(scenario);
  return Boolean(scenario.costLocks[key]);
}

function renderCostLockButtons() {
  document.querySelectorAll("[data-cost-lock-key]").forEach(button => {
    const key = button.dataset.costLockKey;
    const locked = isCostLocked(key);
    button.classList.toggle("locked", locked);
    button.textContent = locked ? "Locked" : "Lock";
    button.setAttribute("aria-pressed", locked ? "true" : "false");
  });
}

function renderCostDrilldownView() {
  const isLaborDrilldown = state.costDrilldown === "labor";
  const isLaborFteDrilldown = state.costDrilldown === "laborFte";
  const isNonLaborDrilldown = state.costDrilldown === "nonLabor";
  const isNonLaborCategoryDrilldown = state.costDrilldown === "nonLaborCategory";
  const mainPanels = document.getElementById("cost-main-panels");
  const laborDrilldown = document.getElementById("cost-labor-drilldown");
  const laborFteDrilldown = document.getElementById("cost-labor-fte-drilldown");
  const nonLaborDrilldown = document.getElementById("cost-non-labor-drilldown");
  const nonLaborCategoryDrilldown = document.getElementById("cost-non-labor-category-drilldown");
  const setBottomToTopButton = document.getElementById("set-cost-labor-bottom-to-top");
  const setTopToBottomButton = document.getElementById("set-cost-labor-top-to-bottom");
  const setLaborFteBottomToTopButton = document.getElementById("set-cost-labor-fte-bottom-to-top");
  const setLaborFteTopToBottomButton = document.getElementById("set-cost-labor-fte-top-to-bottom");
  const setNonLaborBottomToTopButton = document.getElementById("set-cost-non-labor-bottom-to-top");
  const setNonLaborTopToBottomButton = document.getElementById("set-cost-non-labor-top-to-bottom");
  const setNonLaborCategoryBottomToTopButton = document.getElementById("set-cost-non-labor-category-bottom-to-top");
  const setNonLaborCategoryTopToBottomButton = document.getElementById("set-cost-non-labor-category-top-to-bottom");
  const anyDrilldown = isLaborDrilldown || isLaborFteDrilldown || isNonLaborDrilldown || isNonLaborCategoryDrilldown;
  if (mainPanels) mainPanels.hidden = anyDrilldown;
  if (laborDrilldown) laborDrilldown.hidden = !isLaborDrilldown;
  if (laborFteDrilldown) laborFteDrilldown.hidden = !isLaborFteDrilldown;
  if (nonLaborDrilldown) nonLaborDrilldown.hidden = !isNonLaborDrilldown;
  if (nonLaborCategoryDrilldown) nonLaborCategoryDrilldown.hidden = !isNonLaborCategoryDrilldown;
  const matched = laborTopDownMatchesBottomUp();
  if (setBottomToTopButton) setBottomToTopButton.disabled = matched;
  if (setTopToBottomButton) setTopToBottomButton.disabled = matched;
  const laborFteMatched = laborFteTopMatchesBottomUp();
  if (setLaborFteBottomToTopButton) setLaborFteBottomToTopButton.disabled = laborFteMatched;
  if (setLaborFteTopToBottomButton) setLaborFteTopToBottomButton.disabled = laborFteMatched;
  const nonLaborMatched = nonLaborTopDownMatchesBottomUp();
  if (setNonLaborBottomToTopButton) setNonLaborBottomToTopButton.disabled = nonLaborMatched;
  if (setNonLaborTopToBottomButton) setNonLaborTopToBottomButton.disabled = nonLaborMatched;
  const nonLaborCategoryMatched = nonLaborCategoryTopMatchesBottomUp();
  if (setNonLaborCategoryBottomToTopButton) setNonLaborCategoryBottomToTopButton.disabled = nonLaborCategoryMatched;
  if (setNonLaborCategoryTopToBottomButton) setNonLaborCategoryTopToBottomButton.disabled = nonLaborCategoryMatched;
  renderCostChartModeButtons();
  renderCostChartLineModeButtons();
  renderLaborAllocationControls();
  renderLaborFteControls();
  renderNonLaborAllocationControls();
  renderNonLaborCategoryControls();
  refreshHeaderKpis();
}

function resizeCostCharts() {
  setTimeout(() => {
    Object.values(state.costCharts).forEach(chart => chart.resize());
    state.outputCharts.costLaborDepartmentMix?.resize();
    state.outputCharts.costNonLaborCategoryMix?.resize();
    state.outputCharts.costNonLaborCategoryDepartmentMix?.resize();
  }, 0);
}

function toggleCostLock(key) {
  if (!COST_LOCK_KEYS.includes(key)) return;
  pushUndoSnapshot();
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const nextLocked = !scenario.costLocks[key];
  scenario.costLocks[key] = nextLocked;
  if (nextLocked && key === "labor") {
    scenario.costLocks.nonLabor = false;
  } else if (nextLocked && key === "nonLabor") {
    scenario.costLocks.labor = false;
  }
  saveScenarios();
  renderCostLockButtons();
}

function laborBottomUpInterpolatedDollarValue(year, index) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const chart = state.costCharts.laborDrilldownTopDown;
  if (chart && !isCostChartYoY("laborDrilldownTopDown")) {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (Number.isFinite(value)) return Math.max(0, value * 1000000);
  }
  return laborDepartmentTotalForYear(scenario.costs.laborDepartments, index);
}

function setLaborTopToBottom() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (laborTopDownMatchesBottomUp(scenario)) return;
  pushUndoSnapshot();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    const targetLabor = laborBottomUpInterpolatedDollarValue(year, index);
    scenario.costs.labor[index] = targetLabor;
    addCostControlPointIfControlled("labor", year);
    addCostControlPointIfControlled("total", year);
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function setLaborBottomToTop() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (laborTopDownMatchesBottomUp(scenario)) return;
  pushUndoSnapshot();
  const selectedKeys = selectedLaborAllocationKeys();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    setLaborDepartmentBottomUpTotalForYear(scenario, index, Number(scenario.costs.labor[index] || 0), {
      lockOthers: state.laborAllocationOthersLocked,
      selectedKeys,
    });
    addCostControlPointIfControlled("laborBottomUp", year);
    addCostControlPointIfControlled("laborAllocationControl", year);
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function laborFteBuildInterpolatedDollarValue(year, index) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const chart = state.costCharts.laborFteSelectedCost;
  if (chart) {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (Number.isFinite(value)) return Math.max(0, value * 1000000);
  }
  return laborFteBuildCostValues(scenario, activeLaborFteKeys())[index];
}

function setLaborFteTopToBottom() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (laborFteTopMatchesBottomUp(scenario)) return;
  pushUndoSnapshot();
  const keys = activeLaborFteKeys();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    setLaborDepartmentAllocationToFteBuildForYear(scenario, index, keys, laborFteBuildInterpolatedDollarValue(year, index));
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function setLaborFteBottomToTop() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (laborFteTopMatchesBottomUp(scenario)) return;
  pushUndoSnapshot();
  const keys = activeLaborFteKeys();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    setLaborFteBuildCostForYear(scenario, index, keys, laborFteSelectedCostValues(scenario, keys)[index]);
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function nonLaborBottomUpInterpolatedDollarValue(year, index) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const chart = state.costCharts.nonLaborDrilldownTopDown;
  if (chart && !isCostChartYoY("nonLaborDrilldownTopDown")) {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (Number.isFinite(value)) return Math.max(0, value * 1000000);
  }
  return nonLaborCategoryTotalForYear(scenario.costs.nonLaborCategories, index);
}

function setNonLaborTopToBottom() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (nonLaborTopDownMatchesBottomUp(scenario)) return;
  pushUndoSnapshot();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    const targetNonLabor = nonLaborBottomUpInterpolatedDollarValue(year, index);
    scenario.costs.nonLabor[index] = targetNonLabor;
    addCostControlPointIfControlled("nonLabor", year);
    addCostControlPointIfControlled("total", year);
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function setNonLaborBottomToTop() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  if (nonLaborTopDownMatchesBottomUp(scenario)) return;
  pushUndoSnapshot();
  const selectedKeys = selectedNonLaborAllocationKeys();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    setNonLaborCategoryBottomUpTotalForYear(scenario, index, Number(scenario.costs.nonLabor[index] || 0), {
      lockOthers: state.nonLaborAllocationOthersLocked,
      selectedKeys,
    });
    addCostControlPointIfControlled("nonLaborBottomUp", year);
    addCostControlPointIfControlled("nonLaborAllocationControl", year);
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function selectedNonLaborCategoryBottomUpInterpolatedDollarValue(year, index) {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const categoryKey = selectedNonLaborCategoryKey();
  const chart = state.costCharts.nonLaborCategorySelectedSpend;
  if (chart && !isCostChartYoY("nonLaborCategorySelectedSpend")) {
    const value = interpolateNapkinLineValue(chart.lines[0].data, year);
    if (Number.isFinite(value)) return Math.max(0, value * 1000000);
  }
  return nonLaborCategoryDepartmentBottomUpValues(scenario, categoryKey)[index];
}

function setNonLaborCategoryTopToBottom() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const categoryKey = selectedNonLaborCategoryKey();
  if (nonLaborCategoryTopMatchesBottomUp(scenario, categoryKey)) return;
  pushUndoSnapshot();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    scenario.costs.nonLaborCategories[categoryKey][index] = selectedNonLaborCategoryBottomUpInterpolatedDollarValue(year, index);
    addCostControlPointIfControlled(nonLaborCategoryChartKey(categoryKey), year);
    addCostControlPointIfControlled("nonLaborBottomUp", year);
    addCostControlPointIfControlled("nonLaborAllocationControl", year);
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function setNonLaborCategoryBottomToTop() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const categoryKey = selectedNonLaborCategoryKey();
  if (nonLaborCategoryTopMatchesBottomUp(scenario, categoryKey)) return;
  pushUndoSnapshot();
  const selectedKeys = selectedNonLaborCategoryDepartmentKeys();
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    setNonLaborCategoryDepartmentBottomUpTotalForYear(scenario, index, categoryKey, Number(scenario.costs.nonLaborCategories[categoryKey][index] || 0), {
      lockOthers: state.nonLaborCategoryOthersLocked,
      selectedKeys,
    });
  });
  saveScenarios();
  syncCostCharts();
  renderAll();
}

function renderLaborAllocationControls() {
  const list = document.getElementById("labor-allocation-line-list");
  const mode = costChartMode("laborAllocationControl");
  const isPctTotal = mode === "pct";
  const lockButton = document.getElementById("toggle-labor-allocation-top-lock");
  if (lockButton) {
    lockButton.classList.toggle("locked", state.laborAllocationTopLocked);
    lockButton.textContent = state.laborAllocationTopLocked ? "Top Locked" : "Lock Top";
    lockButton.setAttribute("aria-pressed", state.laborAllocationTopLocked ? "true" : "false");
  }
  const lockOthersButton = document.getElementById("toggle-labor-allocation-others-lock");
  if (lockOthersButton) {
    lockOthersButton.classList.toggle("locked", state.laborAllocationOthersLocked);
    lockOthersButton.textContent = state.laborAllocationOthersLocked ? "Others Locked" : "Lock Others";
    lockOthersButton.setAttribute("aria-pressed", state.laborAllocationOthersLocked ? "true" : "false");
    lockOthersButton.hidden = isPctTotal;
  }
  if (!list) return;
  const selected = new Set(selectedLaborAllocationKeys());
  const chart = document.getElementById("cost-labor-allocation-control-chart");
  const emptyState = document.getElementById("cost-labor-allocation-empty");
  const pctChart = document.getElementById("cost-labor-allocation-pct-chart");
  const pctEmptyState = document.getElementById("cost-labor-allocation-pct-empty");
  const axisControls = document.querySelector("[data-napkin-y-axis-controls='laborAllocationControl']");
  const hasSelection = selected.size > 0;
  const hasPctSelection = selected.size > 1;
  if (chart) chart.hidden = isPctTotal || !hasSelection;
  if (emptyState) emptyState.hidden = isPctTotal || hasSelection;
  if (pctChart) pctChart.hidden = !isPctTotal || !hasPctSelection;
  if (pctEmptyState) pctEmptyState.hidden = !isPctTotal || hasPctSelection;
  if (axisControls) axisControls.hidden = isPctTotal;
  list.innerHTML = LABOR_DEPARTMENT_KEYS.map(key => {
    const meta = laborDepartmentMeta[key];
    const checked = selected.has(key) ? " checked" : "";
    return `
      <label class="labor-allocation-chip">
        <input type="checkbox" value="${key}"${checked}>
        <span style="--chip-color: ${meta.color}">${displayLabel(meta.label)}</span>
      </label>
    `;
  }).join("");
}

function renderLaborFteControls() {
  const list = document.getElementById("labor-fte-line-list");
  const label = document.getElementById("labor-fte-selection-label");
  if (label) label.textContent = laborFteFilterLabel();
  const selectedCostPctMode = costChartMode("laborFteSelectedCost") === "pct";
  const fteCountPctMode = costChartMode("laborFteCount") === "pct";
  const activeKeys = activeLaborFteKeys();
  const hasPctSelection = activeKeys.length > 1;
  const selectedCostChart = document.getElementById("cost-labor-fte-selected-cost-chart");
  const selectedCostPctChart = document.getElementById("cost-labor-fte-selected-cost-pct-chart");
  const selectedCostPctEmpty = document.getElementById("cost-labor-fte-selected-cost-pct-empty");
  const selectedCostAxis = document.querySelector("[data-napkin-y-axis-controls='laborFteSelectedCost']");
  if (selectedCostChart) selectedCostChart.hidden = selectedCostPctMode;
  if (selectedCostPctChart) selectedCostPctChart.hidden = !selectedCostPctMode || !hasPctSelection;
  if (selectedCostPctEmpty) selectedCostPctEmpty.hidden = !selectedCostPctMode || hasPctSelection;
  if (selectedCostAxis) selectedCostAxis.hidden = selectedCostPctMode;
  const fteCountChart = document.getElementById("cost-labor-fte-count-chart");
  const fteCountPctChart = document.getElementById("cost-labor-fte-count-pct-chart");
  const fteCountPctEmpty = document.getElementById("cost-labor-fte-count-pct-empty");
  const fteCountAxis = document.querySelector("[data-napkin-y-axis-controls='laborFteCount']");
  if (fteCountChart) fteCountChart.hidden = fteCountPctMode;
  if (fteCountPctChart) fteCountPctChart.hidden = !fteCountPctMode || !hasPctSelection;
  if (fteCountPctEmpty) fteCountPctEmpty.hidden = !fteCountPctMode || hasPctSelection;
  if (fteCountAxis) fteCountAxis.hidden = fteCountPctMode;
  renderCostChartLineModeButtons();
  const topLockButton = document.getElementById("toggle-labor-fte-top-lock");
  if (topLockButton) {
    topLockButton.classList.toggle("locked", state.laborFteTopLocked);
    topLockButton.textContent = state.laborFteTopLocked ? "Locked" : "Lock";
    topLockButton.setAttribute("aria-pressed", state.laborFteTopLocked ? "true" : "false");
  }
  document.querySelectorAll("[data-labor-fte-lock]").forEach(button => {
    const lockKey = button.dataset.laborFteLock;
    const locked = state.laborFteLockedDriver === lockKey;
    button.classList.toggle("locked", locked);
    button.textContent = locked ? "Locked" : "Lock";
    button.setAttribute("aria-pressed", locked ? "true" : "false");
  });
  if (!list) return;
  const selected = new Set(selectedLaborFteKeys());
  list.innerHTML = LABOR_DEPARTMENT_KEYS.map(key => {
    const meta = laborDepartmentMeta[key];
    const checked = selected.has(key) ? " checked" : "";
    return `
      <label class="labor-allocation-chip">
        <input type="checkbox" value="${key}"${checked}>
        <span style="--chip-color: ${meta.color}">${displayLabel(meta.label)}</span>
      </label>
    `;
  }).join("");
}

function renderNonLaborAllocationControls() {
  const list = document.getElementById("non-labor-allocation-line-list");
  const lockButton = document.getElementById("toggle-non-labor-allocation-top-lock");
  if (lockButton) {
    lockButton.classList.toggle("locked", state.nonLaborAllocationTopLocked);
    lockButton.textContent = state.nonLaborAllocationTopLocked ? "Top Locked" : "Lock Top";
    lockButton.setAttribute("aria-pressed", state.nonLaborAllocationTopLocked ? "true" : "false");
  }
  const lockOthersButton = document.getElementById("toggle-non-labor-allocation-others-lock");
  if (lockOthersButton) {
    lockOthersButton.classList.toggle("locked", state.nonLaborAllocationOthersLocked);
    lockOthersButton.textContent = state.nonLaborAllocationOthersLocked ? "Others Locked" : "Lock Others";
    lockOthersButton.setAttribute("aria-pressed", state.nonLaborAllocationOthersLocked ? "true" : "false");
  }
  if (!list) return;
  const selected = new Set(selectedNonLaborAllocationKeys());
  const chart = document.getElementById("cost-non-labor-allocation-control-chart");
  const emptyState = document.getElementById("cost-non-labor-allocation-empty");
  const hasSelection = selected.size > 0;
  if (chart) chart.hidden = !hasSelection;
  if (emptyState) emptyState.hidden = hasSelection;
  list.innerHTML = NON_LABOR_CATEGORY_KEYS.map(key => {
    const meta = nonLaborCategoryMeta[key];
    const checked = selected.has(key) ? " checked" : "";
    return `
      <label class="labor-allocation-chip">
        <input type="checkbox" value="${key}"${checked}>
        <span style="--chip-color: ${meta.color}">${displayLabel(meta.label)}</span>
      </label>
    `;
  }).join("");
}

function renderNonLaborCategoryControls() {
  const categoryKey = selectedNonLaborCategoryKey();
  const categoryMeta = nonLaborCategoryMeta[categoryKey];
  const picker = document.getElementById("non-labor-category-picker");
  if (picker) {
    picker.innerHTML = NON_LABOR_CATEGORY_KEYS.map(key => {
      const meta = nonLaborCategoryMeta[key];
      const checked = key === categoryKey ? " checked" : "";
      return `
        <label class="labor-allocation-chip">
          <input type="radio" name="non-labor-category-picker" value="${key}"${checked}>
          <span style="--chip-color: ${meta.color}">${displayLabel(meta.label)}</span>
        </label>
      `;
    }).join("");
  }
  const title = document.getElementById("non-labor-category-selected-title");
  if (title) title.textContent = `${displayLabel(categoryMeta.label)} Spend`;
  const label = document.getElementById("non-labor-category-selection-label");
  if (label) label.textContent = displayLabel(categoryMeta.label);
  const topLockButton = document.getElementById("toggle-non-labor-category-top-lock");
  if (topLockButton) {
    topLockButton.classList.toggle("locked", state.nonLaborCategoryTopLocked);
    topLockButton.textContent = state.nonLaborCategoryTopLocked ? "Top Locked" : "Lock Top";
    topLockButton.setAttribute("aria-pressed", state.nonLaborCategoryTopLocked ? "true" : "false");
  }
  const othersLockButton = document.getElementById("toggle-non-labor-category-others-lock");
  if (othersLockButton) {
    othersLockButton.classList.toggle("locked", state.nonLaborCategoryOthersLocked);
    othersLockButton.textContent = state.nonLaborCategoryOthersLocked ? "Others Locked" : "Lock Others";
    othersLockButton.setAttribute("aria-pressed", state.nonLaborCategoryOthersLocked ? "true" : "false");
  }
  const list = document.getElementById("non-labor-category-department-line-list");
  if (!list) return;
  const selected = new Set(selectedNonLaborCategoryDepartmentKeys());
  const hasSelection = selected.size > 0;
  const chart = document.getElementById("cost-non-labor-category-department-control-chart");
  const emptyState = document.getElementById("cost-non-labor-category-department-empty");
  if (chart) chart.hidden = !hasSelection;
  if (emptyState) emptyState.hidden = hasSelection;
  list.innerHTML = LABOR_DEPARTMENT_KEYS.map(key => {
    const meta = laborDepartmentMeta[key];
    const checked = selected.has(key) ? " checked" : "";
    return `
      <label class="labor-allocation-chip">
        <input type="checkbox" value="${key}"${checked}>
        <span style="--chip-color: ${meta.color}">${displayLabel(meta.label)}</span>
      </label>
    `;
  }).join("");
}

function renderCostChartModeButtons() {
  document.querySelectorAll("[data-cost-chart-mode]").forEach(group => {
    const chartKey = group.dataset.costChartMode;
    const mode = costChartMode(chartKey);
    group.querySelectorAll("[data-cost-chart-mode-value]").forEach(button => {
      const active = button.dataset.costChartModeValue === mode;
      button.classList.toggle("active", active);
      button.setAttribute("aria-pressed", active ? "true" : "false");
    });
  });
}

function renderCostChartLineModeButtons() {
  document.querySelectorAll("[data-cost-line-mode]").forEach(group => {
    const chartKey = group.dataset.costLineMode;
    const visible = costChartMode(chartKey) === "value";
    group.hidden = !visible;
    const mode = costChartLineMode(chartKey);
    group.querySelectorAll("[data-cost-line-mode-value]").forEach(button => {
      const active = button.dataset.costLineModeValue === mode;
      button.classList.toggle("active", active);
      button.setAttribute("aria-pressed", active ? "true" : "false");
    });
  });
}

function setCostChartMode(chartKey, mode) {
  if (!["value", "yoy", "pct"].includes(mode)) return;
  state.costChartModes[chartKey] = mode;
  syncCostCharts();
  resizeCostCharts();
}

function setCostChartLineMode(chartKey, mode) {
  if (!["grouped", "separated"].includes(mode)) return;
  state.costChartLineModes[chartKey] = mode;
  syncCostCharts();
  resizeCostCharts();
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
  syncCostCharts();
  syncRevUnitCharts();
  syncRevenueDrilldownCharts();
  syncProfitCharts();
  syncOtherRevenueCharts();
  syncStrRevenueCharts();
  syncGrowthRevenueCharts();
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

function editableCostYear(year) {
  return year >= FIRST_COST_EDIT_YEAR;
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

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function rawDataSources() {
  const planRows = YEARS.map((year, index) => ({
    year,
    planRevenue: planRevenue[index],
  }));
  const driverRows = YEARS.map((year, index) => ({
    year,
    retention: baseDrivers.retention[index],
    newCustomers: baseDrivers.newCustomers[index],
    profilesReturning: baseDrivers.profilesReturning[index],
    profilesNew: baseDrivers.profilesNew[index],
    revReturningProfile: baseDrivers.revReturningProfile[index],
    revNewProfile: baseDrivers.revNewProfile[index],
  }));
  const returningRows = historicalReturningCustomers.map((customers, index) => ({
    year: YEARS[index],
    returningCustomers: customers,
  }));
  const cohortRows = Object.entries(baseNewCustomerCohortCounts)
    .flatMap(([year, cohorts]) => Object.entries(cohorts).map(([cohortYear, newCustomers]) => ({
      year: Number(year),
      cohortYear: Number(cohortYear),
      newCustomers,
    })))
    .sort((left, right) => left.year - right.year || left.cohortYear - right.cohortYear);
  const revUnitSourceRows = REV_UNIT_YEARS.map(year => ({
    year,
    newUnits: baseRevUnitNewUnits[year] ?? null,
    newProperties: baseRevUnitNewProperties[year] ?? null,
    hhpRevenue: baseRevUnitRevenue[year] ?? null,
    unaffiliatedRevenue: baseUnaffiliatedRevenue[year] ?? null,
  }));
  const revUnitRpuRows = Object.entries(baseRevUnitRpu)
    .flatMap(([year, cohorts]) => Object.entries(cohorts).map(([cohortYear, revPerUnit]) => ({
      year: Number(year),
      cohortYear: Number(cohortYear),
      cohortAge: Number(year) - Number(cohortYear),
      revPerUnit,
    })))
    .sort((left, right) => left.year - right.year || left.cohortYear - right.cohortYear);
  const initiativeTeamRows = INITIATIVE_TEAMS.map((team, index) => ({
    team,
    sortOrder: index + 1,
  }));

  return [
    {
      id: "plan-revenue",
      title: "Plan Revenue",
      category: "Revenue Targets",
      description: "Annual plan revenue used as the target line across the model.",
      columns: [
        { key: "year", label: "Year", format: "year" },
        { key: "planRevenue", label: "Plan Revenue", format: "currency0" },
      ],
      rows: planRows,
    },
    {
      id: "hhp-drivers",
      title: "HHP Driver Baselines",
      category: "Household Pet Profile Revenue",
      description: "Top-level historical and forecast driver assumptions for the HHP model.",
      columns: [
        { key: "year", label: "Year", format: "year" },
        { key: "retention", label: "Retention", format: "percent" },
        { key: "newCustomers", label: "New Customers", format: "integer" },
        { key: "profilesReturning", label: "Profiles / Returning Customer", format: "decimal" },
        { key: "profilesNew", label: "Profiles / New Customer", format: "decimal" },
        { key: "revReturningProfile", label: "Rev / Returning Profile", format: "currency2" },
        { key: "revNewProfile", label: "Rev / New Profile", format: "currency2" },
      ],
      rows: driverRows,
    },
    {
      id: "historical-returning-customers",
      title: "Historical Returning Customers",
      category: "Household Pet Profile Revenue",
      description: "Historical returning customer values used before model-calculated retention begins.",
      columns: [
        { key: "year", label: "Year", format: "year" },
        { key: "returningCustomers", label: "Returning Customers", format: "integer" },
      ],
      rows: returningRows,
    },
    {
      id: "new-customer-cohorts",
      title: "New Customers by Property Cohort",
      category: "New Customer Drilldown",
      description: "Baseline customer counts by calendar year and property cohort year.",
      columns: [
        { key: "year", label: "Year", format: "year" },
        { key: "cohortYear", label: "Property Cohort Year", format: "cohort" },
        { key: "newCustomers", label: "New Customers", format: "integer" },
      ],
      rows: cohortRows,
    },
    {
      id: "rev-unit-source",
      title: "Rev / Unit Source Values",
      category: "Rev Per Unit Plan",
      description: "Starting unit, property, HHP revenue, and unaffiliated revenue values for the rev/unit model.",
      columns: [
        { key: "year", label: "Year", format: "year" },
        { key: "newUnits", label: "New Units", format: "integer" },
        { key: "newProperties", label: "New Properties", format: "integer" },
        { key: "hhpRevenue", label: "HHP Revenue", format: "currency0" },
        { key: "unaffiliatedRevenue", label: "Unaffiliated Revenue", format: "currency2" },
      ],
      rows: revUnitSourceRows,
    },
    {
      id: "rev-unit-rpu",
      title: "Rev / Unit Cohort Breakout",
      category: "Rev Per Unit Plan",
      description: "Revenue per unit by calendar year and unit cohort year.",
      columns: [
        { key: "year", label: "Year", format: "year" },
        { key: "cohortYear", label: "Cohort Year", format: "cohort" },
        { key: "cohortAge", label: "Cohort Age", format: "integer" },
        { key: "revPerUnit", label: "Revenue / Unit", format: "currency2" },
      ],
      rows: revUnitRpuRows,
    },
    {
      id: "initiative-teams",
      title: "Initiative Teams",
      category: "Defense Metadata",
      description: "Available team tags for defense initiatives.",
      columns: [
        { key: "sortOrder", label: "Sort Order", format: "integer" },
        { key: "team", label: "Team", format: "text" },
      ],
      rows: initiativeTeamRows,
    },
  ];
}

function renderRawDataSources() {
  window.RawDataViewer?.render("raw-data-module", rawDataSources());
}

function napkinChartByKey(key) {
  const explicitCharts = {
    revenueDrilldownLtr: state.revenueDrilldownCharts.ltr,
    revenueDrilldownGrowth: state.revenueDrilldownCharts.growth,
    revenueDrilldownStr: state.revenueDrilldownCharts.str,
    revenueDrilldownOther: state.revenueDrilldownCharts.other,
    profitTop: state.profitCharts.profit,
    profitRevenue: state.profitCharts.revenue,
    profitCost: state.profitCharts.cost,
  };
  return explicitCharts[key]
    || state.cohortCharts[key]
    || state.revUnitCharts[key]
    || state.costCharts[key]
    || state.charts[key]
    || null;
}

function stepYAxisLeadingDigit(value, direction) {
  if (!Number.isFinite(value) || value <= 0) return null;
  const place = Math.pow(10, Math.floor(Math.log10(value)));
  const normalized = value / place;
  const lead = direction === "up"
    ? Math.floor(normalized + 0.000001)
    : Math.ceil(normalized - 0.000001);
  if (direction === "up") {
    return lead >= 9 ? 10 * place : (lead + 1) * place;
  }
  if (direction === "down") {
    if (lead > 1) return (lead - 1) * place;
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

function costChartPairs(scenario, key) {
  const pairs = costValues(scenario, key).map((value, index) => [YEARS[index], value / 1000000]);
  return controlledCostPairs(scenario, key, pairs);
}

function laborBottomUpChartPairs(scenario) {
  const pairs = laborBottomUpValues(scenario).map((value, index) => [YEARS[index], value / 1000000]);
  return controlledCostPairs(scenario, "laborBottomUp", pairs);
}

function laborAllocationControlChartPairs(scenario) {
  const pairs = laborAllocationControlValues(scenario).map((value, index) => [YEARS[index], value / 1000000]);
  return controlledCostPairs(scenario, "laborAllocationControl", pairs);
}

function costYoYPairsFromValues(scenario, values, controlKey) {
  const pairs = YEARS.map((year, index) => {
    return [year, costYoYPercentFromValues(values, index)];
  });
  return pairs;
}

function costPairsFromValuesForMode(scenario, values, controlKey, chartKey) {
  if (isCostChartYoY(chartKey)) return costYoYPairsFromValues(scenario, values, controlKey);
  const pairs = values.map((value, index) => [YEARS[index], value / 1000000]);
  return controlledCostPairs(scenario, controlKey, pairs);
}

function makeCostLine(key, scenario) {
  const meta = costMeta[key];
  return {
    name: displayScenarioName(scenario),
    color: meta.color,
    editable: true,
    editDomain: {
      moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: costChartPairs(scenario, key),
  };
}

function makeCostComparisonLine(key, scenario) {
  return {
    name: `Comparison: ${displayScenarioName(scenario)}`,
    color: "#98a2b3",
    editable: false,
    data: costChartPairs(scenario, key),
  };
}

function costChartLines(key) {
  const lines = [makeCostLine(key, activeScenario())];
  const comparison = compareScenario();
  if (comparison) lines.push(makeCostComparisonLine(key, comparison));
  return lines;
}

function formatCostAxisValue(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  const numericValue = Number(value);
  const absValue = Math.abs(numericValue);
  if (absValue === 0) return "$0";
  if (absValue >= 1) {
    const decimals = absValue < 10 ? 1 : 0;
    return `$${trimNumber(numericValue, decimals)}M`;
  }
  if (absValue >= 0.001) {
    const thousands = numericValue * 1000;
    const decimals = Math.abs(thousands) < 10 ? 1 : 0;
    return `$${trimNumber(thousands, decimals)}k`;
  }
  return formatCurrency(numericValue * 1000000, 0);
}

function formatCostTooltipValue(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return formatCurrency(Number(value) * 1000000, 0);
}

function costChartMode(chartKey) {
  const mode = state.costChartModes?.[chartKey];
  return ["yoy", "pct"].includes(mode) ? mode : "value";
}

function costChartLineMode(chartKey) {
  return state.costChartLineModes?.[chartKey] === "separated" ? "separated" : "grouped";
}

function isCostChartSeparated(chartKey) {
  return costChartMode(chartKey) === "value" && costChartLineMode(chartKey) === "separated";
}

function isCostChartYoY(chartKey) {
  return costChartMode(chartKey) === "yoy";
}

function formatCostChartAxisValue(value, chartKey) {
  if (isCostChartYoY(chartKey)) return formatPercentMetric(Number(value), 0);
  return formatCostAxisValue(value);
}

function formatCostChartTooltipValue(value, chartKey) {
  if (isCostChartYoY(chartKey)) return formatPercentMetric(Number(value), 1);
  return formatCostTooltipValue(value);
}

function formatFteAxisValue(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return trimNumber(Number(value), 0);
}

function formatFteTooltipValue(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return trimNumber(Number(value), 1);
}

function formatCostPerFteAxisValue(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return `$${trimNumber(Number(value), 0)}k`;
}

function formatCostPerFteTooltipValue(value) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return formatCurrency(Number(value) * 1000, 0);
}

function costChartYAxisConfig(chartKey, valueMax) {
  if (isCostChartYoY(chartKey)) {
    return {
      min: -100,
      max: 100,
      axisLabel: { formatter: value => formatCostChartAxisValue(value, chartKey) },
    };
  }
  return {
    min: 0,
    max: valueMax,
    axisLabel: { formatter: value => formatCostChartAxisValue(value, chartKey) },
  };
}

function applyCostChartYAxisMode(chart, chartKey, valueMax) {
  if (!chart?.baseOption?.yAxis) return;
  const nextYAxis = {
    ...chart.baseOption.yAxis,
    ...costChartYAxisConfig(chartKey, valueMax),
  };
  chart.baseOption.yAxis = nextYAxis;
  chart.chart.setOption({ yAxis: nextYAxis }, false);
}

function setScenarioCostValue(scenario, key, year, value) {
  const index = YEARS.indexOf(year);
  if (index < 0 || !editableCostYear(year) || !Number.isFinite(value)) return;
  normalizeCostPlan(scenario);
  const actualValue = Math.max(0, value * 1000000);
  const locks = scenario.costLocks || {};
  if (key === "total") {
    const currentLabor = Number(scenario.costs.labor[index] || 0);
    const currentNonLabor = Number(scenario.costs.nonLabor[index] || 0);
    const currentTotal = currentLabor + currentNonLabor;
    if (locks.labor && locks.nonLabor) return;
    if (locks.labor) {
      setNonLaborCategoryTotalForYear(scenario, index, Math.max(0, actualValue - currentLabor));
      addCostControlPointIfControlled("nonLabor", year);
      return;
    }
    if (locks.nonLabor) {
      setLaborDepartmentTotalForYear(scenario, index, Math.max(0, actualValue - currentNonLabor));
      addCostControlPointIfControlled("labor", year);
      return;
    }
    const baseLabor = baseCosts.labor[index] || 0;
    const baseNonLabor = baseCosts.nonLabor[index] || 0;
    const baseTotal = baseLabor + baseNonLabor;
    const laborShare = currentTotal > 0
      ? currentLabor / currentTotal
      : baseTotal > 0 ? baseLabor / baseTotal : 0.5;
    setLaborDepartmentTotalForYear(scenario, index, actualValue * laborShare);
    setNonLaborCategoryTotalForYear(scenario, index, actualValue * (1 - laborShare));
    addCostControlPointIfControlled("labor", year);
    addCostControlPointIfControlled("nonLabor", year);
    return;
  }
  const siblingKey = key === "labor" ? "nonLabor" : "labor";
  const currentTotal = scenario.costs.labor[index] + scenario.costs.nonLabor[index];
  if (key === "labor") {
    setLaborDepartmentTotalForYear(scenario, index, actualValue);
  } else {
    setNonLaborCategoryTotalForYear(scenario, index, actualValue);
  }
  if (locks.total && !locks[siblingKey]) {
    if (siblingKey === "labor") {
      setLaborDepartmentTotalForYear(scenario, index, Math.max(0, currentTotal - actualValue));
    } else {
      setNonLaborCategoryTotalForYear(scenario, index, Math.max(0, currentTotal - actualValue));
    }
    addCostControlPointIfControlled(siblingKey, year);
  } else {
    addCostControlPointIfControlled("total", year);
  }
}

function initCostChart(key) {
  const meta = costMeta[key];
  const chart = new NapkinChart(
    meta.chartId,
    costChartLines(key),
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
        ...costChartYAxisConfig("laborDrilldownTopDown", meta.yMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: formatCostTooltipValue,
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
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatCostChartTooltipValue(Number(rawValue), "laborDrilldownTopDown")}`);
            });
          if (!rows.length && items[0]) {
            const item = items[0];
            const data = Array.isArray(item.data) ? item.data : item.value;
            year = Array.isArray(data) ? data[0] : item.axisValue;
            const rawValue = Array.isArray(data) ? data[1] : item.value;
            rows.push(`${item.marker || ""} ${item.seriesName || displayLabel(meta.label)}: ${formatCostTooltipValue(Number(rawValue))}`);
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
  applyCostChartYAxisMode(chart, "laborDrilldownTopDown", meta.yMax);
  chart._refreshChart();
  styleComparisonSeries(chart);
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (laborFteDependentEditBlocked("laborFteCount")) {
      rejectLaborFteDependentEdit(chart);
      return;
    }
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    YEARS.forEach(year => {
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      setScenarioCostValue(activeScenario(), key, year, value);
    });
    rememberCostControlPoints(key, chart.lines[0].data);
    saveScenarios();
    styleComparisonSeries(chart);
    syncCostCharts({ excludeChart: chart });
  };
  state.costCharts[key] = chart;
}

function laborDepartmentChartKey(key) {
  return `laborDepartment:${key}`;
}

function laborDepartmentChartPairs(scenario, key) {
  const chartKey = laborDepartmentChartKey(key);
  return costPairsFromValuesForMode(scenario, laborDepartmentValues(scenario, key), chartKey, chartKey);
}

function laborDepartmentChartLines(key) {
  const meta = laborDepartmentMeta[key];
  const lines = [{
    name: displayLabel(meta.label),
    color: meta.color,
    editable: true,
    editDomain: {
      moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: laborDepartmentChartPairs(activeScenario(), key),
  }];
  const comparison = compareScenario();
  if (comparison) {
    lines.push({
      name: `Comparison: ${displayScenarioName(comparison)}`,
      color: "#98a2b3",
      editable: false,
      data: laborDepartmentChartPairs(comparison, key),
    });
  }
  return lines;
}

function laborDrilldownTopDownChartLines() {
  return [
    {
      name: "Bottom Up",
      color: costMeta.labor.color,
      editable: true,
      editDomain: {
        moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: costPairsFromValuesForMode(activeScenario(), laborBottomUpValues(activeScenario()), "laborBottomUp", "laborDrilldownTopDown"),
    },
    {
      name: "Top Down",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(activeScenario(), costValues(activeScenario(), "labor"), "labor", "laborDrilldownTopDown"),
    },
  ];
}

function laborAllocationControlChartLines() {
  const keys = selectedLaborAllocationKeys();
  if (!keys.length) {
    return [{
      name: "Select Departments",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(activeScenario(), YEARS.map(() => 0), "laborAllocationControl", "laborAllocationControl"),
    }];
  }
  const label = keys.length === 1
    ? displayLabel(laborDepartmentMeta[keys[0]].label)
    : `${keys.length} Departments`;
  return [{
    name: label,
    color: "#2f6f73",
    editable: true,
    editDomain: {
      moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: costPairsFromValuesForMode(activeScenario(), laborAllocationControlValues(activeScenario()), "laborAllocationControl", "laborAllocationControl"),
  }];
}

function laborAllocationPctStorageKey(keys = selectedLaborAllocationKeys()) {
  return `labor:${keys.join("|")}`;
}

function defaultLaborAllocationPctLines(keys, scenario = activeScenario()) {
  if (keys.length < 2) return [];
  normalizeCostPlan(scenario);
  return keys.slice(0, -1).map((_key, boundaryIndex) => {
    return editableCostYears().map(year => {
      const yearIndex = YEARS.indexOf(year);
      const selectedTotal = laborDepartmentGroupTotalForYear(scenario.costs.laborDepartments, keys, yearIndex);
      const boundaryTotal = keys.slice(0, boundaryIndex + 1).reduce((sum, key) => {
        return sum + Number(scenario.costs.laborDepartments?.[key]?.[yearIndex] || 0);
      }, 0);
      const equalShareBoundary = ((boundaryIndex + 1) / keys.length) * 100;
      const boundary = selectedTotal > 0 ? (boundaryTotal / selectedTotal) * 100 : equalShareBoundary;
      return [year, Math.max(0, Math.min(100, boundary))];
    });
  });
}

function laborAllocationPctLineData(scenario = activeScenario(), keys = selectedLaborAllocationKeys()) {
  normalizeCostPlan(scenario);
  if (keys.length < 2) return [];
  const storageKey = laborAllocationPctStorageKey(keys);
  const expectedCount = keys.length - 1;
  const stored = scenario.costPctTotalLines?.[storageKey];
  if (!Array.isArray(stored) || stored.length !== expectedCount) {
    scenario.costPctTotalLines[storageKey] = defaultLaborAllocationPctLines(keys, scenario);
  }
  return scenario.costPctTotalLines[storageKey];
}

function laborAllocationPctLineObjects(scenario = activeScenario()) {
  const keys = selectedLaborAllocationKeys();
  const lineData = laborAllocationPctLineData(scenario, keys);
  return lineData
    .map((data, index) => {
      const key = keys[index];
      const meta = laborDepartmentMeta[key];
      return {
        name: displayLabel(meta.label),
        color: meta.color,
        bandIndex: index,
        editable: true,
        editDomain: {
          moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
          addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
          deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        },
        data: clone(data),
      };
    })
    .reverse();
}

function laborAllocationPctSharesForYear(scenario, keys, year) {
  if (keys.length <= 1) return keys.length ? [1] : [];
  const boundaries = laborAllocationPctLineData(scenario, keys)
    .map(data => interpolateNapkinLineValue(data, year))
    .map(value => Number.isFinite(value) ? Math.max(0, Math.min(100, value)) : 0);
  const shares = [];
  let previous = 0;
  boundaries.forEach(rawBoundary => {
    const boundary = Math.max(previous, rawBoundary);
    shares.push(Math.max(0, boundary - previous) / 100);
    previous = boundary;
  });
  shares.push(Math.max(0, 100 - previous) / 100);
  return shares;
}

function applyLaborAllocationPctShares() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const keys = selectedLaborAllocationKeys();
  if (keys.length < 2) return;
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    const selectedTotal = laborDepartmentGroupTotalForYear(scenario.costs.laborDepartments, keys, index);
    const shares = laborAllocationPctSharesForYear(scenario, keys, year);
    keys.forEach((key, keyIndex) => {
      scenario.costs.laborDepartments[key][index] = selectedTotal * (shares[keyIndex] || 0);
      addCostControlPointIfControlled(laborDepartmentChartKey(key), year);
    });
    addCostControlPointIfControlled("laborBottomUp", year);
    addCostControlPointIfControlled("laborAllocationControl", year);
  });
}

function laborFtePctStorageKey(chartKey, keys = activeLaborFteKeys()) {
  return `${chartKey}:${keys.join("|")}`;
}

function laborFtePctMetricValue(scenario, chartKey, key, index) {
  normalizeCostPlan(scenario);
  if (chartKey === "laborFteCount") {
    return Number(scenario.costs.laborFte?.[key]?.[index] || 0);
  }
  return Number(scenario.costs.laborDepartments?.[key]?.[index] || 0);
}

function laborFtePctMetricTotal(scenario, chartKey, keys, index) {
  return keys.reduce((sum, key) => sum + laborFtePctMetricValue(scenario, chartKey, key, index), 0);
}

function defaultLaborFtePctLines(chartKey, keys, scenario = activeScenario()) {
  if (keys.length < 2) return [];
  normalizeCostPlan(scenario);
  return keys.slice(0, -1).map((_key, boundaryIndex) => {
    return editableCostYears().map(year => {
      const yearIndex = YEARS.indexOf(year);
      const selectedTotal = laborFtePctMetricTotal(scenario, chartKey, keys, yearIndex);
      const boundaryTotal = keys.slice(0, boundaryIndex + 1).reduce((sum, key) => {
        return sum + laborFtePctMetricValue(scenario, chartKey, key, yearIndex);
      }, 0);
      const equalShareBoundary = ((boundaryIndex + 1) / keys.length) * 100;
      const boundary = selectedTotal > 0 ? (boundaryTotal / selectedTotal) * 100 : equalShareBoundary;
      return [year, Math.max(0, Math.min(100, boundary))];
    });
  });
}

function laborFtePctLineData(chartKey, scenario = activeScenario(), keys = activeLaborFteKeys()) {
  normalizeCostPlan(scenario);
  if (keys.length < 2) return [];
  const storageKey = laborFtePctStorageKey(chartKey, keys);
  const expectedCount = keys.length - 1;
  const stored = scenario.costPctTotalLines?.[storageKey];
  if (!Array.isArray(stored) || stored.length !== expectedCount) {
    scenario.costPctTotalLines[storageKey] = defaultLaborFtePctLines(chartKey, keys, scenario);
  }
  return scenario.costPctTotalLines[storageKey];
}

function laborFtePctLineObjects(chartKey, scenario = activeScenario()) {
  const keys = activeLaborFteKeys();
  const lineData = laborFtePctLineData(chartKey, scenario, keys);
  return lineData
    .map((data, index) => {
      const key = keys[index];
      const meta = laborDepartmentMeta[key];
      return {
        name: displayLabel(meta.label),
        color: meta.color,
        bandIndex: index,
        editable: true,
        editDomain: {
          moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
          addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
          deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        },
        data: clone(data),
      };
    })
    .reverse();
}

function laborFtePctSharesForYear(chartKey, scenario, keys, year) {
  if (keys.length <= 1) return keys.length ? [1] : [];
  const boundaries = laborFtePctLineData(chartKey, scenario, keys)
    .map(data => interpolateNapkinLineValue(data, year))
    .map(value => Number.isFinite(value) ? Math.max(0, Math.min(100, value)) : 0);
  const shares = [];
  let previous = 0;
  boundaries.forEach(rawBoundary => {
    const boundary = Math.max(previous, rawBoundary);
    shares.push(Math.max(0, boundary - previous) / 100);
    previous = boundary;
  });
  shares.push(Math.max(0, 100 - previous) / 100);
  return shares;
}

function applyLaborFteSelectedCostPctShares() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const keys = activeLaborFteKeys();
  if (keys.length < 2) return;
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    const selectedTotal = laborFtePctMetricTotal(scenario, "laborFteSelectedCost", keys, index);
    const shares = laborFtePctSharesForYear("laborFteSelectedCost", scenario, keys, year);
    keys.forEach((key, keyIndex) => {
      scenario.costs.laborDepartments[key][index] = selectedTotal * (shares[keyIndex] || 0);
      addCostControlPointIfControlled(laborDepartmentChartKey(key), year);
    });
    addCostControlPointIfControlled("laborFteSelectedCost", year);
    addCostControlPointIfControlled("laborBottomUp", year);
    addCostControlPointIfControlled("laborAllocationControl", year);
  });
  syncLaborTotalsFromDepartments(scenario);
}

function applyLaborFteCountPctShares() {
  const scenario = activeScenario();
  normalizeCostPlan(scenario);
  const keys = activeLaborFteKeys();
  if (keys.length < 2) return;
  YEARS.forEach((year, index) => {
    if (!editableCostYear(year)) return;
    const lockedBuild = laborFteBuildCostValues(scenario, keys)[index];
    const selectedTotal = laborFtePctMetricTotal(scenario, "laborFteCount", keys, index);
    const shares = laborFtePctSharesForYear("laborFteCount", scenario, keys, year);
    keys.forEach((key, keyIndex) => {
      scenario.costs.laborFte[key][index] = selectedTotal * (shares[keyIndex] || 0);
    });
    if (state.laborFteTopLocked) {
      adjustLaborCostPerFteToBuildCost(scenario, index, keys, lockedBuild);
      addCostControlPointIfControlled("laborFteCostPerFte", year);
      addCostControlPointIfControlled("laborFteBuildCost", year);
    }
    addCostControlPointIfControlled("laborFteCount", year);
  });
}

function laborFteMetricPairs(values, controlKey, scale = 1) {
  const pairs = values.map((value, index) => [YEARS[index], value / scale]);
  return controlledCostPairs(activeScenario(), controlKey, pairs);
}

function laborFteEditableDomain() {
  return {
    moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
  };
}

function laborFteDepartmentBuildValues(scenario, key) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => {
    return Number(scenario.costs.laborFte?.[key]?.[index] || 0)
      * Number(scenario.costs.laborCostPerFte?.[key]?.[index] || 0);
  });
}

function laborFteDepartmentMetricPairs(scenario, values, controlKey, scale = 1) {
  const pairs = values.map((value, index) => [YEARS[index], value / scale]);
  return controlledCostPairs(scenario, controlKey, pairs);
}

function laborFteSelectedCostChartLines() {
  normalizeCostPlan(activeScenario());
  const keys = activeLaborFteKeys();
  if (isCostChartSeparated("laborFteSelectedCost")) {
    return [
      ...keys.map(key => {
        const meta = laborDepartmentMeta[key];
        return {
          name: displayLabel(meta.label),
          color: meta.color,
          departmentKey: key,
          editable: true,
          editDomain: laborFteEditableDomain(),
          data: laborFteDepartmentMetricPairs(
            activeScenario(),
            laborFteDepartmentBuildValues(activeScenario(), key),
            `laborFteBuildCost:${key}`,
            1000000
          ),
        };
      }),
      {
        name: "Department Allocation",
        color: "#98a2b3",
        editable: false,
        data: costPairsFromValuesForMode(activeScenario(), laborFteSelectedCostValues(activeScenario(), keys), "laborFteSelectedCost", "laborFteSelectedCost"),
      },
    ];
  }
  return [
    {
      name: "FTE Build",
      color: costMeta.labor.color,
      editable: true,
      editDomain: laborFteEditableDomain(),
      data: costPairsFromValuesForMode(activeScenario(), laborFteBuildCostValues(activeScenario(), keys), "laborFteBuildCost", "laborFteSelectedCost"),
    },
    {
      name: "Department Allocation",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(activeScenario(), laborFteSelectedCostValues(activeScenario(), keys), "laborFteSelectedCost", "laborFteSelectedCost"),
    },
  ];
}

function laborFteCountChartLines() {
  normalizeCostPlan(activeScenario());
  const keys = activeLaborFteKeys();
  if (isCostChartSeparated("laborFteCount")) {
    return keys.map(key => {
      const meta = laborDepartmentMeta[key];
      return {
        name: displayLabel(meta.label),
        color: meta.color,
        departmentKey: key,
        editable: true,
        editDomain: laborFteEditableDomain(),
        data: laborFteDepartmentMetricPairs(
          activeScenario(),
          YEARS.map((_year, index) => Number(activeScenario().costs.laborFte?.[key]?.[index] || 0)),
          `laborFteCount:${key}`
        ),
      };
    });
  }
  return [{
    name: `${laborFteFilterLabel()} FTE / Equivalent`,
    color: "#2f6f73",
    editable: true,
    editDomain: laborFteEditableDomain(),
    data: laborFteMetricPairs(laborFteCountValues(activeScenario(), keys), "laborFteCount"),
  }];
}

function laborFteCostPerFteChartLines() {
  normalizeCostPlan(activeScenario());
  const keys = activeLaborFteKeys();
  if (isCostChartSeparated("laborFteCostPerFte")) {
    return keys.map(key => {
      const meta = laborDepartmentMeta[key];
      return {
        name: displayLabel(meta.label),
        color: meta.color,
        departmentKey: key,
        editable: true,
        editDomain: laborFteEditableDomain(),
        data: laborFteDepartmentMetricPairs(
          activeScenario(),
          YEARS.map((_year, index) => Number(activeScenario().costs.laborCostPerFte?.[key]?.[index] || 0)),
          `laborFteCostPerFte:${key}`,
          1000
        ),
      };
    });
  }
  return [{
    name: `${laborFteFilterLabel()} Cost / FTE`,
    color: "#6fa76b",
    editable: true,
    editDomain: laborFteEditableDomain(),
    data: laborFteMetricPairs(laborFteCostPerFteValues(activeScenario(), keys), "laborFteCostPerFte", 1000),
  }];
}

function nonLaborCategoryChartKey(key) {
  return `nonLaborCategory:${key}`;
}

function nonLaborDrilldownTopDownChartLines() {
  return [
    {
      name: "Bottom Up",
      color: costMeta.nonLabor.color,
      editable: true,
      editDomain: {
        moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: costPairsFromValuesForMode(activeScenario(), nonLaborBottomUpValues(activeScenario()), "nonLaborBottomUp", "nonLaborDrilldownTopDown"),
    },
    {
      name: "Top Down",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(activeScenario(), costValues(activeScenario(), "nonLabor"), "nonLabor", "nonLaborDrilldownTopDown"),
    },
  ];
}

function nonLaborAllocationControlChartLines() {
  const keys = selectedNonLaborAllocationKeys();
  if (!keys.length) {
    return [{
      name: "Select Categories",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(activeScenario(), YEARS.map(() => 0), "nonLaborAllocationControl", "nonLaborAllocationControl"),
    }];
  }
  const label = keys.length === 1
    ? displayLabel(nonLaborCategoryMeta[keys[0]].label)
    : `${keys.length} Categories`;
  return [{
    name: label,
    color: "#2f6f73",
    editable: true,
    editDomain: {
      moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: costPairsFromValuesForMode(activeScenario(), nonLaborAllocationControlValues(activeScenario()), "nonLaborAllocationControl", "nonLaborAllocationControl"),
  }];
}

function nonLaborCategorySelectedSpendChartLines() {
  const categoryKey = selectedNonLaborCategoryKey();
  const meta = nonLaborCategoryMeta[categoryKey];
  return [
    {
      name: "Department Build",
      color: meta.color,
      editable: true,
      editDomain: {
        moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: costPairsFromValuesForMode(
        activeScenario(),
        nonLaborCategoryDepartmentBottomUpValues(activeScenario(), categoryKey),
        `nonLaborCategoryDepartmentBottomUp:${categoryKey}`,
        "nonLaborCategorySelectedSpend"
      ),
    },
    {
      name: "Category Allocation",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(
        activeScenario(),
        nonLaborCategoryValues(activeScenario(), categoryKey),
        nonLaborCategoryChartKey(categoryKey),
        "nonLaborCategorySelectedSpend"
      ),
    },
  ];
}

function nonLaborCategoryDepartmentControlChartLines() {
  const categoryKey = selectedNonLaborCategoryKey();
  const selected = selectedNonLaborCategoryDepartmentKeys();
  if (!selected.length) {
    return [{
      name: "Select Departments",
      color: "#98a2b3",
      editable: false,
      data: costPairsFromValuesForMode(activeScenario(), YEARS.map(() => 0), `nonLaborCategoryDepartmentControl:${categoryKey}`, "nonLaborCategoryDepartmentControl"),
    }];
  }
  const editDomain = {
    moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
  };
  if (isCostChartSeparated("nonLaborCategoryDepartmentControl")) {
    return selected.map(key => {
      const meta = laborDepartmentMeta[key];
      return {
        name: displayLabel(meta.label),
        color: meta.color,
        departmentKey: key,
        editable: true,
        editDomain,
        data: costPairsFromValuesForMode(
          activeScenario(),
          nonLaborCategoryDepartmentValues(activeScenario(), categoryKey, key),
          nonLaborCategoryDepartmentChartKey(categoryKey, key),
          "nonLaborCategoryDepartmentControl"
        ),
      };
    });
  }
  const label = selected.length === 1
    ? displayLabel(laborDepartmentMeta[selected[0]].label)
    : `${selected.length} Departments`;
  return [{
    name: label,
    color: nonLaborCategoryMeta[categoryKey].color,
    editable: true,
    editDomain,
    data: costPairsFromValuesForMode(
      activeScenario(),
      nonLaborCategoryDepartmentControlValues(activeScenario()),
      `nonLaborCategoryDepartmentControl:${categoryKey}`,
      "nonLaborCategoryDepartmentControl"
    ),
  }];
}

function styleCostReferenceSeries(chart) {
  const option = chart.chart.getOption();
  const series = (option.series || []).map(seriesItem => {
    const name = String(seriesItem.name || "");
    if (name !== "Top Down" && name !== "Department Allocation" && name !== "Category Allocation") return seriesItem;
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

function initLaborDrilldownTopDownChart() {
  const meta = costMeta.labor;
  const chart = new NapkinChart(
    "cost-labor-topdown-chart",
    laborDrilldownTopDownChartLines(),
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
        ...costChartYAxisConfig("laborDrilldownTopDown", meta.yMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: value => formatCostChartTooltipValue(value, "laborDrilldownTopDown"),
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
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatCostChartTooltipValue(Number(rawValue), "laborDrilldownTopDown")}`);
            });
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
  applyCostChartYAxisMode(chart, "laborDrilldownTopDown", meta.yMax);
  chart._refreshChart();
  styleCostReferenceSeries(chart);
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const bottomUpValues = laborBottomUpValues(activeScenario());
    const selectedKeys = selectedLaborAllocationKeys();
    YEARS.forEach(year => {
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      const index = YEARS.indexOf(year);
      if (index >= 0 && editableCostYear(year)) {
        const previousValue = index > 0 ? Number(bottomUpValues[index - 1] || 0) : 0;
        const currentValue = Number(bottomUpValues[index] || 0);
        const targetValue = valueFromEditedCostPoint(previousValue, value, "laborDrilldownTopDown", currentValue, bottomUpValues, index);
        if (targetValue !== null) {
          setLaborDepartmentBottomUpTotalForYear(activeScenario(), index, targetValue, {
            lockOthers: state.laborAllocationOthersLocked,
            selectedKeys,
          });
          bottomUpValues[index] = laborDepartmentTotalForYear(activeScenario().costs.laborDepartments, index);
        }
      }
    });
    if (isCostChartYoY("laborDrilldownTopDown")) {
      rememberCostControlYears("laborBottomUp", editableCostYears());
    } else {
      rememberCostControlPoints("laborBottomUp", chart.lines[0].data);
    }
    saveScenarios();
    styleCostReferenceSeries(chart);
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.laborDrilldownTopDown = chart;
}

function initLaborAllocationControlChart() {
  const chart = new NapkinChart(
    "cost-labor-allocation-control-chart",
    laborAllocationControlChartLines(),
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
        ...costChartYAxisConfig("laborAllocationControl", costMeta.labor.yMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: formatCostTooltipValue,
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = [];
          let year = items[0]?.axisValue;
          items
            .filter(item => item && item.seriesType === "line" && item.seriesIndex % 2 === 1)
            .forEach(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              year = Array.isArray(data) ? data[0] : item.axisValue;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatCostChartTooltipValue(Number(rawValue), "laborAllocationControl")}`);
            });
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
  applyCostChartYAxisMode(chart, "laborAllocationControl", costMeta.labor.yMax);
  chart._refreshChart();
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    const selectedKeys = selectedLaborAllocationKeys();
    if (!selectedKeys.length) {
      syncCostCharts({ excludeChart: chart });
      renderCostDrilldownView();
      return;
    }
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const allocationValues = laborAllocationControlValues(activeScenario());
    YEARS.forEach((year, index) => {
      if (!editableCostYear(year)) return;
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      const previousValue = index > 0 ? Number(allocationValues[index - 1] || 0) : 0;
      const currentValue = Number(allocationValues[index] || 0);
      const targetValue = valueFromEditedCostPoint(previousValue, value, "laborAllocationControl", currentValue, allocationValues, index);
      if (targetValue === null) return;
      setLaborAllocationControlValue(activeScenario(), index, selectedKeys, targetValue, {
        lockTop: state.laborAllocationTopLocked,
      });
      allocationValues[index] = laborDepartmentGroupTotalForYear(activeScenario().costs.laborDepartments, selectedKeys, index);
    });
    if (isCostChartYoY("laborAllocationControl")) {
      rememberCostControlYears("laborAllocationControl", editableCostYears());
      rememberCostControlYears("laborBottomUp", editableCostYears());
    } else {
      rememberCostControlPoints("laborAllocationControl", chart.lines[0].data);
    }
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.laborAllocationControl = chart;
}

function initLaborAllocationPctChart() {
  const element = document.getElementById("cost-labor-allocation-pct-chart");
  if (!element || typeof NapkinChartArea === "undefined") return;
  const baseOption = {
    animation: false,
    xAxis: {
      type: "value",
      min: FIRST_COST_EDIT_YEAR,
      max: YEARS[YEARS.length - 1],
      minInterval: 1,
      axisLabel: { formatter: formatAxisYear },
    },
    yAxis: {
      type: "value",
      min: 0,
      max: 100,
      axisLabel: { formatter: value => formatPercentMetric(Number(value), 0) },
    },
    topAreaLabel: "Unassigned",
    topAreaColor: "#98a2b3",
    areaTint: 0.42,
    grid: { left: 12, right: 18, top: 18, bottom: 34, containLabel: true },
    tooltip: { trigger: "axis" },
  };
  const chart = new NapkinChartArea(
    "cost-labor-allocation-pct-chart",
    laborAllocationPctLineObjects(),
    true,
    baseOption,
    "none"
  );
  chart.enableZoomBar = false;
  chart.windowStartX = FIRST_COST_EDIT_YEAR;
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    const keys = selectedLaborAllocationKeys();
    if (keys.length < 2) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const storedLines = [];
    chart.lines.forEach((line, fallbackIndex) => {
      const bandIndex = Number.isInteger(line.bandIndex)
        ? line.bandIndex
        : chart.lines.length - fallbackIndex - 1;
      storedLines[bandIndex] = clone(line.data);
    });
    normalizeCostPlan(activeScenario());
    activeScenario().costPctTotalLines[laborAllocationPctStorageKey(keys)] = storedLines.filter(Array.isArray);
    applyLaborAllocationPctShares();
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.laborAllocationPct = chart;
  syncLaborAllocationPctChart();
}

function applyLaborFteSeparatedLegend(chart, chartKey) {
  if (!chart?.baseOption) return;
  const separated = isCostChartSeparated(chartKey);
  chart.baseOption.legend = separated
    ? {
      show: true,
      type: "scroll",
      selectedMode: false,
      top: 0,
      left: 8,
      right: 8,
      itemWidth: 10,
      itemHeight: 8,
      textStyle: { color: "#475467", fontSize: 11 },
    }
    : { show: false };
  chart.baseOption.grid = {
    ...(chart.baseOption.grid || {}),
    top: separated ? 48 : 14,
  };
}

function laborFteTooltipRows(params) {
  const items = Array.isArray(params) ? params : [params];
  const seen = new Set();
  return items.filter(item => {
    if (!item || item.seriesType !== "line") return false;
    if (item.seriesIndex % 2 !== 1) return false;
    if (seen.has(item.seriesName)) return false;
    seen.add(item.seriesName);
    return true;
  });
}

function formatLaborFteSeparatedTooltip(params, chartKey, valueFormatter) {
  const items = Array.isArray(params) ? params : [params];
  const chart = state.costCharts[chartKey];
  const hoveredName = isCostChartSeparated(chartKey) ? chart?._appHoveredLineName : "";
  let year = items[0]?.axisValue;
  const rows = laborFteTooltipRows(items).map(item => {
    const data = Array.isArray(item.data) ? item.data : item.value;
    year = Array.isArray(data) ? data[0] : item.axisValue;
    const rawValue = Array.isArray(data) ? data[1] : item.value;
    const label = escapeHtml(displayLabel(item.seriesName || ""));
    const row = `${item.marker || ""} ${label}: ${valueFormatter(Number(rawValue))}`;
    return hoveredName && item.seriesName === hoveredName ? `<strong>${row}</strong>` : row;
  });
  return [tooltipHeader(year), ...rows].join("<br/>");
}

function setLaborFteSeparatedHover(chart, chartKey, lineName) {
  const nextName = isCostChartSeparated(chartKey) ? String(lineName || "") : "";
  if (chart?._appHoveredLineName === nextName) return;
  chart._appHoveredLineName = nextName;
  styleLaborFteSeparatedLegendSeries(chart, chartKey);
}

function bindLaborFteSeparatedHover(chart, chartKey) {
  if (!chart?.chart || chart._appSeparatedHoverBound) return;
  chart._appSeparatedHoverBound = true;
  chart.chart.on("mouseover", params => {
    if (!isCostChartSeparated(chartKey)) return;
    const lineName = params?.seriesName || (params?.componentType === "legend" ? params.name : "");
    if (!lineName) return;
    setLaborFteSeparatedHover(chart, chartKey, lineName);
  });
  chart.chart.on("mouseout", params => {
    if (!isCostChartSeparated(chartKey)) return;
    if (!params?.seriesName && params?.componentType !== "legend") return;
    setLaborFteSeparatedHover(chart, chartKey, "");
  });
  chart.chart.getZr().on("globalout", () => setLaborFteSeparatedHover(chart, chartKey, ""));
}

function styleLaborFteSeparatedLegendSeries(chart, chartKey) {
  if (!chart?.chart) return;
  bindLaborFteSeparatedHover(chart, chartKey);
  if (!isCostChartSeparated(chartKey)) {
    chart._appHoveredLineName = "";
  }
  const hoveredName = isCostChartSeparated(chartKey) ? chart._appHoveredLineName : "";
  const option = chart.chart.getOption();
  const series = (option.series || []).map(seriesItem => {
    if (seriesItem.type !== "line") return seriesItem;
    const baseWidth = Number(seriesItem._appBaseLineWidth || seriesItem.lineStyle?.width || 2);
    const isHovered = hoveredName && seriesItem.name === hoveredName;
    return {
      ...seriesItem,
      _appBaseLineWidth: baseWidth,
      legendHoverLink: false,
      emphasis: {
        ...(seriesItem.emphasis || {}),
        disabled: true,
      },
      lineStyle: {
        ...(seriesItem.lineStyle || {}),
        width: isHovered ? Math.max(baseWidth + 2, 4) : baseWidth,
      },
    };
  });
  chart.chart.setOption({ legend: chart.baseOption.legend, series }, false);
}

function initLaborFteSelectedCostChart() {
  const chart = new NapkinChart(
    "cost-labor-fte-selected-cost-chart",
    laborFteSelectedCostChartLines(),
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
        max: costMeta.labor.yMax,
        axisLabel: { formatter: formatCostAxisValue },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatLaborFteSeparatedTooltip(params, "laborFteSelectedCost", formatCostTooltipValue),
      },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  applyLaborFteSeparatedLegend(chart, "laborFteSelectedCost");
  chart._refreshChart();
  styleCostReferenceSeries(chart);
  styleLaborFteSeparatedLegendSeries(chart, "laborFteSelectedCost");
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const keys = activeLaborFteKeys();
    if (isCostChartSeparated("laborFteSelectedCost")) {
      chart.lines
        .filter(line => line.departmentKey)
        .forEach(line => {
          const key = line.departmentKey;
          YEARS.forEach((year, index) => {
            if (!editableCostYear(year)) return;
            const value = interpolateNapkinLineValue(line.data, year);
            if (Number.isFinite(value)) {
              setLaborFteBuildCostForYear(activeScenario(), index, [key], value * 1000000);
            }
          });
          rememberCostControlPoints(`laborFteBuildCost:${key}`, line.data);
        });
    } else {
      YEARS.forEach((year, index) => {
        if (!editableCostYear(year)) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        if (Number.isFinite(value)) {
          setLaborFteBuildCostForYear(activeScenario(), index, keys, value * 1000000);
        }
      });
      rememberCostControlPoints("laborFteBuildCost", chart.lines[0].data);
    }
    saveScenarios();
    styleCostReferenceSeries(chart);
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.laborFteSelectedCost = chart;
}

function laborFtePctStateKey(chartKey) {
  return `${chartKey}Pct`;
}

function initLaborFtePctChart(chartKey, domId, applyShares) {
  const element = document.getElementById(domId);
  if (!element || typeof NapkinChartArea === "undefined") return;
  const baseOption = {
    animation: false,
    xAxis: {
      type: "value",
      min: FIRST_COST_EDIT_YEAR,
      max: YEARS[YEARS.length - 1],
      minInterval: 1,
      axisLabel: { formatter: formatAxisYear },
    },
    yAxis: {
      type: "value",
      min: 0,
      max: 100,
      axisLabel: { formatter: value => formatPercentMetric(Number(value), 0) },
    },
    topAreaLabel: "Unassigned",
    topAreaColor: "#98a2b3",
    areaTint: 0.42,
    grid: { left: 12, right: 18, top: 18, bottom: 34, containLabel: true },
    tooltip: { trigger: "axis" },
  };
  const chart = new NapkinChartArea(domId, laborFtePctLineObjects(chartKey), true, baseOption, "none");
  chart.enableZoomBar = false;
  chart.windowStartX = FIRST_COST_EDIT_YEAR;
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (chartKey === "laborFteCount" && laborFteDependentEditBlocked(chartKey)) {
      rejectLaborFteDependentEdit(chart);
      return;
    }
    const keys = activeLaborFteKeys();
    if (keys.length < 2) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const storedLines = [];
    chart.lines.forEach((line, fallbackIndex) => {
      const bandIndex = Number.isInteger(line.bandIndex)
        ? line.bandIndex
        : chart.lines.length - fallbackIndex - 1;
      storedLines[bandIndex] = clone(line.data);
    });
    normalizeCostPlan(activeScenario());
    activeScenario().costPctTotalLines[laborFtePctStorageKey(chartKey, keys)] = storedLines.filter(Array.isArray);
    applyShares();
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts[laborFtePctStateKey(chartKey)] = chart;
  syncLaborFtePctChart(chartKey);
}

function initLaborFteCountChart() {
  const chart = new NapkinChart(
    "cost-labor-fte-count-chart",
    laborFteCountChartLines(),
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
        max: 500,
        axisLabel: { formatter: formatFteAxisValue },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatLaborFteSeparatedTooltip(params, "laborFteCount", formatFteTooltipValue),
      },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  applyLaborFteSeparatedLegend(chart, "laborFteCount");
  chart._refreshChart();
  styleLaborFteSeparatedLegendSeries(chart, "laborFteCount");
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (laborFteDependentEditBlocked("laborFteCount")) {
      rejectLaborFteDependentEdit(chart);
      return;
    }
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const keys = activeLaborFteKeys();
    if (isCostChartSeparated("laborFteCount")) {
      const lockedBuildValues = laborFteBuildCostValues(activeScenario(), keys);
      chart.lines
        .filter(line => line.departmentKey)
        .forEach(line => {
          const key = line.departmentKey;
          YEARS.forEach((year, index) => {
            if (!editableCostYear(year)) return;
            const value = interpolateNapkinLineValue(line.data, year);
            if (Number.isFinite(value)) {
              activeScenario().costs.laborFte[key][index] = Math.max(0, value);
            }
          });
          rememberCostControlPoints(`laborFteCount:${key}`, line.data);
        });
      if (state.laborFteTopLocked) {
        YEARS.forEach((year, index) => {
          if (!editableCostYear(year)) return;
          adjustLaborCostPerFteToBuildCost(activeScenario(), index, keys, lockedBuildValues[index]);
        });
      }
    } else {
      YEARS.forEach((year, index) => {
        if (!editableCostYear(year)) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        if (Number.isFinite(value)) {
          const lockedBuild = laborFteBuildCostValues(activeScenario(), keys)[index];
          distributeLaborFteAcrossKeys(activeScenario(), index, keys, value);
          if (state.laborFteTopLocked) {
            adjustLaborCostPerFteToBuildCost(activeScenario(), index, keys, lockedBuild);
          }
        }
      });
      rememberCostControlPoints("laborFteCount", chart.lines[0].data);
    }
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.laborFteCount = chart;
}

function initLaborFteCostPerFteChart() {
  const chart = new NapkinChart(
    "cost-labor-fte-cost-per-fte-chart",
    laborFteCostPerFteChartLines(),
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
        max: 200,
        axisLabel: { formatter: formatCostPerFteAxisValue },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatLaborFteSeparatedTooltip(params, "laborFteCostPerFte", formatCostPerFteTooltipValue),
      },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  applyLaborFteSeparatedLegend(chart, "laborFteCostPerFte");
  chart._refreshChart();
  styleLaborFteSeparatedLegendSeries(chart, "laborFteCostPerFte");
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (laborFteDependentEditBlocked("laborFteCostPerFte")) {
      rejectLaborFteDependentEdit(chart);
      return;
    }
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const keys = activeLaborFteKeys();
    if (isCostChartSeparated("laborFteCostPerFte")) {
      const lockedBuildValues = laborFteBuildCostValues(activeScenario(), keys);
      chart.lines
        .filter(line => line.departmentKey)
        .forEach(line => {
          const key = line.departmentKey;
          YEARS.forEach((year, index) => {
            if (!editableCostYear(year)) return;
            const value = interpolateNapkinLineValue(line.data, year);
            if (Number.isFinite(value)) {
              activeScenario().costs.laborCostPerFte[key][index] = Math.max(0, value * 1000);
            }
          });
          rememberCostControlPoints(`laborFteCostPerFte:${key}`, line.data);
        });
      if (state.laborFteTopLocked) {
        YEARS.forEach((year, index) => {
          if (!editableCostYear(year)) return;
          adjustLaborFteToBuildCost(activeScenario(), index, keys, lockedBuildValues[index]);
        });
      }
    } else {
      YEARS.forEach((year, index) => {
        if (!editableCostYear(year)) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        if (Number.isFinite(value)) {
          const lockedBuild = laborFteBuildCostValues(activeScenario(), keys)[index];
          scaleLaborCostPerFteAcrossKeys(activeScenario(), index, keys, value * 1000);
          if (state.laborFteTopLocked) {
            adjustLaborFteToBuildCost(activeScenario(), index, keys, lockedBuild);
          }
        }
      });
      rememberCostControlPoints("laborFteCostPerFte", chart.lines[0].data);
    }
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.laborFteCostPerFte = chart;
}

function initNonLaborDrilldownTopDownChart() {
  const meta = costMeta.nonLabor;
  const chart = new NapkinChart(
    "cost-non-labor-topdown-chart",
    nonLaborDrilldownTopDownChartLines(),
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
        ...costChartYAxisConfig("nonLaborDrilldownTopDown", meta.yMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: value => formatCostChartTooltipValue(value, "nonLaborDrilldownTopDown"),
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
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatCostChartTooltipValue(Number(rawValue), "nonLaborDrilldownTopDown")}`);
            });
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
  applyCostChartYAxisMode(chart, "nonLaborDrilldownTopDown", meta.yMax);
  chart._refreshChart();
  styleCostReferenceSeries(chart);
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const bottomUpValues = nonLaborBottomUpValues(activeScenario());
    const selectedKeys = selectedNonLaborAllocationKeys();
    YEARS.forEach(year => {
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      const index = YEARS.indexOf(year);
      if (index >= 0 && editableCostYear(year)) {
        const previousValue = index > 0 ? Number(bottomUpValues[index - 1] || 0) : 0;
        const currentValue = Number(bottomUpValues[index] || 0);
        const targetValue = valueFromEditedCostPoint(previousValue, value, "nonLaborDrilldownTopDown", currentValue, bottomUpValues, index);
        if (targetValue !== null) {
          setNonLaborCategoryBottomUpTotalForYear(activeScenario(), index, targetValue, {
            lockOthers: state.nonLaborAllocationOthersLocked,
            selectedKeys,
          });
          bottomUpValues[index] = nonLaborCategoryTotalForYear(activeScenario().costs.nonLaborCategories, index);
        }
      }
    });
    if (isCostChartYoY("nonLaborDrilldownTopDown")) {
      rememberCostControlYears("nonLaborBottomUp", editableCostYears());
    } else {
      rememberCostControlPoints("nonLaborBottomUp", chart.lines[0].data);
    }
    saveScenarios();
    styleCostReferenceSeries(chart);
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.nonLaborDrilldownTopDown = chart;
}

function initNonLaborAllocationControlChart() {
  const chart = new NapkinChart(
    "cost-non-labor-allocation-control-chart",
    nonLaborAllocationControlChartLines(),
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
        ...costChartYAxisConfig("nonLaborAllocationControl", costMeta.nonLabor.yMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: formatCostTooltipValue,
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = [];
          let year = items[0]?.axisValue;
          items
            .filter(item => item && item.seriesType === "line" && item.seriesIndex % 2 === 1)
            .forEach(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              year = Array.isArray(data) ? data[0] : item.axisValue;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatCostChartTooltipValue(Number(rawValue), "nonLaborAllocationControl")}`);
            });
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
  applyCostChartYAxisMode(chart, "nonLaborAllocationControl", costMeta.nonLabor.yMax);
  chart._refreshChart();
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    const selectedKeys = selectedNonLaborAllocationKeys();
    if (!selectedKeys.length) {
      syncCostCharts({ excludeChart: chart });
      renderCostDrilldownView();
      return;
    }
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const allocationValues = nonLaborAllocationControlValues(activeScenario());
    YEARS.forEach((year, index) => {
      if (!editableCostYear(year)) return;
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      const previousValue = index > 0 ? Number(allocationValues[index - 1] || 0) : 0;
      const currentValue = Number(allocationValues[index] || 0);
      const targetValue = valueFromEditedCostPoint(previousValue, value, "nonLaborAllocationControl", currentValue, allocationValues, index);
      if (targetValue === null) return;
      setNonLaborAllocationControlValue(activeScenario(), index, selectedKeys, targetValue, {
        lockTop: state.nonLaborAllocationTopLocked,
      });
      allocationValues[index] = nonLaborCategoryGroupTotalForYear(activeScenario().costs.nonLaborCategories, selectedKeys, index);
    });
    if (isCostChartYoY("nonLaborAllocationControl")) {
      rememberCostControlYears("nonLaborAllocationControl", editableCostYears());
      rememberCostControlYears("nonLaborBottomUp", editableCostYears());
    } else {
      rememberCostControlPoints("nonLaborAllocationControl", chart.lines[0].data);
    }
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.nonLaborAllocationControl = chart;
}

function initNonLaborCategorySelectedSpendChart() {
  const selectedYMax = nonLaborCategoryMeta[selectedNonLaborCategoryKey()]?.yMax || costMeta.nonLabor.yMax;
  const chart = new NapkinChart(
    "cost-non-labor-category-selected-chart",
    nonLaborCategorySelectedSpendChartLines(),
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
        ...costChartYAxisConfig("nonLaborCategorySelectedSpend", selectedYMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatLaborFteSeparatedTooltip(params, "nonLaborCategorySelectedSpend", value => {
          return formatCostChartTooltipValue(value, "nonLaborCategorySelectedSpend");
        }),
      },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  applyCostChartYAxisMode(chart, "nonLaborCategorySelectedSpend", selectedYMax);
  chart._refreshChart();
  styleCostReferenceSeries(chart);
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const categoryKey = selectedNonLaborCategoryKey();
    const bottomUpValues = nonLaborCategoryDepartmentBottomUpValues(activeScenario(), categoryKey);
    const selectedKeys = selectedNonLaborCategoryDepartmentKeys();
    YEARS.forEach((year, index) => {
      if (!editableCostYear(year)) return;
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      const previousValue = index > 0 ? Number(bottomUpValues[index - 1] || 0) : 0;
      const currentValue = Number(bottomUpValues[index] || 0);
      let targetValue = valueFromEditedCostPoint(previousValue, value, "nonLaborCategorySelectedSpend", currentValue, bottomUpValues, index);
      if (targetValue === null) return;
      if (state.nonLaborCategoryTopLocked) {
        targetValue = Number(activeScenario().costs.nonLaborCategories[categoryKey][index] || 0);
      }
      setNonLaborCategoryDepartmentBottomUpTotalForYear(activeScenario(), index, categoryKey, targetValue, {
        lockOthers: state.nonLaborCategoryOthersLocked,
        selectedKeys,
      });
      bottomUpValues[index] = nonLaborCategoryDepartmentBottomUpValues(activeScenario(), categoryKey)[index];
    });
    if (isCostChartYoY("nonLaborCategorySelectedSpend")) {
      rememberCostControlYears(`nonLaborCategoryDepartmentBottomUp:${categoryKey}`, editableCostYears());
    } else {
      rememberCostControlPoints(`nonLaborCategoryDepartmentBottomUp:${categoryKey}`, chart.lines[0].data);
    }
    saveScenarios();
    styleCostReferenceSeries(chart);
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.nonLaborCategorySelectedSpend = chart;
}

function initNonLaborCategoryDepartmentControlChart() {
  const selectedYMax = nonLaborCategoryMeta[selectedNonLaborCategoryKey()]?.yMax || costMeta.nonLabor.yMax;
  const chart = new NapkinChart(
    "cost-non-labor-category-department-control-chart",
    nonLaborCategoryDepartmentControlChartLines(),
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
        ...costChartYAxisConfig("nonLaborCategoryDepartmentControl", selectedYMax),
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => formatLaborFteSeparatedTooltip(params, "nonLaborCategoryDepartmentControl", value => {
          return formatCostChartTooltipValue(value, "nonLaborCategoryDepartmentControl");
        }),
      },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  applyCostChartYAxisMode(chart, "nonLaborCategoryDepartmentControl", selectedYMax);
  applyLaborFteSeparatedLegend(chart, "nonLaborCategoryDepartmentControl");
  chart._refreshChart();
  styleLaborFteSeparatedLegendSeries(chart, "nonLaborCategoryDepartmentControl");
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    const categoryKey = selectedNonLaborCategoryKey();
    const selectedKeys = selectedNonLaborCategoryDepartmentKeys();
    if (!selectedKeys.length) {
      syncCostCharts({ excludeChart: chart });
      renderCostDrilldownView();
      return;
    }
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    if (isCostChartSeparated("nonLaborCategoryDepartmentControl")) {
      chart.lines
        .filter(line => line.departmentKey)
        .forEach(line => {
          const key = line.departmentKey;
          const departmentValues = nonLaborCategoryDepartmentValues(activeScenario(), categoryKey, key);
          YEARS.forEach((year, index) => {
            if (!editableCostYear(year)) return;
            const value = interpolateNapkinLineValue(line.data, year);
            const previousValue = index > 0 ? Number(departmentValues[index - 1] || 0) : 0;
            const currentValue = Number(departmentValues[index] || 0);
            const targetValue = valueFromEditedCostPoint(previousValue, value, "nonLaborCategoryDepartmentControl", currentValue, departmentValues, index);
            if (targetValue !== null) {
              activeScenario().costs.nonLaborCategoryDepartments[categoryKey][key][index] = targetValue;
              addCostControlPointIfControlled(nonLaborCategoryDepartmentChartKey(categoryKey, key), year);
              if (state.nonLaborCategoryTopLocked) {
                const topValue = Number(activeScenario().costs.nonLaborCategories[categoryKey][index] || 0);
                setNonLaborCategoryDepartmentBottomUpTotalForYear(activeScenario(), index, categoryKey, topValue);
              }
            }
          });
          if (isCostChartYoY("nonLaborCategoryDepartmentControl")) {
            rememberCostControlYears(nonLaborCategoryDepartmentChartKey(categoryKey, key), editableCostYears());
          } else {
            rememberCostControlPoints(nonLaborCategoryDepartmentChartKey(categoryKey, key), line.data);
          }
        });
    } else {
      const allocationValues = nonLaborCategoryDepartmentControlValues(activeScenario());
      YEARS.forEach((year, index) => {
        if (!editableCostYear(year)) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        const previousValue = index > 0 ? Number(allocationValues[index - 1] || 0) : 0;
        const currentValue = Number(allocationValues[index] || 0);
        const targetValue = valueFromEditedCostPoint(previousValue, value, "nonLaborCategoryDepartmentControl", currentValue, allocationValues, index);
        if (targetValue === null) return;
        setNonLaborCategoryDepartmentControlValue(activeScenario(), index, categoryKey, selectedKeys, targetValue, {
          lockTop: state.nonLaborCategoryTopLocked,
        });
        allocationValues[index] = nonLaborCategoryDepartmentGroupTotalForYear(
          activeScenario().costs.nonLaborCategoryDepartments,
          categoryKey,
          selectedKeys,
          index
        );
      });
      if (isCostChartYoY("nonLaborCategoryDepartmentControl")) {
        rememberCostControlYears(`nonLaborCategoryDepartmentControl:${categoryKey}`, editableCostYears());
      } else {
        rememberCostControlPoints(`nonLaborCategoryDepartmentControl:${categoryKey}`, chart.lines[0].data);
      }
    }
    saveScenarios();
    syncCostCharts({ excludeChart: chart });
    renderCostDrilldownView();
  };
  state.costCharts.nonLaborCategoryDepartmentControl = chart;
}

function initLaborDepartmentChart(key) {
  const meta = laborDepartmentMeta[key];
  const chartKey = laborDepartmentChartKey(key);
  const chart = new NapkinChart(
    meta.chartId,
    laborDepartmentChartLines(key),
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
        axisLabel: { formatter: formatCostAxisValue },
      },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        valueFormatter: formatCostTooltipValue,
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
              rows.push(`${item.marker || ""} ${item.seriesName}: ${formatCostTooltipValue(Number(rawValue))}`);
            });
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
    if (state.syncingCostCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingCostCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    YEARS.forEach(year => {
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      setLaborDepartmentValue(activeScenario(), key, year, value);
    });
    rememberCostControlPoints(chartKey, chart.lines[0].data);
    saveScenarios();
    styleComparisonSeries(chart);
    syncCostCharts({ excludeChart: chart });
  };
  state.costCharts[chartKey] = chart;
}

function laborDepartmentMixEchartSeriesValues(key) {
  const values = laborDepartmentValues(activeScenario(), key);
  if (isCostChartYoY("laborDepartmentMix")) {
    return YEARS.map((_year, index) => {
      return costYoYPercentFromValues(values, index);
    });
  }
  return values.map(value => value / 1000000);
}

function renderLaborDepartmentMixEchart() {
  const chart = state.outputCharts.costLaborDepartmentMix;
  if (!chart) return;
  const isYoY = isCostChartYoY("laborDepartmentMix");
  const series = LABOR_DEPARTMENT_KEYS.map(key => {
    const meta = laborDepartmentMeta[key];
    return {
      name: displayLabel(meta.label),
      type: "line",
      data: laborDepartmentMixEchartSeriesValues(key),
      smooth: false,
      symbolSize: 5,
      lineStyle: { width: 2, color: meta.color },
      itemStyle: { color: meta.color },
    };
  });
  chart.setOption({
    animation: false,
    color: LABOR_DEPARTMENT_KEYS.map(key => laborDepartmentMeta[key].color),
    legend: {
      type: "scroll",
      top: 0,
      left: 0,
      right: 0,
      itemWidth: 12,
      itemHeight: 8,
      textStyle: { color: "#475467", fontSize: 11 },
    },
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const items = Array.isArray(params) ? params : [params];
        return [
          tooltipHeader(items[0]?.axisValue),
          ...items.map(item => {
            const value = Number(item.value);
            const formatted = isYoY ? formatPercentMetric(value, 1) : formatCostTooltipValue(value);
            return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatted}`;
          }),
        ].join("<br/>");
      },
    },
    grid: { left: 12, right: 18, top: 52, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String), axisLabel: { formatter: formatAxisYear } },
    yAxis: {
      type: "value",
      axisLabel: {
        formatter: value => isYoY ? formatPercentMetric(Number(value), 0) : formatCostAxisValue(value),
      },
    },
    series,
  }, true);
}

function nonLaborCategoryMixEchartSeriesValues(key) {
  const values = nonLaborCategoryValues(activeScenario(), key);
  if (isCostChartYoY("nonLaborCategoryMix")) {
    return YEARS.map((_year, index) => {
      return costYoYPercentFromValues(values, index);
    });
  }
  return values.map(value => value / 1000000);
}

function renderNonLaborCategoryMixEchart() {
  const chart = state.outputCharts.costNonLaborCategoryMix;
  if (!chart) return;
  const isYoY = isCostChartYoY("nonLaborCategoryMix");
  const series = NON_LABOR_CATEGORY_KEYS.map(key => {
    const meta = nonLaborCategoryMeta[key];
    return {
      name: displayLabel(meta.label),
      type: "line",
      data: nonLaborCategoryMixEchartSeriesValues(key),
      smooth: false,
      symbolSize: 5,
      lineStyle: { width: 2, color: meta.color },
      itemStyle: { color: meta.color },
    };
  });
  chart.setOption({
    animation: false,
    color: NON_LABOR_CATEGORY_KEYS.map(key => nonLaborCategoryMeta[key].color),
    legend: {
      type: "scroll",
      top: 0,
      left: 0,
      right: 0,
      itemWidth: 12,
      itemHeight: 8,
      textStyle: { color: "#475467", fontSize: 11 },
    },
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const items = Array.isArray(params) ? params : [params];
        return [
          tooltipHeader(items[0]?.axisValue),
          ...items.map(item => {
            const value = Number(item.value);
            const formatted = isYoY ? formatPercentMetric(value, 1) : formatCostTooltipValue(value);
            return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatted}`;
          }),
        ].join("<br/>");
      },
    },
    grid: { left: 12, right: 18, top: 52, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String), axisLabel: { formatter: formatAxisYear } },
    yAxis: {
      type: "value",
      axisLabel: {
        formatter: value => isYoY ? formatPercentMetric(Number(value), 0) : formatCostAxisValue(value),
      },
    },
    series,
  }, true);
}

function nonLaborCategoryDepartmentMixEchartSeriesValues(departmentKey) {
  const categoryKey = selectedNonLaborCategoryKey();
  const values = nonLaborCategoryDepartmentValues(activeScenario(), categoryKey, departmentKey);
  if (isCostChartYoY("nonLaborCategoryDepartmentControl")) {
    return YEARS.map((_year, index) => costYoYPercentFromValues(values, index));
  }
  return values.map(value => value / 1000000);
}

function renderNonLaborCategoryDepartmentMixEchart() {
  const chart = state.outputCharts.costNonLaborCategoryDepartmentMix;
  if (!chart) return;
  const isYoY = isCostChartYoY("nonLaborCategoryDepartmentControl");
  const series = LABOR_DEPARTMENT_KEYS.map(key => {
    const meta = laborDepartmentMeta[key];
    return {
      name: displayLabel(meta.label),
      type: "line",
      data: nonLaborCategoryDepartmentMixEchartSeriesValues(key),
      smooth: false,
      symbolSize: 5,
      lineStyle: { width: 2, color: meta.color },
      itemStyle: { color: meta.color },
    };
  });
  chart.setOption({
    animation: false,
    color: LABOR_DEPARTMENT_KEYS.map(key => laborDepartmentMeta[key].color),
    legend: {
      type: "scroll",
      top: 0,
      left: 0,
      right: 0,
      itemWidth: 12,
      itemHeight: 8,
      textStyle: { color: "#475467", fontSize: 11 },
    },
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const items = Array.isArray(params) ? params : [params];
        return [
          tooltipHeader(items[0]?.axisValue),
          ...items.map(item => {
            const value = Number(item.value);
            const formatted = isYoY ? formatPercentMetric(value, 1) : formatCostTooltipValue(value);
            return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatted}`;
          }),
        ].join("<br/>");
      },
    },
    grid: { left: 12, right: 18, top: 52, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String), axisLabel: { formatter: formatAxisYear } },
    yAxis: {
      type: "value",
      axisLabel: {
        formatter: value => isYoY ? formatPercentMetric(Number(value), 0) : formatCostAxisValue(value),
      },
    },
    series,
  }, true);
}

function syncCostCharts({ excludeChart = null } = {}) {
  state.syncingCostCharts = true;
  Object.entries(costMeta).forEach(([key]) => {
    const chart = state.costCharts[key];
    if (!chart || chart === excludeChart) return;
    chart.lines = costChartLines(key);
    chart._refreshChart();
    styleComparisonSeries(chart);
  });
  LABOR_DEPARTMENT_KEYS.forEach(key => {
    const chart = state.costCharts[laborDepartmentChartKey(key)];
    if (!chart || chart === excludeChart) return;
    chart.lines = laborDepartmentChartLines(key);
    chart._refreshChart();
    styleComparisonSeries(chart);
  });
  const laborTopDownChart = state.costCharts.laborDrilldownTopDown;
  if (laborTopDownChart && laborTopDownChart !== excludeChart) {
    laborTopDownChart.lines = laborDrilldownTopDownChartLines();
    applyCostChartYAxisMode(laborTopDownChart, "laborDrilldownTopDown", costMeta.labor.yMax);
    laborTopDownChart._refreshChart();
    styleCostReferenceSeries(laborTopDownChart);
  }
  renderLaborDepartmentMixEchart();
  const laborAllocationChart = state.costCharts.laborAllocationControl;
  if (laborAllocationChart && laborAllocationChart !== excludeChart) {
    laborAllocationChart.lines = laborAllocationControlChartLines();
    applyCostChartYAxisMode(laborAllocationChart, "laborAllocationControl", costMeta.labor.yMax);
    laborAllocationChart._refreshChart();
  }
  const laborAllocationPctChart = state.costCharts.laborAllocationPct;
  if (laborAllocationPctChart && laborAllocationPctChart !== excludeChart) {
    syncLaborAllocationPctChart();
  }
  const laborFteSelectedCostChart = state.costCharts.laborFteSelectedCost;
  if (laborFteSelectedCostChart && laborFteSelectedCostChart !== excludeChart) {
    laborFteSelectedCostChart.lines = laborFteSelectedCostChartLines();
    applyLaborFteSeparatedLegend(laborFteSelectedCostChart, "laborFteSelectedCost");
    laborFteSelectedCostChart._refreshChart();
    styleCostReferenceSeries(laborFteSelectedCostChart);
    styleLaborFteSeparatedLegendSeries(laborFteSelectedCostChart, "laborFteSelectedCost");
  }
  const laborFteSelectedCostPctChart = state.costCharts.laborFteSelectedCostPct;
  if (laborFteSelectedCostPctChart && laborFteSelectedCostPctChart !== excludeChart) {
    syncLaborFtePctChart("laborFteSelectedCost");
  }
  const laborFteCountChart = state.costCharts.laborFteCount;
  if (laborFteCountChart && laborFteCountChart !== excludeChart) {
    laborFteCountChart.lines = laborFteCountChartLines();
    applyLaborFteSeparatedLegend(laborFteCountChart, "laborFteCount");
    laborFteCountChart._refreshChart();
    styleLaborFteSeparatedLegendSeries(laborFteCountChart, "laborFteCount");
  }
  const laborFteCountPctChart = state.costCharts.laborFteCountPct;
  if (laborFteCountPctChart && laborFteCountPctChart !== excludeChart) {
    syncLaborFtePctChart("laborFteCount");
  }
  const laborFteCostPerFteChart = state.costCharts.laborFteCostPerFte;
  if (laborFteCostPerFteChart && laborFteCostPerFteChart !== excludeChart) {
    laborFteCostPerFteChart.lines = laborFteCostPerFteChartLines();
    applyLaborFteSeparatedLegend(laborFteCostPerFteChart, "laborFteCostPerFte");
    laborFteCostPerFteChart._refreshChart();
    styleLaborFteSeparatedLegendSeries(laborFteCostPerFteChart, "laborFteCostPerFte");
  }
  const nonLaborTopDownChart = state.costCharts.nonLaborDrilldownTopDown;
  if (nonLaborTopDownChart && nonLaborTopDownChart !== excludeChart) {
    nonLaborTopDownChart.lines = nonLaborDrilldownTopDownChartLines();
    applyCostChartYAxisMode(nonLaborTopDownChart, "nonLaborDrilldownTopDown", costMeta.nonLabor.yMax);
    nonLaborTopDownChart._refreshChart();
    styleCostReferenceSeries(nonLaborTopDownChart);
  }
  renderNonLaborCategoryMixEchart();
  renderNonLaborCategoryDepartmentMixEchart();
  const nonLaborAllocationChart = state.costCharts.nonLaborAllocationControl;
  if (nonLaborAllocationChart && nonLaborAllocationChart !== excludeChart) {
    nonLaborAllocationChart.lines = nonLaborAllocationControlChartLines();
    applyCostChartYAxisMode(nonLaborAllocationChart, "nonLaborAllocationControl", costMeta.nonLabor.yMax);
    nonLaborAllocationChart._refreshChart();
  }
  const nonLaborCategorySelectedSpendChart = state.costCharts.nonLaborCategorySelectedSpend;
  if (nonLaborCategorySelectedSpendChart && nonLaborCategorySelectedSpendChart !== excludeChart) {
    const selectedYMax = nonLaborCategoryMeta[selectedNonLaborCategoryKey()]?.yMax || costMeta.nonLabor.yMax;
    nonLaborCategorySelectedSpendChart.lines = nonLaborCategorySelectedSpendChartLines();
    applyCostChartYAxisMode(nonLaborCategorySelectedSpendChart, "nonLaborCategorySelectedSpend", selectedYMax);
    nonLaborCategorySelectedSpendChart._refreshChart();
    styleCostReferenceSeries(nonLaborCategorySelectedSpendChart);
  }
  const nonLaborCategoryDepartmentControlChart = state.costCharts.nonLaborCategoryDepartmentControl;
  if (nonLaborCategoryDepartmentControlChart && nonLaborCategoryDepartmentControlChart !== excludeChart) {
    const selectedYMax = nonLaborCategoryMeta[selectedNonLaborCategoryKey()]?.yMax || costMeta.nonLabor.yMax;
    nonLaborCategoryDepartmentControlChart.lines = nonLaborCategoryDepartmentControlChartLines();
    applyCostChartYAxisMode(nonLaborCategoryDepartmentControlChart, "nonLaborCategoryDepartmentControl", selectedYMax);
    applyLaborFteSeparatedLegend(nonLaborCategoryDepartmentControlChart, "nonLaborCategoryDepartmentControl");
    nonLaborCategoryDepartmentControlChart._refreshChart();
    styleLaborFteSeparatedLegendSeries(nonLaborCategoryDepartmentControlChart, "nonLaborCategoryDepartmentControl");
  }
  state.syncingCostCharts = false;
  renderCostDrilldownView();
}

function syncLaborAllocationPctChart() {
  const chart = state.costCharts.laborAllocationPct;
  if (!chart) return;
  const keys = selectedLaborAllocationKeys();
  if (keys.length < 2) {
    chart.lines = [];
    chart._refreshChart();
    return;
  }
  const topKey = keys[keys.length - 1];
  const topMeta = laborDepartmentMeta[topKey];
  chart.lines = laborAllocationPctLineObjects();
  chart.topAreaLabel = displayLabel(topMeta.label);
  chart.topAreaColor = topMeta.color;
  chart.baseOption.topAreaLabel = chart.topAreaLabel;
  chart.baseOption.topAreaColor = chart.topAreaColor;
  chart._refreshChart();
  chart.resize();
  requestAnimationFrame(() => chart.resize());
}

function syncLaborFtePctChart(chartKey) {
  const chart = state.costCharts[laborFtePctStateKey(chartKey)];
  if (!chart) return;
  const keys = activeLaborFteKeys();
  if (keys.length < 2) {
    chart.lines = [];
    chart._refreshChart();
    return;
  }
  const topKey = keys[keys.length - 1];
  const topMeta = laborDepartmentMeta[topKey];
  chart.lines = laborFtePctLineObjects(chartKey);
  chart.topAreaLabel = displayLabel(topMeta.label);
  chart.topAreaColor = topMeta.color;
  chart.baseOption.topAreaLabel = chart.topAreaLabel;
  chart.baseOption.topAreaColor = chart.topAreaColor;
  chart._refreshChart();
  chart.resize();
  requestAnimationFrame(() => chart.resize());
}

function initEchartById(id) {
  const element = document.getElementById(id);
  return element ? echarts.init(element) : null;
}

function initOutputCharts() {
  state.outputCharts.revenue = initEchartById("revenue-chart");
  state.outputCharts.gap = initEchartById("gap-chart");
  state.outputCharts.revenueDrilldownTotal = initEchartById("revenue-drilldown-total-chart");
  state.outputCharts.retentionDefense = initEchartById("retention-defense-chart");
  state.outputCharts.retentionRevenueImpact = initEchartById("retention-revenue-impact-chart");
  state.outputCharts.initiativeImpact = initEchartById("initiative-impact-chart");
  state.outputCharts.initiativeTeamImpact = initEchartById("initiative-team-impact-chart");
  state.outputCharts.newCustomerTotal = initEchartById("new-customers-total-chart");
  state.outputCharts.newCustomerUnitBridge = initEchartById("new-customers-unit-bridge-chart");
  state.outputCharts.newCustomerNewUnitsBridge = initEchartById("new-customers-new-units-bridge-chart");
  state.outputCharts.newCustomerAllCohorts = initEchartById("new-customers-all-cohorts-chart");
  state.outputCharts.revUnitRevenueComparison = initEchartById("rev-unit-revenue-comparison-chart");
  state.outputCharts.revUnitAggregateRpu = initEchartById("rev-unit-aggregate-rpu-chart");
  state.outputCharts.treeRevenue = initEchartById("tree-revenue-chart");
  state.outputCharts.treeLtrRevenue = initEchartById("tree-ltr-revenue-chart");
  state.outputCharts.treeStrRevenue = initEchartById("tree-str-revenue-chart");
  state.outputCharts.treeStrProperties = initEchartById("tree-str-properties-chart");
  state.outputCharts.treeStrRevPerProperty = initEchartById("tree-str-rev-per-property-chart");
  state.outputCharts.treeGrowthRevenue = initEchartById("tree-growth-revenue-chart");
  state.outputCharts.treeGrowthNewPayingCustomers = initEchartById("tree-growth-new-paying-customers-chart");
  state.outputCharts.treeGrowthRevPerNewPayingCustomer = initEchartById("tree-growth-rev-per-new-paying-customer-chart");
  state.outputCharts.treeOtherRevenue = initEchartById("tree-other-revenue-chart");
  state.outputCharts.treeProfit = initEchartById("tree-profit-chart");
  state.outputCharts.treeCost = initEchartById("tree-cost-chart");
  state.outputCharts.treeLaborCost = initEchartById("tree-labor-cost-chart");
  state.outputCharts.treeNonLaborCost = initEchartById("tree-non-labor-cost-chart");
  state.outputCharts.treeLaborFte = initEchartById("tree-labor-fte-chart");
  state.outputCharts.treeLaborCostPerFte = initEchartById("tree-labor-cost-per-fte-chart");
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
  state.outputCharts.saasRuleOf40 = initEchartById("saas-rule-of-40-chart");
  state.outputCharts.saasCac = initEchartById("saas-cac-chart");
  state.outputCharts.saasRevenuePerFte = initEchartById("saas-revenue-per-fte-chart");
  state.outputCharts.growthRevenueProxyCustomers = initEchartById("growth-revenue-proxy-customers-chart");
  state.outputCharts.costLaborDepartmentMix = initEchartById("cost-labor-department-mix-chart");
  state.outputCharts.costNonLaborCategoryMix = initEchartById("cost-non-labor-category-mix-chart");
  state.outputCharts.costNonLaborCategoryDepartmentMix = initEchartById("cost-non-labor-category-department-mix-chart");
  window.addEventListener("resize", () => {
    Object.values(state.outputCharts).forEach(chart => chart?.resize());
    Object.values(state.growthCharts).forEach(chart => chart?.resize());
    Object.values(state.strCharts).forEach(chart => chart?.resize());
    Object.values(state.otherRevenueCharts).forEach(chart => chart?.resize());
    Object.values(state.revenueDrilldownCharts).forEach(chart => chart?.resize());
    Object.values(state.profitCharts).forEach(chart => chart?.resize());
  });
}

function initRetentionDefenseCharts() {
  const element = document.getElementById("retention-pct-total-chart");
  if (!element || typeof NapkinChartArea === "undefined") return;
  const baseOption = {
    animation: false,
    xAxis: {
      type: "value",
      min: FIRST_FORECAST_YEAR,
      max: YEARS[YEARS.length - 1],
      minInterval: 1,
      axisLabel: { formatter: formatAxisYear },
    },
    yAxis: {
      type: "value",
      min: 0,
      max: 100,
      axisLabel: { formatter: value => `${trimNumber(Number(value), 0)}%` },
    },
    topAreaLabel: "Unassigned",
    topAreaColor: "#98a2b3",
    areaTint: 0.42,
    grid: { left: 12, right: 18, top: 18, bottom: 34, containLabel: true },
    tooltip: { trigger: "axis" },
  };
  const chart = new NapkinChartArea(
    "retention-pct-total-chart",
    retentionPctTotalLineObjects(),
    true,
    baseOption,
    "none"
  );
  chart.enableZoomBar = false;
  chart.windowStartX = FIRST_FORECAST_YEAR;
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const storedLines = [];
    chart.lines.forEach((line, fallbackIndex) => {
      const bandIndex = Number.isInteger(line.bandIndex)
        ? line.bandIndex
        : chart.lines.length - fallbackIndex - 1;
      storedLines[bandIndex] = clone(line.data);
    });
    defenseEntry().pctTotalLines = storedLines.filter(Array.isArray);
    saveScenarios();
    renderRetentionRevenueImpact();
  };
  state.defenseCharts.retentionPctTotal = chart;
  syncRetentionPctTotalChart();
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
  renderLtrSetControls(outputs);

  if (!state.outputCharts.gap) return;
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

function ltrTopMatchesBottomUp(scenario = activeScenario(), outputs = calculateOutputs(scenario)) {
  const topDown = ltrRevenueTopDownValues(scenario, outputs);
  const bottomUp = ltrRevenueBottomUpValues(scenario, outputs);
  return YEARS.every((_year, index) => Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0)) < 1);
}

function renderLtrSetControls(outputs = calculateOutputs(activeScenario())) {
  const button = document.getElementById("set-ltr-bottom-to-top");
  if (!button) return;
  button.disabled = ltrTopMatchesBottomUp(activeScenario(), outputs);
}

function setLtrBottomToTop() {
  pushUndoSnapshot();
  const scenario = activeScenario();
  const outputs = calculateOutputs(scenario);
  normalizeRevenuePaths(scenario);
  scenario.revenuePaths.ltr.topDown = ltrRevenueBottomUpValues(scenario, outputs).map(value => Math.max(0, Number(value) || 0));
  scenario.revenuePaths.ltr.edited = true;
  scenario.revenuePaths.ltr.controlPoints.topDown = YEARS.filter(year => year >= FIRST_FORECAST_YEAR);
  saveScenarios();
  syncRevenueDrilldownCharts();
  syncProfitCharts();
  renderAll();
}

const leafDefenseConfigs = {
  retentionRate: {
    title: "Retention Rate",
    shortTitle: "Retention",
    format: "percent",
    color: driverMeta.retention.color,
    focusChart: "chart-retention",
    treeAction: "chart",
    values: scenario => driverValues(scenario, "retention"),
    baseValues: () => baseDrivers.retention,
    applyBaseline: scenario => {
      scenario.drivers.retention = clone(baseDrivers.retention);
    },
  },
  profilesReturningCustomer: {
    title: "Profiles / Returning Customer",
    shortTitle: "Profiles / Returning Customer",
    format: "decimal",
    color: driverMeta.profilesReturning.color,
    focusChart: "chart-profiles-returning",
    treeAction: "chart",
    values: scenario => driverValues(scenario, "profilesReturning"),
    baseValues: () => baseDrivers.profilesReturning,
    applyBaseline: scenario => {
      scenario.drivers.profilesReturning = clone(baseDrivers.profilesReturning);
    },
  },
  profilesNewCustomer: {
    title: "Profiles / New Customer",
    shortTitle: "Profiles / New Customer",
    format: "decimal",
    color: driverMeta.profilesNew.color,
    focusChart: "chart-profiles-new",
    treeAction: "chart",
    values: scenario => driverValues(scenario, "profilesNew"),
    baseValues: () => baseDrivers.profilesNew,
    applyBaseline: scenario => {
      scenario.drivers.profilesNew = clone(baseDrivers.profilesNew);
    },
  },
  revenueReturningProfile: {
    title: "Revenue / Returning Profile",
    shortTitle: "Rev / Returning Profile",
    format: "currency2",
    color: driverMeta.revReturningProfile.color,
    focusChart: "chart-rev-returning",
    treeAction: "chart",
    values: scenario => driverValues(scenario, "revReturningProfile"),
    baseValues: () => baseDrivers.revReturningProfile,
    applyBaseline: scenario => {
      scenario.drivers.revReturningProfile = clone(baseDrivers.revReturningProfile);
    },
  },
  revenueNewProfile: {
    title: "Revenue / New Profile",
    shortTitle: "Rev / New Profile",
    format: "currency2",
    color: driverMeta.revNewProfile.color,
    focusChart: "chart-rev-new",
    treeAction: "chart",
    values: scenario => driverValues(scenario, "revNewProfile"),
    baseValues: () => baseDrivers.revNewProfile,
    applyBaseline: scenario => {
      scenario.drivers.revNewProfile = clone(baseDrivers.revNewProfile);
    },
  },
  existingPropertyNewPayingCustomers: {
    title: "New Paying Customers From Existing Properties",
    shortTitle: "Existing Properties",
    format: "integer",
    color: TOP_DOWN_COLOR,
    treeAction: "newCustomers",
    values: scenario => YEARS.map(year => existingPropertyNewCustomersTotal(scenario.newCustomerDrilldown.counts, year)),
    baseValues: () => YEARS.map(year => existingPropertyNewCustomersTotal(baseNewCustomerCohortCounts, year)),
    newCustomerDelta: scenario => {
      const activeValues = YEARS.map(year => existingPropertyNewCustomersTotal(scenario.newCustomerDrilldown.counts, year));
      const baseValues = YEARS.map(year => existingPropertyNewCustomersTotal(baseNewCustomerCohortCounts, year));
      return activeValues.map((value, index) => value - baseValues[index]);
    },
  },
  growthNewPayingCustomers: {
    title: "Growth New Paying Customers",
    shortTitle: "Growth New Customers",
    format: "integer",
    color: TOP_DOWN_COLOR,
    treeAction: "growthRevenue",
    values: scenario => growthRevenueProxyCustomers(scenario),
    baseValues: () => baseDrivers.newCustomers,
  },
  growthRevPerNewPayingCustomer: {
    title: "Growth Revenue / New Paying Customer",
    shortTitle: "Growth Rev / New Customer",
    format: "currency2",
    color: BOTTOM_UP_COLOR,
    treeAction: "growthRevenue",
    values: scenario => growthRevenueRevPerNewCustomerValues(scenario),
    baseValues: () => createDefaultGrowthRevenuePlan().revPerNewPayingCustomer,
    applyBaseline: scenario => {
      growthRevenuePlan(scenario).revPerNewPayingCustomer = clone(createDefaultGrowthRevenuePlan().revPerNewPayingCustomer);
    },
  },
  strProperties: {
    title: "STR Properties",
    shortTitle: "STR Properties",
    format: "integer",
    color: "#6fa76b",
    treeAction: "strRevenue",
    values: scenario => strRevenuePropertiesValues(scenario),
    baseValues: () => createDefaultStrRevenuePlan().numStrProperties,
    applyBaseline: scenario => {
      strRevenuePlan(scenario).numStrProperties = clone(createDefaultStrRevenuePlan().numStrProperties);
    },
  },
  strRevPerProperty: {
    title: "Revenue / STR Property",
    shortTitle: "Rev / STR Property",
    format: "currency0",
    color: BOTTOM_UP_COLOR,
    treeAction: "strRevenue",
    values: scenario => strRevenueRevPerPropertyValues(scenario),
    baseValues: () => createDefaultStrRevenuePlan().revPerStrProperty,
    applyBaseline: scenario => {
      strRevenuePlan(scenario).revPerStrProperty = clone(createDefaultStrRevenuePlan().revPerStrProperty);
    },
  },
  otherRevenue: {
    title: "Other Revenue",
    shortTitle: "Other Revenue",
    format: "currency0",
    color: "#a96c50",
    treeAction: "otherRevenue",
    values: scenario => otherRevenueTopDownValues(scenario),
    baseValues: () => createDefaultOtherRevenuePlan().topDown,
    applyBaseline: scenario => {
      otherRevenuePlan(scenario).topDown = clone(createDefaultOtherRevenuePlan().topDown);
    },
  },
  nonLaborCost: {
    title: "Non-Labor Cost",
    shortTitle: "Non-Labor Cost",
    format: "currency0",
    color: costMeta.nonLabor.color,
    treeAction: "nonLaborCost",
    values: scenario => costValues(scenario, "nonLabor"),
    baseValues: () => baseCosts.nonLabor,
    applyBaseline: scenario => {
      normalizeCostPlan(scenario);
      scenario.costs.nonLabor = clone(baseCosts.nonLabor);
    },
  },
  laborFte: {
    title: "Labor FTE / Equivalent",
    shortTitle: "Labor FTE",
    format: "decimal",
    color: "#4f7fb8",
    treeAction: "laborFte",
    values: scenario => laborFteTotalValues(scenario),
    baseValues: () => {
      const base = { costs: clone(baseCosts) };
      normalizeCostPlan(base);
      return laborFteTotalValues(base);
    },
    applyBaseline: scenario => {
      normalizeCostPlan(scenario);
      const base = { costs: clone(baseCosts) };
      normalizeCostPlan(base);
      scenario.costs.laborFte = clone(base.costs.laborFte);
    },
  },
  laborCostPerFte: {
    title: "Labor Cost / FTE Equivalent",
    shortTitle: "Cost / FTE",
    format: "currency0",
    color: "#7b6fb8",
    treeAction: "laborFte",
    values: scenario => laborCostPerFteAggregateValues(scenario),
    baseValues: () => {
      const base = { costs: clone(baseCosts) };
      normalizeCostPlan(base);
      return laborCostPerFteAggregateValues(base);
    },
  },
  newCustomersPerUnit: {
    title: "New Customers per Unit",
    shortTitle: "New Customers / Unit",
    format: "decimal",
    color: BOTTOM_UP_COLOR,
    treeAction: "unitDrilldown",
    values: scenario => YEARS.map(year => newCustomerUnitCustomersPerUnitValue(scenario, year)),
    baseValues: () => {
      const base = makeBaseScenario();
      return YEARS.map(year => newCustomerUnitCustomersPerUnitValue(base, year));
    },
    newCustomerDelta: scenario => {
      const base = makeBaseScenario();
      return YEARS.map(year => {
        const units = newCustomerUnitNewUnitsValue(scenario, year);
        return units * (newCustomerUnitCustomersPerUnitValue(scenario, year) - newCustomerUnitCustomersPerUnitValue(base, year));
      });
    },
  },
  newProperties: {
    title: "New Properties",
    shortTitle: "New Properties",
    format: "integer",
    color: BOTTOM_UP_COLOR,
    treeAction: "newUnitsDrilldown",
    values: scenario => YEARS.map(year => newCustomerNewUnitsNewPropertiesValue(scenario, year)),
    baseValues: () => {
      const base = makeBaseScenario();
      return YEARS.map(year => newCustomerNewUnitsNewPropertiesValue(base, year));
    },
    newCustomerDelta: scenario => {
      const base = makeBaseScenario();
      return YEARS.map(year => {
        const unitsPerProperty = newCustomerNewUnitsUnitsPerPropertyValue(scenario, year);
        const customersPerUnit = newCustomerUnitCustomersPerUnitValue(scenario, year);
        const activeUnits = newCustomerNewUnitsNewPropertiesValue(scenario, year) * unitsPerProperty;
        const baseUnits = newCustomerNewUnitsNewPropertiesValue(base, year) * unitsPerProperty;
        return (activeUnits - baseUnits) * customersPerUnit;
      });
    },
  },
  newUnitsPerProperty: {
    title: "New Units per Property",
    shortTitle: "Units / New Property",
    format: "decimal",
    color: BOTTOM_UP_COLOR,
    treeAction: "newUnitsDrilldown",
    values: scenario => YEARS.map(year => newCustomerNewUnitsUnitsPerPropertyValue(scenario, year)),
    baseValues: () => {
      const base = makeBaseScenario();
      return YEARS.map(year => newCustomerNewUnitsUnitsPerPropertyValue(base, year));
    },
    newCustomerDelta: scenario => {
      const base = makeBaseScenario();
      return YEARS.map(year => {
        const properties = newCustomerNewUnitsNewPropertiesValue(scenario, year);
        const customersPerUnit = newCustomerUnitCustomersPerUnitValue(scenario, year);
        const activeUnits = properties * newCustomerNewUnitsUnitsPerPropertyValue(scenario, year);
        const baseUnits = properties * newCustomerNewUnitsUnitsPerPropertyValue(base, year);
        return (activeUnits - baseUnits) * customersPerUnit;
      });
    },
  },
  priorPayingCustomers: {
    title: "Prior Year Paying Customers",
    shortTitle: "Prior Paying Customers",
    format: "integer",
    color: "#98a2b3",
    treeAction: "newCustomers",
    values: scenario => {
      const outputs = calculateOutputs(scenario);
      return outputs.totalCustomers.map((value, index) => index === 0 ? null : outputs.totalCustomers[index - 1]);
    },
    baseValues: () => {
      const outputs = calculateOutputs({ drivers: clone(baseDrivers), newCustomerSource: "topDown", newCustomerDrilldown: createDefaultNewCustomerDrilldown() });
      return outputs.totalCustomers.map((value, index) => index === 0 ? null : outputs.totalCustomers[index - 1]);
    },
  },
};

function activeDefenseConfig() {
  return leafDefenseConfigs[state.activeDefenseKey] || leafDefenseConfigs.retentionRate;
}

function defenseEntry(scenario = activeScenario(), key = state.activeDefenseKey) {
  normalizeScenario(scenario);
  if (!scenario.defenses[key]) {
    scenario.defenses[key] = { initiatives: [], pctTotalLines: [] };
  }
  return scenario.defenses[key];
}

function formatDefenseMetricValue(value, config = activeDefenseConfig()) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  if (config.format === "percent") return `${trimNumber(Number(value) * 100, 1)}%`;
  if (config.format === "currency0") return formatCurrency(Number(value), 0);
  if (config.format === "currency2") return formatCurrency(Number(value), 2);
  if (config.format === "integer") return Math.round(Number(value)).toLocaleString("en-US");
  if (config.format === "decimal") return trimNumber(Number(value), 2);
  return formatCompactNumber(Number(value));
}

function formatDefenseAxisValue(value, config = activeDefenseConfig()) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  if (config.format === "percent") return `${trimNumber(Number(value) * 100, 0)}%`;
  if (config.format === "currency0") return formatCompactCurrency(Number(value));
  if (config.format === "currency2") return formatCompactCurrency(Number(value));
  if (config.format === "integer") return formatCompactNumber(Number(value));
  if (config.format === "decimal") return trimNumber(Number(value), 2);
  return formatCompactNumber(Number(value));
}

function openLeafDefense(key) {
  if (!leafDefenseConfigs[key]) return;
  state.activeDefenseKey = key;
  setView("retentionDefense");
  renderRetentionDefense();
}

function retentionDefenseInitiatives(scenario = activeScenario(), key = state.activeDefenseKey) {
  normalizeScenario(scenario);
  return defenseEntry(scenario, key).initiatives;
}

function retentionPctTotalBands(scenario = activeScenario(), key = state.activeDefenseKey) {
  const initiatives = retentionDefenseInitiatives(scenario, key);
  return initiatives.length ? initiatives : ["Unknown"];
}

function retentionPctTotalBandLabel(band, index) {
  const name = typeof band === "string" ? band : band?.name;
  if (!isAnonymizedView()) return name || "Unknown";
  return band === "Unknown" ? "Unknown" : `Initiative ${index + 1}`;
}

function retentionInitiativeTeamLabel(initiative) {
  const teams = initiative?.teams || [];
  return teams.length ? teams.join(", ") : "No team tags";
}

function retentionPctTotalLineData(scenario = activeScenario(), key = state.activeDefenseKey, bands = retentionPctTotalBands(scenario, key)) {
  normalizeScenario(scenario);
  const stored = defenseEntry(scenario, key).pctTotalLines;
  const expectedCount = Math.max(0, bands.length - 1);
  if (stored.length !== expectedCount) {
    defenseEntry(scenario, key).pctTotalLines = defaultRetentionPctTotalLines(bands);
  }
  return defenseEntry(scenario, key).pctTotalLines;
}

function retentionPctTotalLineObjects(scenario = activeScenario(), key = state.activeDefenseKey) {
  const bands = retentionPctTotalBands(scenario, key);
  const lineData = retentionPctTotalLineData(scenario, key, bands);
  return lineData
    .map((data, index) => ({
      name: retentionPctTotalBandLabel(bands[index], index),
      color: metricTreeScenarioPalette[index % metricTreeScenarioPalette.length],
      bandIndex: index,
      editable: true,
      editDomain: {
        moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: clone(data),
    }))
    .reverse();
}

function retentionPctTotalSharesForYear(scenario, year, key = state.activeDefenseKey) {
  const bands = retentionPctTotalBands(scenario, key);
  if (bands.length <= 1) return [1];
  const boundaries = retentionPctTotalLineData(scenario, key, bands)
    .map(data => interpolateNapkinLineValue(data, year))
    .map(value => Number.isFinite(value) ? Math.max(0, Math.min(100, value)) : 0);
  const shares = [];
  let previous = 0;
  boundaries.forEach(boundary => {
    const share = Math.max(0, boundary - previous) / 100;
    shares.push(share);
    previous = boundary;
  });
  shares.push(Math.max(0, 100 - previous) / 100);
  return shares;
}

function retentionRevenueDeltaToBase(scenario = activeScenario(), key = state.activeDefenseKey) {
  const config = leafDefenseConfigs[key] || activeDefenseConfig();
  const activeOutputs = calculateOutputs(scenario);
  if (typeof config.newCustomerDelta === "function") {
    const deltas = config.newCustomerDelta(scenario);
    const baselineNewCustomers = driverValues(scenario, "newCustomers").map((value, index) => Math.max(0, value - (deltas[index] || 0)));
    const baseOutputs = calculateOutputsWithNewCustomers(scenario, baselineNewCustomers);
    return activeOutputs.revenue.map((value, index) => value - baseOutputs.revenue[index]);
  }
  if (typeof config.applyBaseline !== "function") {
    return YEARS.map(() => 0);
  }
  const baseRetentionScenario = clone(scenario);
  config.applyBaseline(baseRetentionScenario);
  const baseRetentionOutputs = calculateOutputs(baseRetentionScenario);
  return activeOutputs.revenue.map((value, index) => value - baseRetentionOutputs.revenue[index]);
}

function retentionInitiativeRevenueImpacts(scenario = activeScenario(), key = state.activeDefenseKey) {
  const bands = retentionPctTotalBands(scenario, key);
  const deltas = retentionRevenueDeltaToBase(scenario, key);
  return bands.map((band, bandIndex) => {
    const yearly = YEARS.map((year, yearIndex) => {
      const shares = retentionPctTotalSharesForYear(scenario, year, key);
      return deltas[yearIndex] * (shares[bandIndex] || 0);
    });
    return {
      name: retentionPctTotalBandLabel(band, bandIndex),
      color: metricTreeScenarioPalette[bandIndex % metricTreeScenarioPalette.length],
      yearly,
      forecastTotal: yearly.reduce((sum, value, index) => editableYear(YEARS[index]) ? sum + value : sum, 0),
      finalYear: yearly[yearly.length - 1],
    };
  });
}

function renderRetentionRevenueImpact() {
  const container = document.getElementById("retention-revenue-impact");
  if (!container) return;
  const impacts = retentionInitiativeRevenueImpacts();
  const totalForecast = impacts.reduce((sum, item) => sum + item.forecastTotal, 0);
  const totalFinalYear = impacts.reduce((sum, item) => sum + item.finalYear, 0);
  container.innerHTML = `
    <div class="revenue-impact-total">
      <span>2026-2029 Delta to BAU</span>
      <strong>${formatCurrency(totalForecast, 0)}</strong>
      <small>2029: ${formatCurrency(totalFinalYear, 0)}</small>
    </div>
    ${impacts.map(item => `
      <div class="revenue-impact-item">
        <span class="impact-color" style="background:${item.color}"></span>
        <div>
          <strong>${escapeHtml(item.name)}</strong>
          <small>2026-2029</small>
        </div>
        <div>
          <strong>${formatCurrency(item.forecastTotal, 0)}</strong>
          <small>2029: ${formatCurrency(item.finalYear, 0)}</small>
        </div>
      </div>
    `).join("")}
  `;
  renderRetentionRevenueImpactChart(impacts);
}

function renderRetentionRevenueImpactChart(impacts = retentionInitiativeRevenueImpacts()) {
  const chartElement = document.getElementById("retention-revenue-impact-chart");
  const chartTypeControl = document.getElementById("retention-impact-chart-type");
  const cumulativeButton = document.getElementById("retention-impact-cumulative");
  const lineButton = document.getElementById("retention-impact-line");
  const barButton = document.getElementById("retention-impact-bar");
  const chart = state.outputCharts.retentionRevenueImpact;
  if (!chartElement || !chartTypeControl || !chart) return;

  chartElement.hidden = false;
  chartTypeControl.hidden = false;
  if (cumulativeButton) cumulativeButton.classList.toggle("active", state.retentionImpactCumulative);
  if (lineButton) lineButton.classList.toggle("active", state.retentionImpactChartType === "line");
  if (barButton) barButton.classList.toggle("active", state.retentionImpactChartType === "bar");

  const isBar = state.retentionImpactChartType === "bar";
  const seriesImpacts = impacts.map(item => {
    let running = 0;
    return {
      ...item,
      chartData: state.retentionImpactCumulative
        ? item.yearly.map(value => {
          running += value;
          return running;
        })
        : item.yearly,
    };
  });
  chart.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(item.value), 0)}`),
      ].join("<br/>"),
    },
    legend: { type: "scroll", top: 0 },
    grid: { left: 12, right: 18, top: 48, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
    series: seriesImpacts.map(item => ({
      name: item.name,
      type: isBar ? "bar" : "line",
      data: item.chartData,
      ...(isBar ? { stack: "retentionImpact", barMaxWidth: 46 } : { symbolSize: 5 }),
      lineStyle: isBar ? undefined : { color: item.color, width: 2 },
      itemStyle: { color: item.color },
    })),
  }, true);
  chart.resize();
  requestAnimationFrame(() => chart.resize());
}

function syncRetentionPctTotalChart() {
  const chart = state.defenseCharts.retentionPctTotal;
  const element = document.getElementById("retention-pct-total-chart");
  const empty = document.getElementById("retention-pct-total-empty");
  if (!chart || !element || !empty) return;
  const bands = retentionPctTotalBands();
  empty.hidden = true;
  element.hidden = false;

  if (bands.length <= 1) {
    chart.lines = [];
    chart.chart.setOption({
      animation: false,
      tooltip: {
        trigger: "axis",
        formatter: params => [
          tooltipHeader(params[0]?.axisValue),
          ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatPercentMetric(Number(item.value), 0)}`),
        ].join("<br/>"),
      },
      legend: { top: 0 },
      grid: { left: 12, right: 18, top: 42, bottom: 34, containLabel: true },
      xAxis: {
        type: "value",
        min: FIRST_FORECAST_YEAR,
        max: YEARS[YEARS.length - 1],
        minInterval: 1,
        axisLabel: { formatter: formatAxisYear },
      },
      yAxis: {
        type: "value",
        min: 0,
        max: 100,
        axisLabel: { formatter: value => formatPercentMetric(Number(value), 0) },
      },
      series: [{
        name: retentionPctTotalBandLabel(bands[0], 0),
        type: "line",
        data: [[FIRST_FORECAST_YEAR, 100], [YEARS[YEARS.length - 1], 100]],
        showSymbol: false,
        lineStyle: { width: 0 },
        itemStyle: { color: metricTreeScenarioPalette[0] },
        areaStyle: { opacity: 0.52, color: metricTreeScenarioPalette[0] },
      }],
    }, true);
    chart.resize();
    requestAnimationFrame(() => chart.resize());
    renderRetentionRevenueImpact();
    return;
  }

  chart.lines = retentionPctTotalLineObjects();
  chart.topAreaLabel = retentionPctTotalBandLabel(bands[bands.length - 1], bands.length - 1);
  chart.topAreaColor = metricTreeScenarioPalette[(bands.length - 1) % metricTreeScenarioPalette.length];
  chart.baseOption.topAreaLabel = chart.topAreaLabel;
  chart.baseOption.topAreaColor = chart.topAreaColor;
  chart._refreshChart();
  chart.resize();
  requestAnimationFrame(() => chart.resize());
  renderRetentionRevenueImpact();
}

function renderRetentionDefense() {
  const config = activeDefenseConfig();
  const title = document.getElementById("defense-view-title");
  const subtitle = document.getElementById("defense-view-subtitle");
  const chartTitle = document.getElementById("defense-metric-chart-title");
  if (title) title.textContent = `${displayLabel(config.shortTitle)} Defense`;
  if (subtitle) {
    subtitle.textContent = `Capture the rationale behind ${displayLabel(config.shortTitle).toLowerCase()} assumptions, especially where they depart from BAU.`;
  }
  if (chartTitle) chartTitle.textContent = `${displayLabel(config.shortTitle)} vs Base`;

  const chart = state.outputCharts.retentionDefense;
  if (chart) {
    const scenario = activeScenario();
    const scenarioName = displayScenarioName(scenario);
    const activeValues = config.values(scenario);
    const baseValues = config.baseValues(scenario);
    chart.setOption({
      animation: false,
      tooltip: {
        trigger: "axis",
        formatter: params => [
          tooltipHeader(params[0]?.axisValue),
          ...params.map(item => `${item.marker} ${displayLabel(item.seriesName)}: ${formatDefenseMetricValue(Number(item.value), config)}`),
        ].join("<br/>"),
      },
      legend: { top: 0 },
      grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
      xAxis: { type: "category", data: YEARS.map(String) },
      yAxis: { type: "value", axisLabel: { formatter: value => formatDefenseAxisValue(value, config) } },
      series: [
        {
          name: scenarioName,
          type: "line",
          data: activeValues,
          symbolSize: 6,
          lineStyle: { color: config.color, width: 3 },
          itemStyle: { color: config.color },
        },
        {
          name: "Base Scenario",
          type: "line",
          data: baseValues,
          symbolSize: 5,
          lineStyle: { color: "#98a2b3", width: 2, type: "dashed" },
          itemStyle: { color: "#98a2b3" },
        },
      ],
    }, true);
  }

  const list = document.getElementById("retention-initiative-list");
  if (!list) return;
  const initiatives = retentionDefenseInitiatives();
  list.innerHTML = initiatives.length
    ? initiatives.map((initiative, index) => `
      <div class="initiative-chip">
        <div class="initiative-chip-main">
          <span>${escapeHtml(isAnonymizedView() ? `Initiative ${index + 1}` : initiative.name)}</span>
          <small>${escapeHtml(retentionInitiativeTeamLabel(initiative))}</small>
        </div>
        <button type="button" data-remove-retention-initiative="${index}" aria-label="Remove ${escapeHtml(initiative.name)}">Remove</button>
        <div class="initiative-team-tags">
          ${initiative.teams.map(team => `
            <span class="initiative-team-pill">
              ${escapeHtml(team)}
              <button
                type="button"
                data-remove-retention-initiative-team="${index}"
                data-team="${escapeHtml(team)}"
                aria-label="Remove ${escapeHtml(team)} from ${escapeHtml(initiative.name)}"
              >x</button>
            </span>
          `).join("")}
          ${INITIATIVE_TEAMS.some(team => !initiative.teams.includes(team)) ? `
            <button
              type="button"
              class="initiative-team-add"
              data-show-retention-initiative-team-picker="${index}"
              aria-label="Add team tag to ${escapeHtml(initiative.name)}"
            >+</button>
            <select class="initiative-team-picker" data-add-retention-initiative-team="${index}" hidden>
              <option value="">Choose team</option>
              ${INITIATIVE_TEAMS
                .filter(team => !initiative.teams.includes(team))
                .map(team => `<option value="${escapeHtml(team)}">${escapeHtml(team)}</option>`)
                .join("")}
            </select>
          ` : ""}
        </div>
      </div>
    `).join("")
    : `<div class="initiative-empty">No initiatives added yet.</div>`;
  syncRetentionPctTotalChart();
}

function addRetentionInitiative(name) {
  const trimmed = name.trim();
  if (!trimmed) return;
  pushUndoSnapshot();
  const scenario = activeScenario();
  retentionDefenseInitiatives(scenario).push({ name: trimmed, teams: [] });
  defenseEntry(scenario).pctTotalLines = defaultRetentionPctTotalLines(retentionPctTotalBands(scenario));
  saveScenarios();
  renderRetentionDefense();
}

function addRetentionInitiativeTeam(index, team) {
  if (!INITIATIVE_TEAMS.includes(team)) return;
  const scenario = activeScenario();
  const initiative = retentionDefenseInitiatives(scenario)[index];
  if (!initiative || initiative.teams.includes(team)) return;
  pushUndoSnapshot();
  initiative.teams = [...initiative.teams, team];
  saveScenarios();
  renderRetentionDefense();
}

function removeRetentionInitiativeTeam(index, team) {
  const scenario = activeScenario();
  const initiative = retentionDefenseInitiatives(scenario)[index];
  if (!initiative || !initiative.teams.includes(team)) return;
  pushUndoSnapshot();
  initiative.teams = initiative.teams.filter(item => item !== team);
  saveScenarios();
  renderRetentionDefense();
}

function removeRetentionInitiative(index) {
  const scenario = activeScenario();
  const initiatives = retentionDefenseInitiatives(scenario);
  if (!initiatives[index]) return;
  pushUndoSnapshot();
  initiatives.splice(index, 1);
  defenseEntry(scenario).pctTotalLines = defaultRetentionPctTotalLines(retentionPctTotalBands(scenario));
  saveScenarios();
  renderRetentionDefense();
}

function initiativeImpactRows(scenario = activeScenario()) {
  return LEAF_DEFENSE_KEYS.flatMap(defenseKey => {
    const config = leafDefenseConfigs[defenseKey];
    const initiatives = retentionDefenseInitiatives(scenario, defenseKey);
    if (!config || !initiatives.length) return [];
    return retentionInitiativeRevenueImpacts(scenario, defenseKey).slice(0, initiatives.length).map((impact, index) => ({
      id: `${defenseKey}:${index}`,
      defenseKey,
      assumption: displayLabel(config.shortTitle),
      name: impact.name,
      teams: initiatives[index]?.teams || [],
      forecastTotal: impact.forecastTotal,
      finalYear: impact.finalYear,
      yearly: impact.yearly,
      color: impact.color,
    }));
  }).sort((left, right) => right.forecastTotal - left.forecastTotal);
}

function initiativeFilterValues() {
  return {
    team: document.getElementById("initiative-team-filter")?.value || "all",
    assumption: document.getElementById("initiative-assumption-filter")?.value || "all",
  };
}

function filteredInitiativeRows(rows = initiativeImpactRows()) {
  const filters = initiativeFilterValues();
  return rows.filter(row => {
    const teamMatch = filters.team === "all"
      || (filters.team === "untagged" ? row.teams.length === 0 : row.teams.includes(filters.team));
    const assumptionMatch = filters.assumption === "all" || row.defenseKey === filters.assumption;
    return teamMatch && assumptionMatch;
  });
}

function renderInitiativeFilters(rows) {
  const teamSelect = document.getElementById("initiative-team-filter");
  const assumptionSelect = document.getElementById("initiative-assumption-filter");
  if (!teamSelect || !assumptionSelect) return;
  const currentTeam = teamSelect.value || "all";
  const currentAssumption = assumptionSelect.value || "all";
  const usedTeams = new Set(rows.flatMap(row => row.teams));
  const hasUntagged = rows.some(row => !row.teams.length);
  teamSelect.innerHTML = [
    `<option value="all">All teams</option>`,
    ...(hasUntagged ? [`<option value="untagged">Untagged</option>`] : []),
    ...INITIATIVE_TEAMS
      .filter(team => usedTeams.has(team))
      .map(team => `<option value="${escapeHtml(team)}">${escapeHtml(team)}</option>`),
  ].join("");
  const teamValues = Array.from(teamSelect.options).map(option => option.value);
  teamSelect.value = teamValues.includes(currentTeam) ? currentTeam : "all";

  const usedAssumptions = new Set(rows.map(row => row.defenseKey));
  assumptionSelect.innerHTML = [
    `<option value="all">All assumptions</option>`,
    ...LEAF_DEFENSE_KEYS
      .filter(key => usedAssumptions.has(key))
      .map(key => `<option value="${key}">${escapeHtml(displayLabel(leafDefenseConfigs[key]?.shortTitle || key))}</option>`),
  ].join("");
  const assumptionValues = Array.from(assumptionSelect.options).map(option => option.value);
  assumptionSelect.value = assumptionValues.includes(currentAssumption) ? currentAssumption : "all";
}

function initiativeTeamRows(rows) {
  const totals = new Map();
  rows.forEach(row => {
    const teams = row.teams.length ? row.teams : ["Untagged"];
    const forecastShare = row.forecastTotal / teams.length;
    const finalShare = row.finalYear / teams.length;
    teams.forEach(team => {
      const existing = totals.get(team) || { team, forecastTotal: 0, finalYear: 0 };
      existing.forecastTotal += forecastShare;
      existing.finalYear += finalShare;
      totals.set(team, existing);
    });
  });
  return Array.from(totals.values()).sort((left, right) => right.forecastTotal - left.forecastTotal);
}

function renderInitiativeCharts(rows) {
  const impactChart = state.outputCharts.initiativeImpact;
  const teamChart = state.outputCharts.initiativeTeamImpact;
  const chartRows = rows.slice(0, 12).reverse();
  if (impactChart) {
    impactChart.setOption({
      animation: false,
      tooltip: {
        trigger: "axis",
        axisPointer: { type: "shadow" },
        formatter: params => {
          const item = params[0];
          return `${escapeHtml(item.name)}<br/>${item.marker} 2026-2029 Impact: ${formatCurrency(Number(item.value), 0)}`;
        },
      },
      grid: { left: 150, right: 20, top: 12, bottom: 28 },
      xAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
      yAxis: {
        type: "category",
        data: chartRows.map(row => row.name),
        axisLabel: { width: 132, overflow: "truncate" },
      },
      series: [{
        type: "bar",
        data: chartRows.map(row => ({
          value: row.forecastTotal,
          itemStyle: { color: row.color },
        })),
        barMaxWidth: 18,
      }],
    }, true);
  }

  const teamRows = initiativeTeamRows(rows).slice(0, 10).reverse();
  if (teamChart) {
    teamChart.setOption({
      animation: false,
      tooltip: {
        trigger: "axis",
        axisPointer: { type: "shadow" },
        formatter: params => {
          const item = params[0];
          return `${escapeHtml(item.name)}<br/>${item.marker} 2026-2029 Impact: ${formatCurrency(Number(item.value), 0)}`;
        },
      },
      grid: { left: 140, right: 20, top: 12, bottom: 28 },
      xAxis: { type: "value", axisLabel: { formatter: formatCompactCurrency } },
      yAxis: {
        type: "category",
        data: teamRows.map(row => displayLabel(row.team)),
        axisLabel: { width: 122, overflow: "truncate" },
      },
      series: [{
        type: "bar",
        data: teamRows.map((row, index) => ({
          value: row.forecastTotal,
          itemStyle: { color: metricTreeScenarioPalette[index % metricTreeScenarioPalette.length] },
        })),
        barMaxWidth: 18,
      }],
    }, true);
  }
}

function renderInitiativesModule() {
  const allRows = initiativeImpactRows();
  renderInitiativeFilters(allRows);
  const rows = filteredInitiativeRows(allRows);
  const teamRows = initiativeTeamRows(rows);
  const count = document.getElementById("initiative-count");
  const totalImpact = document.getElementById("initiative-total-impact");
  const finalImpact = document.getElementById("initiative-final-impact");
  const teamCount = document.getElementById("initiative-team-count");
  const empty = document.getElementById("initiative-empty-state");
  const table = document.getElementById("initiative-table");
  if (count) count.textContent = rows.length.toLocaleString("en-US");
  if (totalImpact) totalImpact.textContent = formatCurrency(rows.reduce((sum, row) => sum + row.forecastTotal, 0), 0);
  if (finalImpact) finalImpact.textContent = formatCurrency(rows.reduce((sum, row) => sum + row.finalYear, 0), 0);
  if (teamCount) teamCount.textContent = teamRows.length.toLocaleString("en-US");
  if (empty) empty.hidden = allRows.length > 0;
  if (table) {
    table.hidden = rows.length === 0;
    table.innerHTML = rows.length
      ? `
        <thead>
          <tr>
            <th>Rank</th>
            <th>Initiative</th>
            <th>Assumption</th>
            <th>Teams</th>
            <th>2026-2029 Impact</th>
            <th>2029 Impact</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((row, index) => `
            <tr>
              <td>${index + 1}</td>
              <td><strong>${escapeHtml(row.name)}</strong></td>
              <td>${escapeHtml(row.assumption)}</td>
              <td>
                <div class="initiative-table-tags">
                  ${(row.teams.length ? row.teams : ["Untagged"]).map(team => `<span>${escapeHtml(displayLabel(team))}</span>`).join("")}
                </div>
              </td>
              <td>${formatCurrency(row.forecastTotal, 0)}</td>
              <td>${formatCurrency(row.finalYear, 0)}</td>
              <td><button class="drilldown-button" data-initiative-defense-key="${row.defenseKey}" type="button">Defend</button></td>
            </tr>
          `).join("")}
        </tbody>
      `
      : "";
  }
  renderInitiativeCharts(rows);
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
  const formatTooltipValue = value => value === null || value === undefined || !Number.isFinite(Number(value))
    ? "-"
    : tooltipFormatter(value);
  return {
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => [
        tooltipHeader(params[0]?.axisValue),
        ...params.map(item => `${item.marker} ${item.seriesName}: ${formatTooltipValue(item.value)}`),
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

function totalCostValues(scenario) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => Number(scenario.costs.labor[index] || 0) + Number(scenario.costs.nonLabor[index] || 0));
}

function revenueTotalTopDownValues(scenario = activeScenario()) {
  normalizeRevenuePaths(scenario);
  return scenario.revenuePaths.total.topDown;
}

function ltrRevenueTopDownValues(scenario = activeScenario(), outputs = calculateOutputs(scenario)) {
  normalizeRevenuePaths(scenario);
  const plan = scenario.revenuePaths.ltr;
  return YEARS.map((_year, index) => {
    const value = Number(plan.topDown[index]);
    return Number.isFinite(value) ? Math.max(0, value) : outputs.revenue[index];
  });
}

function ltrRevenueBottomUpValues(_scenario = activeScenario(), outputs = calculateOutputs(_scenario)) {
  return outputs.revenue;
}

function revenueCategoryValues(scenario, key, outputs = calculateOutputs(scenario)) {
  if (key === "ltr") return ltrRevenueTopDownValues(scenario, outputs);
  if (key === "growth") return growthRevenueTopDownValues(scenario);
  if (key === "str") return strRevenueTopDownValues(scenario);
  if (key === "other") return otherRevenueTopDownValues(scenario);
  return totalRevenueValues(outputs, scenario);
}

function revenueBottomUpValues(scenario = activeScenario(), outputs = calculateOutputs(scenario)) {
  const ltr = revenueCategoryValues(scenario, "ltr", outputs);
  const growth = growthRevenueTopDownValues(scenario);
  const str = strRevenueTopDownValues(scenario);
  const other = otherRevenueTopDownValues(scenario);
  return YEARS.map((_year, index) => {
    return Number(ltr[index] || 0)
      + Number(growth[index] || 0)
      + Number(str[index] || 0)
      + Number(other[index] || 0);
  });
}

function totalRevenueValues(outputs, scenario) {
  return revenueBottomUpValues(scenario, outputs);
}

function profitValues(outputs, scenario) {
  const costs = totalCostValues(scenario);
  const revenue = totalRevenueValues(outputs, scenario);
  return revenue.map((value, index) => Number(value || 0) - Number(costs[index] || 0));
}

function laborFteTotalValues(scenario) {
  normalizeCostPlan(scenario);
  return YEARS.map((_year, index) => {
    return LABOR_DEPARTMENT_KEYS.reduce((sum, key) => {
      return sum + Number(scenario.costs.laborFte?.[key]?.[index] || 0);
    }, 0);
  });
}

function laborCostPerFteAggregateValues(scenario) {
  normalizeCostPlan(scenario);
  const fteValues = laborFteTotalValues(scenario);
  return YEARS.map((_year, index) => {
    const fte = Number(fteValues[index] || 0);
    return fte > 0 ? Number(scenario.costs.labor[index] || 0) / fte : 0;
  });
}

function revenueGrowthPercentValues(revenueValues) {
  return revenueValues.map((value, index) => {
    if (index === 0) return null;
    const prior = Number(revenueValues[index - 1] || 0);
    if (prior <= 0) return null;
    return ((Number(value || 0) / prior) - 1) * 100;
  });
}

function profitMarginPercentValues(revenueValues, profitValueList) {
  return revenueValues.map((value, index) => {
    const revenue = Number(value || 0);
    if (revenue <= 0) return null;
    return (Number(profitValueList[index] || 0) / revenue) * 100;
  });
}

function ruleOf40Values(revenueValues, profitValueList) {
  const revenueGrowth = revenueGrowthPercentValues(revenueValues);
  const profitMargin = profitMarginPercentValues(revenueValues, profitValueList);
  return revenueGrowth.map((value, index) => {
    const margin = profitMargin[index];
    return Number.isFinite(value) && Number.isFinite(margin) ? value + margin : null;
  });
}

function formatSaasPercent(value, decimals = 1) {
  if (isAnonymizedView()) return anonymizedMetricValue();
  return Number.isFinite(value) ? `${trimNumber(value, decimals)}%` : "-";
}

function fullyLoadedCacValues(totalCostValueList, newCustomerValueList) {
  return YEARS.map((_year, index) => {
    const newCustomers = Number(newCustomerValueList[index] || 0);
    if (newCustomers <= 0) return null;
    return Number(totalCostValueList[index] || 0) / newCustomers;
  });
}

function revenuePerFteValues(revenueValues, fteValues) {
  return YEARS.map((_year, index) => {
    const fte = Number(fteValues[index] || 0);
    if (fte <= 0) return null;
    return Number(revenueValues[index] || 0) / fte;
  });
}

function revenuePerProxyCustomerValues(revenueValues, customerValues) {
  return YEARS.map((_year, index) => {
    const revenue = revenueValues[index];
    const customers = Number(customerValues[index] || 0);
    if (revenue === null || revenue === undefined || customers <= 0) return null;
    return Number(revenue || 0) / customers;
  });
}

function growthRevenuePlan(scenario = activeScenario()) {
  normalizeRevenuePaths(scenario);
  return scenario.revenuePaths.growth;
}

function growthRevenueTopDownValues(scenario = activeScenario()) {
  return growthRevenuePlan(scenario).topDown;
}

function growthRevenueProxyCustomers(scenario = activeScenario()) {
  return driverValues(scenario, "newCustomers");
}

function growthRevenueRevPerNewCustomerValues(scenario = activeScenario()) {
  return growthRevenuePlan(scenario).revPerNewPayingCustomer;
}

function growthRevenueBottomUpValues(scenario = activeScenario()) {
  const customers = growthRevenueProxyCustomers(scenario);
  const revPerCustomer = growthRevenueRevPerNewCustomerValues(scenario);
  return YEARS.map((_year, index) => {
    const customerValue = Number(customers[index] || 0);
    const revValue = Number(revPerCustomer[index] || 0);
    return customerValue > 0 ? customerValue * revValue : null;
  });
}

function growthRevenueControlPoints(scenario = activeScenario()) {
  return growthRevenuePlan(scenario).controlPoints;
}

function controlledGrowthRevenuePairs(key, values, scale = 1) {
  const stored = growthRevenueControlPoints()[key];
  if (Array.isArray(stored) && stored.length >= 2) {
    return stored
      .map(year => {
        const index = YEARS.indexOf(Number(year));
        if (index < 0) return null;
        const value = values[index];
        return value === null || value === undefined ? null : [Number(year), value / scale];
      })
      .filter(Boolean);
  }
  return YEARS
    .map((year, index) => {
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / scale];
    })
    .filter(Boolean);
}

function rememberGrowthRevenueControlPoints(key, pairs) {
  const years = (pairs || [])
    .map(point => Number(point?.[0]))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
  if (years.length >= 2) {
    growthRevenueControlPoints()[key] = Array.from(new Set(years));
  } else {
    delete growthRevenueControlPoints()[key];
  }
}

function growthRevenueTopMatchesBottomUp(scenario = activeScenario()) {
  const topDown = growthRevenueTopDownValues(scenario);
  const bottomUp = growthRevenueBottomUpValues(scenario);
  return YEARS.every((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return true;
    const topValue = Number(topDown[index]);
    const bottomValue = Number(bottomUp[index]);
    return Number.isFinite(topValue) && Number.isFinite(bottomValue) && Math.abs(topValue - bottomValue) < 1;
  });
}

function growthRevenueTopDownPairs(scenario = activeScenario()) {
  return controlledGrowthRevenuePairs("topDown", growthRevenueTopDownValues(scenario), 1000000);
}

function growthRevenueBottomUpPairs(scenario = activeScenario()) {
  return controlledGrowthRevenuePairs("bottomUp", growthRevenueBottomUpValues(scenario), 1000000);
}

function growthRevenueRevPerNewCustomerPairs(scenario = activeScenario()) {
  return controlledGrowthRevenuePairs("revPerNewPayingCustomer", growthRevenueRevPerNewCustomerValues(scenario), 1);
}

function growthRevenueTopDownChartLines() {
  return [
    {
      name: "Bottom Up",
      color: BOTTOM_UP_COLOR,
      editable: true,
      editDomain: {
        moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: growthRevenueBottomUpPairs(activeScenario()),
    },
    {
      name: "Top Down",
      color: "#98a2b3",
      editable: false,
      data: growthRevenueTopDownPairs(activeScenario()),
    },
  ];
}

function growthRevenueRevPerNewCustomerChartLines() {
  return [{
    name: "Rev / New Paying Customer",
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: growthRevenueRevPerNewCustomerPairs(activeScenario()),
  }];
}

function styleGrowthReferenceSeries(chart) {
  const option = chart.chart.getOption();
  const series = (option.series || []).map(seriesItem => {
    if (String(seriesItem.name || "") !== "Top Down") return seriesItem;
    return {
      ...seriesItem,
      silent: true,
      showSymbol: false,
      symbolSize: 0,
      z: 2,
      itemStyle: { ...(seriesItem.itemStyle || {}), color: "#98a2b3" },
      lineStyle: { ...(seriesItem.lineStyle || {}), color: "#98a2b3", width: 3, opacity: 1 },
    };
  });
  chart.chart.setOption({ series }, false);
}

function renderGrowthRevenueProxyCustomersEchart() {
  const chart = state.outputCharts.growthRevenueProxyCustomers;
  if (!chart) return;
  const scenario = activeScenario();
  const comparison = compareScenario();
  const scenarioColor = metricTreeScenarioColor(scenario);
  const comparisonColor = metricTreeScenarioColor(comparison);
  const series = [
    metricTreeSeries(displayScenarioName(scenario), growthRevenueProxyCustomers(scenario), scenarioColor),
  ];
  if (comparison) {
    series.push(metricTreeSeries(displayScenarioName(comparison), growthRevenueProxyCustomers(comparison), comparisonColor));
  }
  chart.setOption(metricTreeChartOption(series), true);
}

function renderGrowthRevenueDrilldown() {
  renderGrowthRevenueProxyCustomersEchart();
  const matched = growthRevenueTopMatchesBottomUp();
  const topToBottomButton = document.getElementById("set-growth-revenue-top-to-bottom");
  const bottomToTopButton = document.getElementById("set-growth-revenue-bottom-to-top");
  if (topToBottomButton) topToBottomButton.disabled = matched;
  if (bottomToTopButton) bottomToTopButton.disabled = matched;
}

function syncGrowthRevenueCharts({ excludeChart = null } = {}) {
  state.syncingGrowthCharts = true;
  const topChart = state.growthCharts.growthRevenueTopDown;
  if (topChart && topChart !== excludeChart) {
    topChart.lines = safeNapkinLines(growthRevenueTopDownChartLines());
    topChart._refreshChart();
    styleGrowthReferenceSeries(topChart);
  }
  const revPerChart = state.growthCharts.growthRevenueRevPerNewCustomer;
  if (revPerChart && revPerChart !== excludeChart) {
    revPerChart.lines = safeNapkinLines(growthRevenueRevPerNewCustomerChartLines());
    revPerChart._refreshChart();
  }
  state.syncingGrowthCharts = false;
  renderGrowthRevenueDrilldown();
}

function setGrowthRevenueBottomToTop() {
  if (growthRevenueTopMatchesBottomUp()) return;
  pushUndoSnapshot();
  const scenario = activeScenario();
  const plan = growthRevenuePlan(scenario);
  const bottomUp = growthRevenueBottomUpValues(scenario);
  YEARS.forEach((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return;
    plan.topDown[index] = bottomUp[index];
  });
  plan.controlPoints.topDown = YEARS.filter((year, index) => plan.topDown[index] !== null && plan.topDown[index] !== undefined);
  saveScenarios();
  syncGrowthRevenueCharts();
  renderAll();
}

function setGrowthRevenueTopToBottom() {
  if (growthRevenueTopMatchesBottomUp()) return;
  pushUndoSnapshot();
  const scenario = activeScenario();
  const plan = growthRevenuePlan(scenario);
  YEARS.forEach((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return;
    const topValue = Number(plan.topDown[index]);
    const customers = Number(growthRevenueProxyCustomers(scenario)[index] || 0);
    if (Number.isFinite(topValue) && customers > 0) {
      plan.revPerNewPayingCustomer[index] = topValue / customers;
    }
  });
  plan.controlPoints.revPerNewPayingCustomer = YEARS.slice();
  plan.controlPoints.bottomUp = YEARS.slice();
  saveScenarios();
  syncGrowthRevenueCharts();
  renderAll();
}

function initGrowthRevenueCharts() {
  const topChart = new NapkinChart(
    "growth-revenue-topdown-chart",
    safeNapkinLines(growthRevenueTopDownChartLines()),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 15, axisLabel: { formatter: value => formatCompactCurrency(Number(value) * 1000000) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = items
            .filter(item => item && item.seriesType === "line")
            .map(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(rawValue) * 1000000, 0)}`;
            });
          return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );
  topChart.windowStartX = YEARS[0];
  topChart.windowEndX = YEARS[YEARS.length - 1];
  topChart.globalMaxX = YEARS[YEARS.length - 1];
  styleGrowthReferenceSeries(topChart);
  topChart._appEditSnapshot = null;
  topChart._appEditCommitted = false;
  topChart.chart.getZr().on("mousedown", () => {
    if (state.syncingGrowthCharts) return;
    topChart._appEditSnapshot = snapshotState();
    topChart._appEditCommitted = false;
  });
  topChart.onDataChanged = () => {
    if (state.syncingGrowthCharts) return;
    if (topChart._appEditSnapshot && !topChart._appEditCommitted) {
      pushUndoSnapshot(topChart._appEditSnapshot);
      topChart._appEditCommitted = true;
    }
    const plan = growthRevenuePlan(activeScenario());
    const customers = growthRevenueProxyCustomers(activeScenario());
    YEARS.forEach((year, index) => {
      if (year < FIRST_FORECAST_YEAR) return;
      const value = interpolateNapkinLineValue(topChart.lines[0].data, year);
      const customerValue = Number(customers[index] || 0);
      if (Number.isFinite(value) && customerValue > 0) {
        plan.revPerNewPayingCustomer[index] = Math.max(0, (value * 1000000) / customerValue);
      }
    });
    rememberGrowthRevenueControlPoints("bottomUp", topChart.lines[0].data);
    rememberGrowthRevenueControlPoints("revPerNewPayingCustomer", topChart.lines[0].data);
    saveScenarios();
    syncGrowthRevenueCharts({ excludeChart: topChart });
    renderAll({ syncNewCustomerEditors: false });
  };
  state.growthCharts.growthRevenueTopDown = topChart;

  const revPerChart = new NapkinChart(
    "growth-revenue-rev-per-new-customer-chart",
    safeNapkinLines(growthRevenueRevPerNewCustomerChartLines()),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 25, axisLabel: { formatter: value => formatCurrency(Number(value), 0) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = items
            .filter(item => item && item.seriesType === "line")
            .map(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(rawValue), 2)}`;
            });
          return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );
  revPerChart.windowStartX = YEARS[0];
  revPerChart.windowEndX = YEARS[YEARS.length - 1];
  revPerChart.globalMaxX = YEARS[YEARS.length - 1];
  revPerChart._appEditSnapshot = null;
  revPerChart._appEditCommitted = false;
  revPerChart.chart.getZr().on("mousedown", () => {
    if (state.syncingGrowthCharts) return;
    revPerChart._appEditSnapshot = snapshotState();
    revPerChart._appEditCommitted = false;
  });
  revPerChart.onDataChanged = () => {
    if (state.syncingGrowthCharts) return;
    if (revPerChart._appEditSnapshot && !revPerChart._appEditCommitted) {
      pushUndoSnapshot(revPerChart._appEditSnapshot);
      revPerChart._appEditCommitted = true;
    }
    const plan = growthRevenuePlan(activeScenario());
    YEARS.forEach((year, index) => {
      if (year < FIRST_FORECAST_YEAR) return;
      const value = interpolateNapkinLineValue(revPerChart.lines[0].data, year);
      if (Number.isFinite(value)) plan.revPerNewPayingCustomer[index] = Math.max(0, value);
    });
    rememberGrowthRevenueControlPoints("revPerNewPayingCustomer", revPerChart.lines[0].data);
    rememberGrowthRevenueControlPoints("bottomUp", revPerChart.lines[0].data);
    saveScenarios();
    syncGrowthRevenueCharts({ excludeChart: revPerChart });
    renderAll({ syncNewCustomerEditors: false, syncRevenueDrilldown: false });
  };
  state.growthCharts.growthRevenueRevPerNewCustomer = revPerChart;
}

function strRevenuePlan(scenario = activeScenario()) {
  normalizeRevenuePaths(scenario);
  return scenario.revenuePaths.str;
}

function strRevenueTopDownValues(scenario = activeScenario()) {
  return strRevenuePlan(scenario).topDown;
}

function strRevenuePropertiesValues(scenario = activeScenario()) {
  return strRevenuePlan(scenario).numStrProperties;
}

function strRevenueRevPerPropertyValues(scenario = activeScenario()) {
  return strRevenuePlan(scenario).revPerStrProperty;
}

function strRevenueBottomUpValues(scenario = activeScenario()) {
  const properties = strRevenuePropertiesValues(scenario);
  const revenuePerProperty = strRevenueRevPerPropertyValues(scenario);
  return YEARS.map((_year, index) => {
    const propertyValue = Number(properties[index] || 0);
    const revValue = Number(revenuePerProperty[index] || 0);
    return propertyValue > 0 ? propertyValue * revValue : 0;
  });
}

function strRevenueControlPoints(scenario = activeScenario()) {
  return strRevenuePlan(scenario).controlPoints;
}

function controlledStrRevenuePairs(key, values, scale = 1) {
  const stored = strRevenueControlPoints()[key];
  if (Array.isArray(stored) && stored.length >= 2) {
    return stored
      .map(year => {
        const index = YEARS.indexOf(Number(year));
        if (index < 0) return null;
        const value = values[index];
        return value === null || value === undefined ? null : [Number(year), value / scale];
      })
      .filter(Boolean);
  }
  return YEARS
    .map((year, index) => {
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / scale];
    })
    .filter(Boolean);
}

function rememberStrRevenueControlPoints(key, pairs) {
  const years = (pairs || [])
    .map(point => Number(point?.[0]))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
  if (years.length >= 2) {
    strRevenueControlPoints()[key] = Array.from(new Set(years));
  } else {
    delete strRevenueControlPoints()[key];
  }
}

function strRevenueTopMatchesBottomUp(scenario = activeScenario()) {
  const topDown = strRevenueTopDownValues(scenario);
  const bottomUp = strRevenueBottomUpValues(scenario);
  return YEARS.every((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return true;
    const topValue = Number(topDown[index]);
    const bottomValue = Number(bottomUp[index]);
    return Number.isFinite(topValue) && Number.isFinite(bottomValue) && Math.abs(topValue - bottomValue) < 1;
  });
}

function revenueTopMatchesBottomUp(scenario = activeScenario()) {
  const outputs = calculateOutputs(scenario);
  const topDown = revenueTotalTopDownValues(scenario);
  const bottomUp = revenueBottomUpValues(scenario, outputs);
  return YEARS.every((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return true;
    return Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0)) < 1;
  });
}

function strRevenueTopDownPairs(scenario = activeScenario()) {
  return controlledStrRevenuePairs("topDown", strRevenueTopDownValues(scenario), 1000000);
}

function strRevenueBottomUpPairs(scenario = activeScenario()) {
  return controlledStrRevenuePairs("bottomUp", strRevenueBottomUpValues(scenario), 1000000);
}

function strRevenuePropertiesPairs(scenario = activeScenario()) {
  return controlledStrRevenuePairs("numStrProperties", strRevenuePropertiesValues(scenario), 1);
}

function strRevenueRevPerPropertyPairs(scenario = activeScenario()) {
  return controlledStrRevenuePairs("revPerStrProperty", strRevenueRevPerPropertyValues(scenario), 1);
}

function strRevenueTopDownChartLines() {
  return [
    {
      name: "Bottom Up",
      color: BOTTOM_UP_COLOR,
      editable: true,
      editDomain: {
        moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: strRevenueBottomUpPairs(activeScenario()),
    },
    {
      name: "Top Down",
      color: "#98a2b3",
      editable: false,
      data: strRevenueTopDownPairs(activeScenario()),
    },
  ];
}

function strRevenuePropertiesChartLines() {
  return [{
    name: "Num STR Properties",
    color: "#6fa76b",
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: strRevenuePropertiesPairs(activeScenario()),
  }];
}

function strRevenueRevPerPropertyChartLines() {
  return [{
    name: "Rev / STR Property",
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: strRevenueRevPerPropertyPairs(activeScenario()),
  }];
}

function renderStrRevenueDrilldown() {
  const matched = strRevenueTopMatchesBottomUp();
  const topToBottomButton = document.getElementById("set-str-revenue-top-to-bottom");
  const bottomToTopButton = document.getElementById("set-str-revenue-bottom-to-top");
  if (topToBottomButton) topToBottomButton.disabled = matched;
  if (bottomToTopButton) bottomToTopButton.disabled = matched;
}

function syncStrRevenueCharts({ excludeChart = null } = {}) {
  state.syncingStrCharts = true;
  const topChart = state.strCharts.strRevenueTopDown;
  if (topChart && topChart !== excludeChart) {
    topChart.lines = safeNapkinLines(strRevenueTopDownChartLines());
    topChart._refreshChart();
    styleGrowthReferenceSeries(topChart);
  }
  const propertiesChart = state.strCharts.strRevenueProperties;
  if (propertiesChart && propertiesChart !== excludeChart) {
    propertiesChart.lines = safeNapkinLines(strRevenuePropertiesChartLines());
    propertiesChart._refreshChart();
  }
  const revPerChart = state.strCharts.strRevenueRevPerProperty;
  if (revPerChart && revPerChart !== excludeChart) {
    revPerChart.lines = safeNapkinLines(strRevenueRevPerPropertyChartLines());
    revPerChart._refreshChart();
  }
  state.syncingStrCharts = false;
  renderStrRevenueDrilldown();
}

function setStrRevenueBottomToTop() {
  if (strRevenueTopMatchesBottomUp()) return;
  pushUndoSnapshot();
  const scenario = activeScenario();
  const plan = strRevenuePlan(scenario);
  const bottomUp = strRevenueBottomUpValues(scenario);
  YEARS.forEach((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return;
    plan.topDown[index] = bottomUp[index];
  });
  plan.controlPoints.topDown = YEARS.filter((year, index) => plan.topDown[index] !== null && plan.topDown[index] !== undefined);
  saveScenarios();
  syncStrRevenueCharts();
  renderAll();
}

function setStrRevenueTopToBottom() {
  if (strRevenueTopMatchesBottomUp()) return;
  pushUndoSnapshot();
  const scenario = activeScenario();
  const plan = strRevenuePlan(scenario);
  YEARS.forEach((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return;
    const topValue = Number(plan.topDown[index]);
    const properties = Number(strRevenuePropertiesValues(scenario)[index] || 0);
    if (Number.isFinite(topValue) && properties > 0) {
      plan.revPerStrProperty[index] = topValue / properties;
    }
  });
  plan.controlPoints.revPerStrProperty = YEARS.slice();
  plan.controlPoints.bottomUp = YEARS.slice();
  saveScenarios();
  syncStrRevenueCharts();
  renderAll();
}

function initStrRevenueCharts() {
  const topChart = new NapkinChart(
    "str-revenue-topdown-chart",
    safeNapkinLines(strRevenueTopDownChartLines()),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 10, axisLabel: { formatter: value => formatCompactCurrency(Number(value) * 1000000) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = items
            .filter(item => item && item.seriesType === "line")
            .map(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(rawValue) * 1000000, 0)}`;
            });
          return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );
  topChart.windowStartX = YEARS[0];
  topChart.windowEndX = YEARS[YEARS.length - 1];
  topChart.globalMaxX = YEARS[YEARS.length - 1];
  styleGrowthReferenceSeries(topChart);
  topChart._appEditSnapshot = null;
  topChart._appEditCommitted = false;
  topChart.chart.getZr().on("mousedown", () => {
    if (state.syncingStrCharts) return;
    topChart._appEditSnapshot = snapshotState();
    topChart._appEditCommitted = false;
  });
  topChart.onDataChanged = () => {
    if (state.syncingStrCharts) return;
    if (topChart._appEditSnapshot && !topChart._appEditCommitted) {
      pushUndoSnapshot(topChart._appEditSnapshot);
      topChart._appEditCommitted = true;
    }
    const plan = strRevenuePlan(activeScenario());
    const properties = strRevenuePropertiesValues(activeScenario());
    YEARS.forEach((year, index) => {
      if (year < FIRST_FORECAST_YEAR) return;
      const value = interpolateNapkinLineValue(topChart.lines[0].data, year);
      const propertyValue = Number(properties[index] || 0);
      if (Number.isFinite(value) && propertyValue > 0) {
        plan.revPerStrProperty[index] = Math.max(0, (value * 1000000) / propertyValue);
      }
    });
    rememberStrRevenueControlPoints("bottomUp", topChart.lines[0].data);
    rememberStrRevenueControlPoints("revPerStrProperty", topChart.lines[0].data);
    saveScenarios();
    syncStrRevenueCharts({ excludeChart: topChart });
    renderAll({ syncNewCustomerEditors: false });
  };
  state.strCharts.strRevenueTopDown = topChart;

  const propertiesChart = new NapkinChart(
    "str-revenue-properties-chart",
    safeNapkinLines(strRevenuePropertiesChartLines()),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 250, axisLabel: { formatter: value => formatIntegerMetric(Number(value)) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = items
            .filter(item => item && item.seriesType === "line")
            .map(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatIntegerMetric(Number(rawValue))}`;
            });
          return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );
  propertiesChart.windowStartX = YEARS[0];
  propertiesChart.windowEndX = YEARS[YEARS.length - 1];
  propertiesChart.globalMaxX = YEARS[YEARS.length - 1];
  propertiesChart._appEditSnapshot = null;
  propertiesChart._appEditCommitted = false;
  propertiesChart.chart.getZr().on("mousedown", () => {
    if (state.syncingStrCharts) return;
    propertiesChart._appEditSnapshot = snapshotState();
    propertiesChart._appEditCommitted = false;
  });
  propertiesChart.onDataChanged = () => {
    if (state.syncingStrCharts) return;
    if (propertiesChart._appEditSnapshot && !propertiesChart._appEditCommitted) {
      pushUndoSnapshot(propertiesChart._appEditSnapshot);
      propertiesChart._appEditCommitted = true;
    }
    const plan = strRevenuePlan(activeScenario());
    YEARS.forEach((year, index) => {
      if (year < FIRST_FORECAST_YEAR) return;
      const value = interpolateNapkinLineValue(propertiesChart.lines[0].data, year);
      if (Number.isFinite(value)) plan.numStrProperties[index] = Math.max(0, value);
    });
    rememberStrRevenueControlPoints("numStrProperties", propertiesChart.lines[0].data);
    rememberStrRevenueControlPoints("bottomUp", propertiesChart.lines[0].data);
    saveScenarios();
    syncStrRevenueCharts({ excludeChart: propertiesChart });
    renderAll({ syncNewCustomerEditors: false });
  };
  state.strCharts.strRevenueProperties = propertiesChart;

  const revPerChart = new NapkinChart(
    "str-revenue-rev-per-property-chart",
    safeNapkinLines(strRevenueRevPerPropertyChartLines()),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 150000, axisLabel: { formatter: value => formatCompactCurrency(Number(value)) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = items
            .filter(item => item && item.seriesType === "line")
            .map(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(rawValue), 0)}`;
            });
          return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );
  revPerChart.windowStartX = YEARS[0];
  revPerChart.windowEndX = YEARS[YEARS.length - 1];
  revPerChart.globalMaxX = YEARS[YEARS.length - 1];
  revPerChart._appEditSnapshot = null;
  revPerChart._appEditCommitted = false;
  revPerChart.chart.getZr().on("mousedown", () => {
    if (state.syncingStrCharts) return;
    revPerChart._appEditSnapshot = snapshotState();
    revPerChart._appEditCommitted = false;
  });
  revPerChart.onDataChanged = () => {
    if (state.syncingStrCharts) return;
    if (revPerChart._appEditSnapshot && !revPerChart._appEditCommitted) {
      pushUndoSnapshot(revPerChart._appEditSnapshot);
      revPerChart._appEditCommitted = true;
    }
    const plan = strRevenuePlan(activeScenario());
    YEARS.forEach((year, index) => {
      if (year < FIRST_FORECAST_YEAR) return;
      const value = interpolateNapkinLineValue(revPerChart.lines[0].data, year);
      if (Number.isFinite(value)) plan.revPerStrProperty[index] = Math.max(0, value);
    });
    rememberStrRevenueControlPoints("revPerStrProperty", revPerChart.lines[0].data);
    rememberStrRevenueControlPoints("bottomUp", revPerChart.lines[0].data);
    saveScenarios();
    syncStrRevenueCharts({ excludeChart: revPerChart });
    renderAll({ syncNewCustomerEditors: false });
  };
  state.strCharts.strRevenueRevPerProperty = revPerChart;
}

function otherRevenuePlan(scenario = activeScenario()) {
  normalizeRevenuePaths(scenario);
  return scenario.revenuePaths.other;
}

function otherRevenueTopDownValues(scenario = activeScenario()) {
  return otherRevenuePlan(scenario).topDown;
}

function otherRevenueControlPoints(scenario = activeScenario()) {
  return otherRevenuePlan(scenario).controlPoints;
}

function controlledOtherRevenuePairs(key, values, scale = 1) {
  const stored = otherRevenueControlPoints()[key];
  if (Array.isArray(stored) && stored.length >= 2) {
    return stored
      .map(year => {
        const index = YEARS.indexOf(Number(year));
        if (index < 0) return null;
        const value = values[index];
        return value === null || value === undefined ? null : [Number(year), value / scale];
      })
      .filter(Boolean);
  }
  return YEARS
    .map((year, index) => {
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / scale];
    })
    .filter(Boolean);
}

function rememberOtherRevenueControlPoints(key, pairs) {
  const years = (pairs || [])
    .map(point => Number(point?.[0]))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
  if (years.length >= 2) {
    otherRevenueControlPoints()[key] = Array.from(new Set(years));
  } else {
    delete otherRevenueControlPoints()[key];
  }
}

function otherRevenuePairs(scenario = activeScenario()) {
  return controlledOtherRevenuePairs("topDown", otherRevenueTopDownValues(scenario), 1000000);
}

function otherRevenueChartLines() {
  return [{
    name: "Other Revenue",
    color: "#a96c50",
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: otherRevenuePairs(activeScenario()),
  }];
}

function syncOtherRevenueCharts({ excludeChart = null } = {}) {
  state.syncingOtherRevenueCharts = true;
  const chart = state.otherRevenueCharts.otherRevenue;
  if (chart && chart !== excludeChart) {
    chart.lines = safeNapkinLines(otherRevenueChartLines());
    chart._refreshChart();
  }
  state.syncingOtherRevenueCharts = false;
}

function initOtherRevenueCharts() {
  const chart = new NapkinChart(
    "other-revenue-chart",
    safeNapkinLines(otherRevenueChartLines()),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: 3, axisLabel: { formatter: value => formatCompactCurrency(Number(value) * 1000000) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: {
        trigger: "axis",
        formatter: params => {
          const items = Array.isArray(params) ? params : [params];
          const rows = items
            .filter(item => item && item.seriesType === "line")
            .map(item => {
              const data = Array.isArray(item.data) ? item.data : item.value;
              const rawValue = Array.isArray(data) ? data[1] : item.value;
              return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(rawValue) * 1000000, 0)}`;
            });
          return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
        },
      },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingOtherRevenueCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingOtherRevenueCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    const plan = otherRevenuePlan(activeScenario());
    YEARS.forEach((year, index) => {
      if (year < FIRST_FORECAST_YEAR) return;
      const value = interpolateNapkinLineValue(chart.lines[0].data, year);
      if (Number.isFinite(value)) plan.topDown[index] = Math.max(0, value * 1000000);
    });
    rememberOtherRevenueControlPoints("topDown", chart.lines[0].data);
    saveScenarios();
    syncOtherRevenueCharts({ excludeChart: chart });
    renderAll({ syncNewCustomerEditors: false });
  };
  state.otherRevenueCharts.otherRevenue = chart;
}

function profitPlan(scenario = activeScenario()) {
  normalizeProfitPlan(scenario);
  return scenario.profitPlan;
}

function derivedProfitRevenueValues(scenario = activeScenario()) {
  return totalRevenueValues(calculateOutputs(scenario), scenario);
}

function profitRevenueValues(scenario = activeScenario()) {
  const plan = profitPlan(scenario);
  const derived = derivedProfitRevenueValues(scenario);
  if (!plan.revenueEdited) return derived;
  return YEARS.map((_year, index) => {
    const value = Number(plan.revenue[index]);
    return Number.isFinite(value) ? Math.max(0, value) : derived[index];
  });
}

function profitCostValues(scenario = activeScenario()) {
  return totalCostValues(scenario);
}

function profitBottomUpValues(scenario = activeScenario()) {
  const revenue = profitRevenueValues(scenario);
  const cost = profitCostValues(scenario);
  return YEARS.map((_year, index) => Number(revenue[index] || 0) - Number(cost[index] || 0));
}

function profitTopDownValues(scenario = activeScenario()) {
  const plan = profitPlan(scenario);
  const bottomUp = profitBottomUpValues(scenario);
  if (!plan.profitEdited) return bottomUp;
  return YEARS.map((_year, index) => {
    const value = Number(plan.profit[index]);
    return Number.isFinite(value) ? value : bottomUp[index];
  });
}

function profitPlanControlPoints(scenario = activeScenario()) {
  return profitPlan(scenario).controlPoints;
}

function controlledProfitPairs(key, values, scale = 1) {
  const stored = profitPlanControlPoints()[key];
  if (Array.isArray(stored) && stored.length >= 2) {
    return stored
      .map(year => {
        const index = YEARS.indexOf(Number(year));
        if (index < 0) return null;
        const value = values[index];
        return value === null || value === undefined ? null : [Number(year), value / scale];
      })
      .filter(Boolean);
  }
  return YEARS
    .map((year, index) => {
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / scale];
    })
    .filter(Boolean);
}

function rememberProfitControlPoints(key, pairs) {
  const years = (pairs || [])
    .map(point => Number(point?.[0]))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
  if (years.length >= 2) {
    profitPlanControlPoints()[key] = Array.from(new Set(years));
  } else {
    delete profitPlanControlPoints()[key];
  }
}

function profitValuesMatch(scenario = activeScenario()) {
  const top = profitTopDownValues(scenario);
  const bottom = profitBottomUpValues(scenario);
  return YEARS.every((year, index) => {
    if (year < FIRST_COST_EDIT_YEAR) return true;
    return Math.abs(Number(top[index] || 0) - Number(bottom[index] || 0)) < 1;
  });
}

function profitChartPairs(key, values) {
  return controlledProfitPairs(key, values, 1000000);
}

function profitChartLines() {
  return [
    {
      name: "Profit",
      color: TOP_DOWN_COLOR,
      editable: true,
      editDomain: {
        moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: profitChartPairs("profit", profitTopDownValues(activeScenario())),
    },
    {
      name: "Bottom Up",
      color: "#98a2b3",
      editable: false,
      data: profitChartPairs("bottomUp", profitBottomUpValues(activeScenario())),
    },
  ];
}

function profitRevenueChartLines() {
  return [{
    name: "Revenue",
    color: BOTTOM_UP_COLOR,
    editable: true,
    editDomain: {
      moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: profitChartPairs("revenue", profitRevenueValues(activeScenario())),
  }];
}

function profitCostChartLines() {
  return [{
    name: "Cost",
    color: "#a96c50",
    editable: true,
    editDomain: {
      moveX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_COST_EDIT_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: profitChartPairs("cost", profitCostValues(activeScenario())),
  }];
}

function ensureProfitSnapshot(key, scenario = activeScenario()) {
  const plan = profitPlan(scenario);
  if (key === "profit" && !plan.profitEdited) {
    plan.profit = clone(profitBottomUpValues(scenario));
    plan.profitEdited = true;
  }
  if (key === "revenue" && !plan.revenueEdited) {
    plan.revenue = clone(derivedProfitRevenueValues(scenario));
    plan.revenueEdited = true;
  }
}

function setProfitRevenueDollarValue(scenario, index, value) {
  const actualValue = Math.max(0, Number(value) || 0);
  const plan = profitPlan(scenario);
  ensureProfitSnapshot("revenue", scenario);
  plan.revenue[index] = actualValue;
}

function setProfitCostDollarValue(scenario, index, value) {
  setScenarioCostValue(scenario, "total", YEARS[index], Math.max(0, Number(value) || 0) / 1000000);
}

function setProfitTopDollarValue(scenario, index, value) {
  const plan = profitPlan(scenario);
  ensureProfitSnapshot("profit", scenario);
  plan.profit[index] = Number(value) || 0;
}

function applyProfitTopToBottomForYear(scenario, index, targetProfit) {
  const plan = profitPlan(scenario);
  const locks = plan.locks || {};
  if (locks.revenue && locks.cost) return false;
  const revenue = profitRevenueValues(scenario)[index];
  const cost = profitCostValues(scenario)[index];
  const currentProfit = Number(revenue || 0) - Number(cost || 0);
  const target = Number(targetProfit) || 0;
  if (locks.revenue) {
    setProfitCostDollarValue(scenario, index, Math.max(0, Number(revenue || 0) - target));
    return true;
  }
  if (locks.cost) {
    setProfitRevenueDollarValue(scenario, index, Math.max(0, target + Number(cost || 0)));
    return true;
  }
  const delta = target - currentProfit;
  setProfitRevenueDollarValue(scenario, index, Math.max(0, Number(revenue || 0) + delta / 2));
  setProfitCostDollarValue(scenario, index, Math.max(0, Number(cost || 0) - delta / 2));
  return true;
}

function setProfitTopFromBottomForYear(scenario, index) {
  const plan = profitPlan(scenario);
  if (plan.locks?.profit) return false;
  const bottomUp = profitBottomUpValues(scenario)[index];
  setProfitTopDollarValue(scenario, index, bottomUp);
  return true;
}

function applyProfitRevenueEditForYear(scenario, index, targetRevenue) {
  const plan = profitPlan(scenario);
  const locks = plan.locks || {};
  if (locks.profit && locks.cost) return false;
  setProfitRevenueDollarValue(scenario, index, targetRevenue);
  if (locks.profit) {
    const targetProfit = profitTopDownValues(scenario)[index];
    setProfitCostDollarValue(scenario, index, Math.max(0, Number(targetRevenue || 0) - Number(targetProfit || 0)));
  } else {
    setProfitTopFromBottomForYear(scenario, index);
  }
  return true;
}

function applyProfitCostEditForYear(scenario, index, targetCost) {
  const plan = profitPlan(scenario);
  const locks = plan.locks || {};
  if (locks.profit && locks.revenue) return false;
  setProfitCostDollarValue(scenario, index, targetCost);
  if (locks.profit) {
    const targetProfit = profitTopDownValues(scenario)[index];
    setProfitRevenueDollarValue(scenario, index, Number(targetProfit || 0) + Number(targetCost || 0));
  } else {
    setProfitTopFromBottomForYear(scenario, index);
  }
  return true;
}

function renderProfitDrilldown() {
  const matched = profitValuesMatch();
  const plan = profitPlan();
  const topToBottomButton = document.getElementById("set-profit-top-to-bottom");
  const bottomToTopButton = document.getElementById("set-profit-bottom-to-top");
  if (topToBottomButton) topToBottomButton.disabled = matched || (plan.locks.revenue && plan.locks.cost);
  if (bottomToTopButton) bottomToTopButton.disabled = matched || plan.locks.profit;
  document.querySelectorAll("[data-profit-lock-key]").forEach(button => {
    const key = button.dataset.profitLockKey;
    const locked = Boolean(plan.locks?.[key]);
    button.classList.toggle("locked", locked);
    button.textContent = locked ? "Locked" : "Lock";
    button.setAttribute("aria-pressed", locked ? "true" : "false");
  });
}

function syncProfitCharts({ excludeChart = null } = {}) {
  state.syncingProfitCharts = true;
  const profitChart = state.profitCharts.profit;
  if (profitChart && profitChart !== excludeChart) {
    profitChart.lines = safeNapkinLines(profitChartLines());
    profitChart._refreshChart();
    styleGrowthReferenceSeries(profitChart);
  }
  const revenueChart = state.profitCharts.revenue;
  if (revenueChart && revenueChart !== excludeChart) {
    revenueChart.lines = safeNapkinLines(profitRevenueChartLines());
    revenueChart._refreshChart();
  }
  const costChart = state.profitCharts.cost;
  if (costChart && costChart !== excludeChart) {
    costChart.lines = safeNapkinLines(profitCostChartLines());
    costChart._refreshChart();
  }
  state.syncingProfitCharts = false;
  renderProfitDrilldown();
}

function setProfitBottomToTop() {
  const plan = profitPlan();
  if (plan.locks.profit || profitValuesMatch()) return;
  pushUndoSnapshot();
  YEARS.forEach((year, index) => {
    if (year < FIRST_COST_EDIT_YEAR) return;
    setProfitTopFromBottomForYear(activeScenario(), index);
  });
  plan.controlPoints.profit = YEARS.slice();
  saveScenarios();
  syncProfitCharts();
  renderAll();
}

function setProfitTopToBottom() {
  const plan = profitPlan();
  if ((plan.locks.revenue && plan.locks.cost) || profitValuesMatch()) return;
  pushUndoSnapshot();
  const top = profitTopDownValues(activeScenario());
  YEARS.forEach((year, index) => {
    if (year < FIRST_COST_EDIT_YEAR) return;
    applyProfitTopToBottomForYear(activeScenario(), index, top[index]);
  });
  saveScenarios();
  syncProfitCharts();
  syncCostCharts();
  renderAll();
}

function toggleProfitLock(key) {
  if (!["profit", "revenue", "cost"].includes(key)) return;
  pushUndoSnapshot();
  const plan = profitPlan();
  const nextLocked = !plan.locks[key];
  plan.locks[key] = nextLocked;
  if (nextLocked && key === "revenue") {
    ensureProfitSnapshot("revenue");
    plan.locks.cost = false;
  } else if (nextLocked && key === "cost") {
    plan.locks.revenue = false;
  } else if (nextLocked && key === "profit") {
    ensureProfitSnapshot("profit");
  }
  saveScenarios();
  renderProfitDrilldown();
}

function profitTooltipFormatter(params, formatter) {
  const formatValue = typeof formatter === "function"
    ? formatter
    : value => formatCurrency(Number(value) * 1000000, 0);
  const items = Array.isArray(params) ? params : [params];
  const rows = items
    .filter(item => item && item.seriesType === "line")
    .map(item => {
      const data = Array.isArray(item.data) ? item.data : item.value;
      const rawValue = Array.isArray(data) ? data[1] : item.value;
      return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatValue(Number(rawValue))}`;
    });
  return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
}

function makeProfitNapkinChart(id, lines, yAxis, onValuesChanged) {
  const chart = new NapkinChart(
    id,
    safeNapkinLines(lines),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis,
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: { trigger: "axis", formatter: profitTooltipFormatter },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingProfitCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingProfitCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    onValuesChanged(chart);
    saveScenarios();
    syncProfitCharts({ excludeChart: chart });
    syncCostCharts();
    renderAll({ syncNewCustomerEditors: false });
  };
  return chart;
}

function initProfitCharts() {
  const yAxisCurrency = (min, max) => ({
    type: "value",
    min,
    max,
    axisLabel: { formatter: value => formatCompactCurrency(Number(value) * 1000000) },
  });
  const profitChart = makeProfitNapkinChart(
    "profit-top-chart",
    profitChartLines(),
    yAxisCurrency(-100, 200),
    chart => {
      const scenario = activeScenario();
      YEARS.forEach((year, index) => {
        if (year < FIRST_COST_EDIT_YEAR) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        if (!Number.isFinite(value)) return;
        setProfitTopDollarValue(scenario, index, value * 1000000);
        applyProfitTopToBottomForYear(scenario, index, value * 1000000);
      });
      rememberProfitControlPoints("profit", chart.lines[0].data);
    }
  );
  styleGrowthReferenceSeries(profitChart);
  state.profitCharts.profit = profitChart;

  state.profitCharts.revenue = makeProfitNapkinChart(
    "profit-revenue-chart",
    profitRevenueChartLines(),
    yAxisCurrency(0, 220),
    chart => {
      const scenario = activeScenario();
      YEARS.forEach((year, index) => {
        if (year < FIRST_COST_EDIT_YEAR) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        if (Number.isFinite(value)) applyProfitRevenueEditForYear(scenario, index, value * 1000000);
      });
      rememberProfitControlPoints("revenue", chart.lines[0].data);
    }
  );

  state.profitCharts.cost = makeProfitNapkinChart(
    "profit-cost-chart",
    profitCostChartLines(),
    yAxisCurrency(0, 120),
    chart => {
      const scenario = activeScenario();
      YEARS.forEach((year, index) => {
        if (year < FIRST_COST_EDIT_YEAR) return;
        const value = interpolateNapkinLineValue(chart.lines[0].data, year);
        if (Number.isFinite(value)) applyProfitCostEditForYear(scenario, index, value * 1000000);
      });
      rememberProfitControlPoints("cost", chart.lines[0].data);
    }
  );
}

function renderSaasMetrics(outputs) {
  const scenario = activeScenario();
  const profit = profitValues(outputs, scenario);
  const totalCosts = totalCostValues(scenario);
  const newCustomerValues = effectiveNewCustomers(scenario);
  const laborFteTotals = laborFteTotalValues(scenario);
  const revenueGrowth = revenueGrowthPercentValues(outputs.revenue);
  const profitMargin = profitMarginPercentValues(outputs.revenue, profit);
  const ruleOf40 = ruleOf40Values(outputs.revenue, profit);
  const fullyLoadedCac = fullyLoadedCacValues(totalCosts, newCustomerValues);
  const revenuePerFte = revenuePerFteValues(outputs.revenue, laborFteTotals);
  const last = YEARS.length - 1;

  setText("saas-rule-of-40-2029", formatSaasPercent(ruleOf40[last], 1));
  setText("saas-revenue-growth-2029", formatSaasPercent(revenueGrowth[last], 1));
  setText("saas-profit-margin-2029", formatSaasPercent(profitMargin[last], 1));
  setText("saas-cac-2029", formatOptionalCurrency(fullyLoadedCac[last], 0));
  setText("saas-cac-new-customers-2029", formatIntegerMetric(newCustomerValues[last]));
  setText("saas-cac-total-cost-2029", formatCurrency(totalCosts[last], 0));
  setText("saas-revenue-per-fte-2029", formatOptionalCurrency(revenuePerFte[last], 0));
  setText("saas-revenue-per-fte-revenue-2029", formatCurrency(outputs.revenue[last], 0));
  setText("saas-revenue-per-fte-fte-2029", formatDecimalMetric(laborFteTotals[last], 1));

  state.outputCharts.saasRuleOf40?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const lines = [tooltipHeader(params[0]?.axisValue)];
        params.forEach(item => {
          const value = item.value === null || item.value === undefined ? null : Number(item.value);
          lines.push(`${item.marker} ${displayLabel(item.seriesName)}: ${formatSaasPercent(value, 1)}`);
        });
        return lines.join("<br/>");
      },
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: value => formatSaasPercent(Number(value), 0) } },
    series: [
      {
        name: "Rule of 40",
        type: "line",
        data: ruleOf40,
        symbolSize: 6,
        lineStyle: { color: "#1d4ed8", width: 3 },
        itemStyle: { color: "#1d4ed8" },
      },
      {
        name: "Revenue Growth",
        type: "line",
        data: revenueGrowth,
        symbolSize: 5,
        lineStyle: { color: "#4f7f52", width: 2 },
        itemStyle: { color: "#4f7f52" },
      },
      {
        name: "Profit Margin",
        type: "line",
        data: profitMargin,
        symbolSize: 5,
        lineStyle: { color: "#8b5cf6", width: 2 },
        itemStyle: { color: "#8b5cf6" },
      },
      {
        name: "40 Target",
        type: "line",
        data: YEARS.map(() => 40),
        symbol: "none",
        silent: true,
        lineStyle: { color: "#111827", width: 1.5, type: "dashed" },
        itemStyle: { color: "#111827" },
      },
    ],
  }, true);

  state.outputCharts.saasCac?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const lines = [tooltipHeader(params[0]?.axisValue)];
        params.forEach(item => {
          const value = item.value === null || item.value === undefined ? null : Number(item.value);
          lines.push(`${item.marker} ${displayLabel(item.seriesName)}: ${formatOptionalCurrency(value, 0)}`);
        });
        return lines.join("<br/>");
      },
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: value => formatCompactCurrency(Number(value)) } },
    series: [
      {
        name: "Fully Loaded CAC",
        type: "line",
        data: fullyLoadedCac,
        symbolSize: 6,
        lineStyle: { color: "#0f766e", width: 3 },
        itemStyle: { color: "#0f766e" },
      },
    ],
  }, true);

  state.outputCharts.saasRevenuePerFte?.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      formatter: params => {
        const lines = [tooltipHeader(params[0]?.axisValue)];
        params.forEach(item => {
          const value = item.value === null || item.value === undefined ? null : Number(item.value);
          lines.push(`${item.marker} ${displayLabel(item.seriesName)}: ${formatOptionalCurrency(value, 0)}`);
        });
        return lines.join("<br/>");
      },
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 36, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: value => formatCompactCurrency(Number(value)) } },
    series: [
      {
        name: "Revenue / FTE",
        type: "line",
        data: revenuePerFte,
        symbolSize: 6,
        lineStyle: { color: "#7c3aed", width: 3 },
        itemStyle: { color: "#7c3aed" },
      },
    ],
  }, true);
}

const revenueDrilldownCategoryMeta = {
  ltr: { label: "LTR", chartKey: "ltr", color: "#4f7fb8", yMax: 170 },
  growth: { label: "Growth", chartKey: "growth", color: "#6fa76b", yMax: 25 },
  str: { label: "STR", chartKey: "str", color: "#7b6fb8", yMax: 15 },
  other: { label: "Other", chartKey: "other", color: "#a96c50", yMax: 5 },
};

function revenueDrilldownControlPoints(key) {
  normalizeRevenuePaths(activeScenario());
  if (key === "total") return activeScenario().revenuePaths.total.controlPoints;
  if (key === "ltr") return activeScenario().revenuePaths.ltr.controlPoints;
  if (key === "growth") return growthRevenueControlPoints();
  if (key === "str") return strRevenueControlPoints();
  if (key === "other") return otherRevenueControlPoints();
  return {};
}

function controlledRevenueDrilldownPairs(key, controlKey, values, scale = 1000000) {
  const historicalPairs = YEARS
    .map((year, index) => {
      if (year >= FIRST_FORECAST_YEAR) return null;
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / scale];
    })
    .filter(Boolean);
  const stored = revenueDrilldownControlPoints(key)[controlKey];
  if (Array.isArray(stored) && stored.length >= 2) {
    const storedPairs = stored
      .map(year => {
        const index = YEARS.indexOf(Number(year));
        if (index < 0) return null;
        const value = values[index];
        return value === null || value === undefined ? null : [Number(year), value / scale];
      })
      .filter(Boolean);
    return [...historicalPairs, ...storedPairs]
      .reduce((pairs, pair) => {
        const existingIndex = pairs.findIndex(existingPair => existingPair[0] === pair[0]);
        if (existingIndex >= 0) {
          pairs[existingIndex] = pair;
        } else {
          pairs.push(pair);
        }
        return pairs;
      }, [])
      .sort((left, right) => left[0] - right[0]);
  }
  return YEARS
    .map((year, index) => {
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / scale];
    })
    .filter(Boolean);
}

function rememberRevenueDrilldownControlPoints(key, controlKey, pairs) {
  const years = (pairs || [])
    .map(point => Number(point?.[0]))
    .filter(Number.isFinite)
    .sort((left, right) => left - right);
  const target = revenueDrilldownControlPoints(key);
  if (years.length >= 2) {
    target[controlKey] = Array.from(new Set(years));
  } else {
    delete target[controlKey];
  }
}

function revenueTotalTopDownPairs() {
  return controlledRevenueDrilldownPairs("total", "topDown", revenueTotalTopDownValues(activeScenario()));
}

function revenueBottomUpPairs() {
  return controlledRevenueDrilldownPairs("total", "bottomUp", revenueBottomUpValues(activeScenario()));
}

function totalRevenueTopMatchesBottomUp(scenario = activeScenario(), outputs = calculateOutputs(scenario)) {
  const topDown = revenueTotalTopDownValues(scenario);
  const bottomUp = revenueBottomUpValues(scenario, outputs);
  return YEARS.every((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return true;
    return Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0)) < 1;
  });
}

function renderTotalRevenueDrilldownChart(outputs = calculateOutputs(activeScenario())) {
  const chart = state.outputCharts.revenueDrilldownTotal;
  if (!chart) return;
  const scenario = activeScenario();
  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
  const scenarioName = displayScenarioName(scenario);
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const topDown = revenueTotalTopDownValues(scenario);
  const bottomUp = revenueBottomUpValues(scenario, outputs);
  const series = [
    { name: `${scenarioName} Top-Down`, type: "line", data: topDown, symbolSize: 6, lineStyle: { color: TOP_DOWN_COLOR, width: 3 }, itemStyle: { color: TOP_DOWN_COLOR } },
    { name: `${scenarioName} Bottom-Up`, type: "line", data: bottomUp, symbolSize: 6, lineStyle: { color: BOTTOM_UP_COLOR, width: 3 }, itemStyle: { color: BOTTOM_UP_COLOR } },
  ];
  if (comparison && comparisonOutputs) {
    series.push({
      name: `${comparisonName} Top-Down`,
      type: "line",
      data: revenueTotalTopDownValues(comparison),
      symbolSize: 5,
      lineStyle: { color: "#98a2b3", width: 2 },
      itemStyle: { color: "#98a2b3" },
    });
    series.push({
      name: `${comparisonName} Bottom-Up`,
      type: "line",
      data: revenueBottomUpValues(comparison, comparisonOutputs),
      symbolSize: 5,
      lineStyle: { color: "#98a2b3", width: 2, type: "dashed" },
      itemStyle: { color: "#98a2b3" },
    });
  }
  chart.setOption({
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
    series,
  }, true);

  const setButton = document.getElementById("set-total-revenue-bottom-to-top");
  if (setButton) setButton.disabled = totalRevenueTopMatchesBottomUp(scenario, outputs);
}

function setTotalRevenueBottomToTop() {
  const scenario = activeScenario();
  const outputs = calculateOutputs(scenario);
  if (totalRevenueTopMatchesBottomUp(scenario, outputs)) return;
  pushUndoSnapshot();
  const bottomUp = revenueBottomUpValues(scenario, outputs);
  YEARS.forEach((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return;
    setRevenueTotalTopDownValue(scenario, index, bottomUp[index]);
  });
  scenario.revenuePaths.total.controlPoints.topDown = YEARS.filter(year => year >= FIRST_FORECAST_YEAR);
  saveScenarios();
  syncProfitCharts();
  renderAll();
}

function ltrRevenueBottomUpPairs() {
  const values = ltrRevenueBottomUpValues(activeScenario());
  return YEARS
    .map((year, index) => {
      const value = values[index];
      return value === null || value === undefined ? null : [year, value / 1000000];
    })
    .filter(Boolean);
}

function revenueCategoryPairs(key) {
  return controlledRevenueDrilldownPairs(key, "topDown", revenueCategoryValues(activeScenario(), key));
}

function revenueDrilldownTopLines() {
  return [
    {
      name: "Revenue",
      color: TOP_DOWN_COLOR,
      editable: true,
      editDomain: {
        moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
        deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      },
      data: revenueTotalTopDownPairs(),
    },
    {
      name: "Bottom Up",
      color: "#98a2b3",
      editable: false,
      data: revenueBottomUpPairs(),
    },
  ];
}

function revenueCategoryChartLines(key) {
  const meta = revenueDrilldownCategoryMeta[key];
  const lines = [{
    name: meta.label,
    color: meta.color,
    editable: true,
    editDomain: {
      moveX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      addX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
      deleteX: [[FIRST_FORECAST_YEAR, YEARS[YEARS.length - 1]]],
    },
    data: revenueCategoryPairs(key),
  }];
  if (key === "ltr") {
    lines.push({
      name: "Bottom Up",
      color: "#98a2b3",
      editable: false,
      data: ltrRevenueBottomUpPairs(),
    });
  }
  return lines;
}

function setRevenueTotalTopDownValue(scenario, index, value) {
  normalizeRevenuePaths(scenario);
  scenario.revenuePaths.total.topDown[index] = Math.max(0, Number(value) || 0);
}

function setRevenueCategoryValue(scenario, key, index, value) {
  normalizeRevenuePaths(scenario);
  const numericValue = Math.max(0, Number(value) || 0);
  if (key === "ltr") {
    scenario.revenuePaths.ltr.topDown[index] = numericValue;
    scenario.revenuePaths.ltr.edited = true;
  } else if (key === "growth") {
    growthRevenuePlan(scenario).topDown[index] = numericValue;
  } else if (key === "str") {
    strRevenuePlan(scenario).topDown[index] = numericValue;
  } else if (key === "other") {
    otherRevenuePlan(scenario).topDown[index] = numericValue;
  }
}

function renderRevenueDrilldown() {
  renderTotalRevenueDrilldownChart();
  const ltrChart = state.revenueDrilldownCharts.ltr;
  if (ltrChart) styleRevenueBottomUpReferenceSeries(ltrChart);
}

function styleRevenueBottomUpReferenceSeries(chart) {
  const option = chart.chart.getOption();
  const series = (option.series || []).map(seriesItem => {
    if (String(seriesItem.name || "") !== "Bottom Up") return { ...seriesItem, z: 4 };
    return {
      ...seriesItem,
      silent: true,
      showSymbol: false,
      symbolSize: 0,
      z: 2,
      itemStyle: { ...(seriesItem.itemStyle || {}), color: "#98a2b3" },
      lineStyle: { ...(seriesItem.lineStyle || {}), color: "#98a2b3", width: 3, opacity: 1 },
    };
  });
  chart.chart.setOption({ series }, false);
}

function syncRevenueDrilldownCharts({ excludeChart = null } = {}) {
  state.syncingRevenueDrilldownCharts = true;
  renderTotalRevenueDrilldownChart();
  Object.keys(revenueDrilldownCategoryMeta).forEach(key => {
    const chart = state.revenueDrilldownCharts[key];
    if (chart && chart !== excludeChart) {
      chart.lines = safeNapkinLines(revenueCategoryChartLines(key));
      chart._refreshChart();
      if (key === "ltr") styleRevenueBottomUpReferenceSeries(chart);
    }
  });
  state.syncingRevenueDrilldownCharts = false;
}

function revenueDrilldownTooltipFormatter(params) {
  const items = Array.isArray(params) ? params : [params];
  const rows = items
    .filter(item => item && item.seriesType === "line")
    .map(item => {
      const data = Array.isArray(item.data) ? item.data : item.value;
      const rawValue = Array.isArray(data) ? data[1] : item.value;
      return `${item.marker || ""} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(rawValue) * 1000000, 0)}`;
    });
  return [tooltipHeader(items[0]?.axisValue), ...rows].join("<br/>");
}

function makeRevenueDrilldownNapkinChart(id, lines, yMax, onValuesChanged) {
  const chart = new NapkinChart(
    id,
    safeNapkinLines(lines),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1, axisLabel: { formatter: formatAxisYear } },
      yAxis: { type: "value", min: 0, max: yMax, axisLabel: { formatter: value => formatCompactCurrency(Number(value) * 1000000) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: { trigger: "axis", formatter: revenueDrilldownTooltipFormatter },
    },
    "none",
    false
  );
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart._appEditSnapshot = null;
  chart._appEditCommitted = false;
  chart.chart.getZr().on("mousedown", () => {
    if (state.syncingRevenueDrilldownCharts) return;
    chart._appEditSnapshot = snapshotState();
    chart._appEditCommitted = false;
  });
  chart.onDataChanged = () => {
    if (state.syncingRevenueDrilldownCharts) return;
    if (chart._appEditSnapshot && !chart._appEditCommitted) {
      pushUndoSnapshot(chart._appEditSnapshot);
      chart._appEditCommitted = true;
    }
    onValuesChanged(chart);
    saveScenarios();
    syncRevenueDrilldownCharts({ excludeChart: chart });
    syncGrowthRevenueCharts();
    syncStrRevenueCharts();
    syncOtherRevenueCharts();
    syncProfitCharts();
    renderAll({ syncNewCustomerEditors: false });
  };
  return chart;
}

function initRevenueDrilldownCharts() {
  Object.keys(revenueDrilldownCategoryMeta).forEach(key => {
    const meta = revenueDrilldownCategoryMeta[key];
    state.revenueDrilldownCharts[key] = makeRevenueDrilldownNapkinChart(
      `revenue-drilldown-${key}-chart`,
      revenueCategoryChartLines(key),
      meta.yMax,
      chart => {
        const scenario = activeScenario();
        YEARS.forEach((year, index) => {
          if (year < FIRST_FORECAST_YEAR) return;
          const value = interpolateNapkinLineValue(chart.lines[0].data, year);
          if (Number.isFinite(value)) setRevenueCategoryValue(scenario, key, index, value * 1000000);
        });
        rememberRevenueDrilldownControlPoints(key, "topDown", chart.lines[0].data);
        rememberRevenueDrilldownControlPoints("total", "bottomUp", revenueBottomUpPairs());
      }
    );
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
  const costs = totalCostValues(scenario);
  const comparisonCosts = comparison ? totalCostValues(comparison) : null;
  const totalRevenue = totalRevenueValues(outputs, scenario);
  const comparisonTotalRevenue = comparison && comparisonOutputs ? totalRevenueValues(comparisonOutputs, comparison) : null;
  const profits = profitValues(outputs, scenario);
  const comparisonProfits = comparison && comparisonOutputs ? profitValues(comparisonOutputs, comparison) : null;
  const laborFteTotals = laborFteTotalValues(scenario);
  const comparisonLaborFteTotals = comparison ? laborFteTotalValues(comparison) : null;
  const laborCostPerFte = laborCostPerFteAggregateValues(scenario);
  const comparisonLaborCostPerFte = comparison ? laborCostPerFteAggregateValues(comparison) : null;
  const strRevenueValues = strRevenueTopDownValues(scenario);
  const comparisonStrRevenueValues = comparison ? strRevenueTopDownValues(comparison) : null;
  const strPropertiesValues = strRevenuePropertiesValues(scenario);
  const comparisonStrPropertiesValues = comparison ? strRevenuePropertiesValues(comparison) : null;
  const strRevPerPropertyValues = strRevenueRevPerPropertyValues(scenario);
  const comparisonStrRevPerPropertyValues = comparison ? strRevenueRevPerPropertyValues(comparison) : null;
  const otherRevenueValues = otherRevenueTopDownValues(scenario);
  const comparisonOtherRevenueValues = comparison ? otherRevenueTopDownValues(comparison) : null;
  const growthRevenueValues = growthRevenueTopDownValues(scenario);
  const comparisonGrowthRevenueValues = comparison ? growthRevenueTopDownValues(comparison) : null;
  const growthProxyCustomers = growthRevenueProxyCustomers(scenario);
  const comparisonGrowthProxyCustomers = comparison ? driverValues(comparison, "newCustomers") : null;
  const growthRevenuePerProxyCustomer = growthRevenueRevPerNewCustomerValues(scenario);
  const comparisonGrowthRevenuePerProxyCustomer = comparisonGrowthProxyCustomers
    ? growthRevenueRevPerNewCustomerValues(comparison)
    : null;
  const scenarioName = displayScenarioName(scenario);
  const comparisonName = comparison ? displayScenarioName(comparison) : "";

  setMetricTreeChart("treeProfit", [
    metricTreeSeries(scenarioName, profits, scenarioColor),
    ...(comparisonProfits ? [metricTreeSeries(comparisonName, comparisonProfits, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeRevenue", [
    metricTreeSeries(scenarioName, totalRevenue, scenarioColor),
    metricTreeSeries("Plan", outputs.planRevenue, "#111827"),
    ...(comparisonTotalRevenue ? [metricTreeSeries(comparisonName, comparisonTotalRevenue, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeLtrRevenue", [
    metricTreeSeries(displayLabel(`${scenarioName} LTR`), revenueCategoryValues(scenario, "ltr", outputs), scenarioColor),
    ...(comparisonOutputs ? [metricTreeSeries(displayLabel(`${comparisonName} LTR`), revenueCategoryValues(comparison, "ltr", comparisonOutputs), comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeStrRevenue", [
    metricTreeSeries(scenarioName, strRevenueValues, scenarioColor),
    ...(comparisonStrRevenueValues ? [metricTreeSeries(comparisonName, comparisonStrRevenueValues, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeStrProperties", [
    metricTreeSeries(scenarioName, strPropertiesValues, scenarioColor),
    ...(comparisonStrPropertiesValues ? [metricTreeSeries(comparisonName, comparisonStrPropertiesValues, comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeStrRevPerProperty", [
    metricTreeSeries(scenarioName, strRevPerPropertyValues, scenarioColor),
    ...(comparisonStrRevPerPropertyValues ? [metricTreeSeries(comparisonName, comparisonStrRevPerPropertyValues, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeGrowthRevenue", [
    metricTreeSeries(scenarioName, growthRevenueValues, scenarioColor),
    ...(comparisonGrowthRevenueValues ? [metricTreeSeries(comparisonName, comparisonGrowthRevenueValues, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeGrowthNewPayingCustomers", [
    metricTreeSeries(scenarioName, growthProxyCustomers, scenarioColor),
    ...(comparisonGrowthProxyCustomers ? [metricTreeSeries(comparisonName, comparisonGrowthProxyCustomers, comparisonColor)] : []),
  ]);

  setMetricTreeChart("treeGrowthRevPerNewPayingCustomer", [
    metricTreeSeries(scenarioName, growthRevenuePerProxyCustomer, scenarioColor),
    ...(comparisonGrowthRevenuePerProxyCustomer ? [metricTreeSeries(comparisonName, comparisonGrowthRevenuePerProxyCustomer, comparisonColor)] : []),
  ], { format: "currency2" });

  setMetricTreeChart("treeOtherRevenue", [
    metricTreeSeries(scenarioName, otherRevenueValues, scenarioColor),
    ...(comparisonOtherRevenueValues ? [metricTreeSeries(comparisonName, comparisonOtherRevenueValues, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeCost", [
    metricTreeSeries(scenarioName, costs, scenarioColor),
    ...(comparisonCosts ? [metricTreeSeries(comparisonName, comparisonCosts, comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeLaborCost", [
    metricTreeSeries(scenarioName, costValues(scenario, "labor"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, costValues(comparison, "labor"), comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeNonLaborCost", [
    metricTreeSeries(scenarioName, costValues(scenario, "nonLabor"), scenarioColor),
    ...(comparison ? [metricTreeSeries(comparisonName, costValues(comparison, "nonLabor"), comparisonColor)] : []),
  ], { format: "currency" });

  setMetricTreeChart("treeLaborFte", [
    metricTreeSeries(scenarioName, laborFteTotals, scenarioColor),
    ...(comparisonLaborFteTotals ? [metricTreeSeries(comparisonName, comparisonLaborFteTotals, comparisonColor)] : []),
  ], { format: "decimal" });

  setMetricTreeChart("treeLaborCostPerFte", [
    metricTreeSeries(scenarioName, laborCostPerFte, scenarioColor),
    ...(comparisonLaborCostPerFte ? [metricTreeSeries(comparisonName, comparisonLaborCostPerFte, comparisonColor)] : []),
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
  const scenario = activeScenario();
  const comparison = compareScenario();
  const comparisonOutputs = comparison ? calculateOutputs(comparison) : null;
  const scenarioName = displayScenarioName(activeScenario());
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const ltrTopDown = ltrRevenueTopDownValues(scenario, outputs);
  const ltrBottomUp = ltrRevenueBottomUpValues(scenario, outputs);
  const revenueSeries = [
    { name: `${scenarioName} LTR Top Down`, type: "line", data: ltrTopDown, symbolSize: 6, lineStyle: { color: TOP_DOWN_COLOR, width: 3 }, itemStyle: { color: TOP_DOWN_COLOR } },
    { name: `${scenarioName} Bottom Up`, type: "line", data: ltrBottomUp, symbolSize: 6, lineStyle: { color: BOTTOM_UP_COLOR, width: 3 }, itemStyle: { color: BOTTOM_UP_COLOR } },
  ];
  if (comparisonOutputs) {
    const comparisonTopDown = ltrRevenueTopDownValues(comparison, comparisonOutputs);
    const comparisonBottomUp = ltrRevenueBottomUpValues(comparison, comparisonOutputs);
    revenueSeries.push({
      name: `${comparisonName} LTR Top Down`,
      type: "line",
      data: comparisonTopDown,
      symbolSize: 5,
      lineStyle: { color: "#98a2b3", width: 2 },
      itemStyle: { color: "#98a2b3" },
    });
    revenueSeries.push({
      name: `${comparisonName} Bottom Up`,
      type: "line",
      data: comparisonBottomUp,
      symbolSize: 5,
      lineStyle: { color: "#98a2b3", width: 2, type: "dashed" },
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
          lines.push(`${item.marker} ${displayLabel(item.seriesName)}: ${formatCurrency(Number(item.value), 0)}`);
        });
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
  const scenario = activeScenario();
  const cumulativeProfit = scenarioCumulativeProfitForYears(scenario);
  const comparison = compareScenario();
  const comparisonProfit = comparison ? scenarioCumulativeProfitForYears(comparison) : null;
  const latestVersion = latestScenarioVersion(scenario);
  const latestVersionScenario = latestVersion ? scenarioFromVersion(scenario, latestVersion) : null;
  const topLineMetric = topLineMetricForCurrentView(scenario);
  const versionTopLineMetric = latestVersionScenario ? topLineMetricForCurrentView(latestVersionScenario) : null;
  const compareDelta = comparison ? cumulativeProfit - comparisonProfit : null;
  const currentMetricDelta = topLineMetricPeriodDelta(topLineMetric.values);
  const versionMetricDelta = versionTopLineMetric ? topLineMetricPeriodDelta(versionTopLineMetric.values) : null;
  const versionDelta = versionTopLineMetric ? currentMetricDelta - versionMetricDelta : null;
  const compareElement = document.getElementById("kpi-compare-profit-delta");
  const versionElement = document.getElementById("kpi-version-profit-delta");

  setText("kpi-cumulative-profit", formatCurrency(cumulativeProfit, 0));
  setText("kpi-compare-profit-delta", comparison ? formatCurrency(compareDelta, 0) : "-");
  setText("kpi-version-delta-label", `2021-2029 ${topLineMetric.label} Delta to Previous Version`);
  setText("kpi-version-profit-delta", versionTopLineMetric ? formatTopLineMetricDelta(versionDelta, topLineMetric.format) : "-");
  compareElement?.classList.toggle("negative", Number(compareDelta) < 0);
  versionElement?.classList.toggle("negative", Number(versionDelta) < 0);
  renderKpiReconciliationPills(scenario);
}

function refreshHeaderKpis() {
  renderKpis(calculateOutputs(activeScenario()));
}

function renderKpiReconciliationPills(scenario) {
  const container = document.getElementById("kpi-reconciliation-pills");
  if (!container) return;
  const unmatched = scenarioReconciliationChecks(scenario).filter(check => !check.matched);
  if (!unmatched.length) {
    container.innerHTML = `<span class="summary-status-pill matched">Matched</span>`;
    return;
  }
  container.innerHTML = unmatched.map(check => `
    <button
      class="summary-status-pill open"
      type="button"
      data-reconciliation-action="${escapeHtml(check.action || "")}"
      data-focus-chart="${escapeHtml(check.focusChart || "")}"
      data-category-key="${escapeHtml(check.categoryKey || "")}"
    >${escapeHtml(check.label)}</button>
  `).join("");
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
  refreshHeaderKpis();
}

function resizeNewCustomerUnitDrilldownCharts() {
  state.outputCharts.newCustomerUnitBridge?.resize();
  state.cohortCharts.unitNewUnits?.resize();
  state.cohortCharts.unitCustomersPerUnit?.resize();
}

function resizeNewCustomerNewUnitsDrilldownCharts() {
  state.outputCharts.newCustomerNewUnitsBridge?.resize();
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
  const comparisonBottomUp = comparison ? YEARS.map(year => newCustomerUnitBottomUpValue(comparison, year)) : null;

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

  renderNewCustomerNewUnitsDrilldown();
}

function renderNewCustomerNewUnitsDrilldown() {
  const scenario = activeScenario();
  const comparison = compareScenario();
  const comparisonName = comparison ? displayScenarioName(comparison) : "";
  const topDown = YEARS.map(year => newCustomerUnitNewUnitsValue(scenario, year));
  const bottomUp = YEARS.map(year => visibleNewCustomerNewUnitsBottomUpValue(year));
  const comparisonBottomUp = comparison ? YEARS.map(year => newCustomerNewUnitsBottomUpValue(comparison, year)) : null;

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

function formatScenarioSavedAt(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString([], { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" });
}

function scenarioFromVersion(scenario, version) {
  const restored = clone(scenario);
  restored.drivers = clone(version.drivers || scenario.drivers);
  restored.costs = version.costs ? clone(version.costs) : clone(scenario.costs || baseCosts);
  restored.costControlPoints = version.costControlPoints ? clone(version.costControlPoints) : {};
  restored.costPctTotalLines = version.costPctTotalLines ? clone(version.costPctTotalLines) : {};
  restored.costLocks = version.costLocks ? clone(version.costLocks) : {};
  restored.newCustomerSource = version.newCustomerSource || scenario.newCustomerSource || "topDown";
  restored.newCustomerDrilldown = version.newCustomerDrilldown
    ? clone(version.newCustomerDrilldown)
    : clone(scenario.newCustomerDrilldown || createDefaultNewCustomerDrilldown());
  restored.revUnitPlan = version.revUnitPlan ? clone(version.revUnitPlan) : clone(scenario.revUnitPlan || createDefaultRevUnitPlan());
  restored.revenuePaths = version.revenuePaths
    ? clone(version.revenuePaths)
    : clone(scenario.revenuePaths || { total: createDefaultTotalRevenuePlan(), ltr: createDefaultLtrRevenuePlan(), growth: createDefaultGrowthRevenuePlan(), str: createDefaultStrRevenuePlan(), other: createDefaultOtherRevenuePlan() });
  restored.profitPlan = version.profitPlan ? clone(version.profitPlan) : clone(scenario.profitPlan || createDefaultProfitPlan());
  restored.defenses = version.defenses ? clone(version.defenses) : clone(scenario.defenses || createDefaultDefenses());
  normalizeScenario(restored);
  return restored;
}

function scenarioCumulativeProfit(scenario) {
  const outputs = calculateOutputs(scenario);
  return profitValues(outputs, scenario)
    .reduce((sum, value, index) => YEARS[index] >= FIRST_COST_EDIT_YEAR ? sum + Number(value || 0) : sum, 0);
}

function scenarioCumulativeProfitForYears(scenario, startYear = YEARS[0], endYear = YEARS[YEARS.length - 1]) {
  const outputs = calculateOutputs(scenario);
  return profitValues(outputs, scenario)
    .reduce((sum, value, index) => {
      const year = YEARS[index];
      return year >= startYear && year <= endYear ? sum + Number(value || 0) : sum;
    }, 0);
}

function latestScenarioVersion(scenario) {
  return (scenario.versions || [])
    .slice()
    .sort((left, right) => new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime())[0] || null;
}

function topLineMetricPeriodDelta(values, startYear = YEARS[0], endYear = YEARS[YEARS.length - 1]) {
  const startIndex = YEARS.indexOf(startYear);
  const endIndex = YEARS.indexOf(endYear);
  if (startIndex < 0 || endIndex < 0) return 0;
  return Number(values[endIndex] || 0) - Number(values[startIndex] || 0);
}

function formatTopLineMetricDelta(value, format) {
  if (format === "currency") return formatCurrency(value, 0);
  if (format === "integer") return formatIntegerMetric(value);
  if (format === "decimal") return formatDecimalMetric(value, 2);
  if (format === "percent") return formatPercentMetric(value, 1);
  return formatIntegerMetric(value);
}

function topLineMetricForCurrentView(scenario = activeScenario()) {
  const view = state.currentView;
  if (view === "profit") {
    return { label: "Profit", values: profitTopDownValues(scenario), format: "currency" };
  }
  if (view === "revenueDrilldown") {
    return { label: "Total Revenue", values: revenueTotalTopDownValues(scenario), format: "currency" };
  }
  if (view === "chart") {
    return { label: "LTR Revenue", values: ltrRevenueTopDownValues(scenario, calculateOutputs(scenario)), format: "currency" };
  }
  if (view === "growthRevenue") {
    return { label: "Growth Revenue", values: growthRevenueTopDownValues(scenario), format: "currency" };
  }
  if (view === "strRevenue") {
    return { label: "STR Revenue", values: strRevenueTopDownValues(scenario), format: "currency" };
  }
  if (view === "otherRevenue") {
    return { label: "Other Revenue", values: otherRevenueTopDownValues(scenario), format: "currency" };
  }
  if (view === "newCustomers") {
    if (state.newCustomerNewUnitsDrilldownOpen) {
      return {
        label: "New Units",
        values: YEARS.map(year => newCustomerUnitNewUnitsValue(scenario, year)),
        format: "integer",
      };
    }
    if (state.newCustomerUnitDrilldownOpen) {
      return {
        label: "New Property Customers",
        values: YEARS.map(year => newPropertyCohortValue(scenario.newCustomerDrilldown.counts, year)),
        format: "integer",
      };
    }
    return { label: "New Customers", values: driverValues(scenario, "newCustomers"), format: "integer" };
  }
  if (view === "costs") {
    if (state.costDrilldown === "labor") return { label: "Labor Cost", values: costValues(scenario, "labor"), format: "currency" };
    if (state.costDrilldown === "nonLabor") return { label: "Non-Labor Cost", values: costValues(scenario, "nonLabor"), format: "currency" };
    if (state.costDrilldown === "laborFte") return { label: "Labor Cost", values: costValues(scenario, "labor"), format: "currency" };
    if (state.costDrilldown === "nonLaborCategory") {
      const categoryKey = selectedNonLaborCategoryKey();
      return {
        label: nonLaborCategoryMeta[categoryKey]?.label || "Non-Labor Category",
        values: nonLaborCategoryValues(scenario, categoryKey),
        format: "currency",
      };
    }
    return { label: "Total Cost", values: costValues(scenario, "total"), format: "currency" };
  }
  return { label: "Profit", values: profitTopDownValues(scenario), format: "currency" };
}

function newCustomersTopMatchesBottomUp(scenario) {
  const topDown = driverValues(scenario, "newCustomers");
  const bottomUp = bottomUpNewCustomers(scenario);
  return YEARS.every((year, index) => {
    if (year < FIRST_FORECAST_YEAR) return true;
    return Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0)) < 1;
  });
}

function scenarioReconciliationChecks(scenario) {
  normalizeScenario(scenario);
  return [
    { label: "Profit", matched: profitValuesMatch(scenario), action: "profit" },
    { label: "Revenue", matched: revenueTopMatchesBottomUp(scenario), action: "revenueDrilldown" },
    { label: "LTR", matched: ltrTopMatchesBottomUp(scenario), action: "chart", focusChart: "revenue-chart" },
    { label: "Growth", matched: growthRevenueTopMatchesBottomUp(scenario), action: "growthRevenue" },
    { label: "STR", matched: strRevenueTopMatchesBottomUp(scenario), action: "strRevenue" },
    { label: "New Customers", matched: newCustomersTopMatchesBottomUp(scenario), action: "newCustomers" },
    { label: "Labor", matched: laborTopDownMatchesBottomUp(scenario), action: "laborCost" },
    { label: "Non-Labor", matched: nonLaborTopDownMatchesBottomUp(scenario), action: "nonLaborCost" },
    { label: "Labor FTE", matched: laborFteTopMatchesBottomUp(scenario), action: "laborFte" },
    ...NON_LABOR_CATEGORY_KEYS.map(key => ({
      label: nonLaborCategoryMeta[key]?.label || key,
      matched: nonLaborCategoryTopMatchesBottomUp(scenario, key),
      action: "nonLaborCategoryCost",
      categoryKey: key,
    })),
  ];
}

function scenarioMatchSummary(checks) {
  const unmatched = checks.filter(check => !check.matched);
  return {
    matched: unmatched.length === 0,
    label: unmatched.length ? `${unmatched.length} open` : "Matched",
    detail: unmatched.length ? unmatched.map(check => check.label).join(", ") : "All checked top-down and bottom-up definitions match.",
  };
}

function renderScenarioMatchStatus(checks) {
  const summary = scenarioMatchSummary(checks);
  return `
    <span class="scenario-status-pill ${summary.matched ? "matched" : "open"}">${summary.label}</span>
    <small>${escapeHtml(summary.detail)}</small>
  `;
}

function renderScenarioVersionRows(scenario) {
  const versions = (scenario.versions || [])
    .slice()
    .sort((left, right) => new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime());
  if (!versions.length) {
    return `<div class="scenario-version-empty">No saved versions yet.</div>`;
  }
  return `
    <table class="scenario-version-table">
      <thead>
        <tr>
          <th>Version</th>
          <th>Saved</th>
          <th>2025-2029 Cumulative Profit</th>
          <th>Top Down / Bottom Up</th>
        </tr>
      </thead>
      <tbody>
        ${versions.map((version, index) => {
          const versionScenario = scenarioFromVersion(scenario, version);
          const checks = scenarioReconciliationChecks(versionScenario);
          return `
            <tr>
              <td>${escapeHtml(isAnonymizedView() ? `Version ${index + 1}` : version.label || "Version")}</td>
              <td>${escapeHtml(formatScenarioSavedAt(version.createdAt))}</td>
              <td class="number-cell">${formatCurrency(scenarioCumulativeProfit(versionScenario), 0)}</td>
              <td>${renderScenarioMatchStatus(checks)}</td>
            </tr>
          `;
        }).join("")}
      </tbody>
    </table>
  `;
}

function renderScenariosModule() {
  const container = document.getElementById("scenarios-module");
  if (!container) return;
  const scenarios = Object.values(state.scenarios);
  container.innerHTML = `
    <div class="scenario-table-wrap">
      <table class="scenario-table">
        <thead>
          <tr>
            <th>Scenario</th>
            <th>2025-2029 Cumulative Profit</th>
            <th>Top Down / Bottom Up</th>
            <th>Latest Saved</th>
            <th>Versions</th>
          </tr>
        </thead>
        <tbody>
          ${scenarios.map(scenario => {
            normalizeScenario(scenario);
            const checks = scenarioReconciliationChecks(scenario);
            const cumulativeProfit = scenarioCumulativeProfit(scenario);
            const matchStatus = renderScenarioMatchStatus(checks);
            const versions = scenario.versions || [];
            const latestVersion = versions
              .slice()
              .sort((left, right) => new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime())[0];
            return `
              <tr class="scenario-parent-row">
                <td colspan="5">
                  <details class="scenario-details">
                    <summary>
                      <span>${escapeHtml(displayScenarioName(scenario))}</span>
                      <span class="number-cell">${formatCurrency(cumulativeProfit, 0)}</span>
                      <span>${matchStatus}</span>
                      <span>${escapeHtml(formatScenarioSavedAt(latestVersion?.createdAt))}</span>
                      <span>${versions.length.toLocaleString("en-US")}</span>
                    </summary>
                    <div class="scenario-version-panel">
                      ${renderScenarioVersionRows(scenario)}
                    </div>
                  </details>
                </td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
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

function renderAll({ syncNewCustomerEditors = true, excludeNewCustomerChart = null, syncRevenueDrilldown = true } = {}) {
  const outputs = calculateOutputs(activeScenario());
  if (syncRevenueDrilldown) syncRevenueDrilldownCharts();
  renderKpis(outputs);
  renderGuidedPlan(outputs);
  renderOutputCharts(outputs);
  renderRetentionDefense();
  renderInitiativesModule();
  renderProfitDrilldown();
  renderRevenueDrilldown(outputs);
  renderStrRevenueDrilldown();
  renderGrowthRevenueDrilldown();
  renderSaasMetrics(outputs);
  renderScenariosModule();
  renderMetricTree(outputs);
  renderCanvasFocus();
  renderTable(outputs);
  renderNewCustomerDrilldown({ syncEditors: syncNewCustomerEditors, excludeEditor: excludeNewCustomerChart });
  renderRevUnitPlan();
  renderRawDataSources();
  renderCostLockButtons();
  renderCostDrilldownView();
  renderLaborDepartmentMixEchart();
  renderNonLaborCategoryMixEchart();
  renderNonLaborCategoryDepartmentMixEchart();
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
    costs: clone(activeScenario().costs || baseCosts),
    costControlPoints: clone(activeScenario().costControlPoints || {}),
    costPctTotalLines: clone(activeScenario().costPctTotalLines || {}),
    costLocks: clone(activeScenario().costLocks || {}),
    newCustomerSource: activeScenario().newCustomerSource,
    newCustomerDrilldown: clone(activeScenario().newCustomerDrilldown),
    revUnitPlan: clone(activeScenario().revUnitPlan),
    revenuePaths: clone(activeScenario().revenuePaths || { growth: createDefaultGrowthRevenuePlan(), str: createDefaultStrRevenuePlan(), other: createDefaultOtherRevenuePlan() }),
    profitPlan: clone(activeScenario().profitPlan || createDefaultProfitPlan()),
    defenses: clone(activeScenario().defenses || createDefaultDefenses()),
  };
  addScenarioVersion(state.scenarios[id], "Save As");
  state.activeScenarioId = id;
  state.selectedVersionId = state.scenarios[id].versions[state.scenarios[id].versions.length - 1].id;
  setCompareScenario(previousActiveId);
  saveScenarios();
  renderScenarioSelect();
  syncDriverCharts();
  syncCostCharts();
  syncRevUnitCharts();
  syncRevenueDrilldownCharts();
  syncProfitCharts();
  syncOtherRevenueCharts();
  syncStrRevenueCharts();
  syncGrowthRevenueCharts();
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
  renderScenariosModule();
}

function restoreVersion() {
  const scenario = activeScenario();
  const version = (scenario.versions || []).find(item => item.id === state.selectedVersionId);
  if (!version) return;
  pushUndoSnapshot();
  scenario.drivers = clone(version.drivers);
  scenario.costs = version.costs
    ? clone(version.costs)
    : clone(baseCosts);
  scenario.costControlPoints = version.costControlPoints
    ? clone(version.costControlPoints)
    : {};
  scenario.costPctTotalLines = version.costPctTotalLines
    ? clone(version.costPctTotalLines)
    : {};
  scenario.costLocks = version.costLocks
    ? clone(version.costLocks)
    : {};
  scenario.newCustomerSource = version.newCustomerSource || "topDown";
  scenario.newCustomerDrilldown = version.newCustomerDrilldown
    ? clone(version.newCustomerDrilldown)
    : createDefaultNewCustomerDrilldown();
  scenario.revUnitPlan = version.revUnitPlan
    ? clone(version.revUnitPlan)
    : createDefaultRevUnitPlan();
  scenario.revenuePaths = version.revenuePaths
    ? clone(version.revenuePaths)
    : { total: createDefaultTotalRevenuePlan(), ltr: createDefaultLtrRevenuePlan(), growth: createDefaultGrowthRevenuePlan(), str: createDefaultStrRevenuePlan(), other: createDefaultOtherRevenuePlan() };
  scenario.profitPlan = version.profitPlan
    ? clone(version.profitPlan)
    : createDefaultProfitPlan();
  scenario.defenses = version.defenses
    ? clone(version.defenses)
    : createDefaultDefenses();
  enforceNewCustomerBaseline(scenario);
  saveScenarios();
  syncDriverCharts();
  syncCostCharts();
  syncRevUnitCharts();
  syncRevenueDrilldownCharts();
  syncProfitCharts();
  syncOtherRevenueCharts();
  syncStrRevenueCharts();
  syncGrowthRevenueCharts();
  renderAll();
}

function resetScenario() {
  pushUndoSnapshot();
  if (state.activeScenarioId === "base") {
    state.scenarios.base = makeBaseScenario();
  } else {
    activeScenario().drivers = clone(baseDrivers);
    activeScenario().costs = clone(baseCosts);
    activeScenario().costControlPoints = {};
    activeScenario().costPctTotalLines = {};
    activeScenario().costLocks = {};
    activeScenario().newCustomerSource = "topDown";
    activeScenario().newCustomerDrilldown = createDefaultNewCustomerDrilldown();
    activeScenario().revUnitPlan = createDefaultRevUnitPlan();
    activeScenario().revenuePaths = { total: createDefaultTotalRevenuePlan(), ltr: createDefaultLtrRevenuePlan(), growth: createDefaultGrowthRevenuePlan(), str: createDefaultStrRevenuePlan(), other: createDefaultOtherRevenuePlan() };
    activeScenario().profitPlan = createDefaultProfitPlan();
    activeScenario().defenses = createDefaultDefenses();
  }
  saveScenarios();
  syncDriverCharts();
  syncCostCharts();
  syncRevUnitCharts();
  syncRevenueDrilldownCharts();
  syncProfitCharts();
  syncOtherRevenueCharts();
  syncStrRevenueCharts();
  syncGrowthRevenueCharts();
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
  if (action === "definitions") {
    setView("definitions");
    focusElementById("definitions-view");
    return;
  }
  if (action === "profit") {
    setView("profit");
    syncProfitCharts();
    focusElementById("profit-view");
    return;
  }
  if (action === "revenueDrilldown") {
    setView("revenueDrilldown");
    syncRevenueDrilldownCharts();
    focusElementById("revenue-drilldown-view");
    return;
  }
  if (action === "growthRevenue") {
    setView("growthRevenue");
    syncGrowthRevenueCharts();
    focusElementById("growth-revenue-view");
    return;
  }
  if (action === "strRevenue") {
    setView("strRevenue");
    syncStrRevenueCharts();
    focusElementById("str-revenue-view");
    return;
  }
  if (action === "otherRevenue") {
    setView("otherRevenue");
    syncOtherRevenueCharts();
    focusElementById("other-revenue-view");
    return;
  }
  if (action === "costs") {
    state.costDrilldown = "";
    setView("costs");
    renderCostDrilldownView();
    syncCostCharts();
    focusElementById("cost-total-chart");
    return;
  }
  if (action === "laborCost") {
    state.costDrilldown = "labor";
    setView("costs");
    renderCostDrilldownView();
    syncCostCharts();
    focusElementById("cost-labor-drilldown");
    return;
  }
  if (action === "nonLaborCost") {
    state.costDrilldown = "nonLabor";
    setView("costs");
    renderCostDrilldownView();
    syncCostCharts();
    focusElementById("cost-non-labor-drilldown");
    return;
  }
  if (action === "nonLaborCategoryCost") {
    if (focusChart && NON_LABOR_CATEGORY_KEYS.includes(focusChart)) {
      state.selectedNonLaborCategoryKey = focusChart;
    }
    state.costDrilldown = "nonLaborCategory";
    setView("costs");
    renderCostDrilldownView();
    syncCostCharts();
    focusElementById("cost-non-labor-category-drilldown");
    return;
  }
  if (action === "laborFte") {
    state.costDrilldown = "laborFte";
    setView("costs");
    renderCostDrilldownView();
    syncCostCharts();
    focusElementById("cost-labor-fte-drilldown");
    return;
  }
  setView("chart");
  focusElementById(focusChart || "revenue-chart");
}

function openNapkinViewEntry() {
  setView("profit");
  syncProfitCharts();
  focusElementById("profit-view");
}

function resizeChartInstance(chart) {
  chart?.resize?.();
}

function resizeChartCollection(collection) {
  Object.values(collection).forEach(resizeChartInstance);
}

function resizeAllRenderedCharts() {
  resizeChartCollection(state.charts);
  resizeChartCollection(state.costCharts);
  resizeChartCollection(state.cohortCharts);
  resizeChartCollection(state.outputCharts);
  resizeChartCollection(state.revUnitCharts);
  resizeChartCollection(state.defenseCharts);
  resizeChartCollection(state.growthCharts);
  resizeChartCollection(state.strCharts);
  resizeChartCollection(state.otherRevenueCharts);
  resizeChartCollection(state.revenueDrilldownCharts);
  resizeChartCollection(state.profitCharts);
}

function scheduleViewResize(view) {
  setTimeout(() => {
    resizeAllRenderedCharts();
    requestAnimationFrame(resizeAllRenderedCharts);
  }, 0);
}

function bindControls() {
  bindNapkinYAxisControls();
  document.getElementById("anonymized-view-toggle")?.addEventListener("change", event => {
    state.anonymizedView = event.target.checked;
    syncDriverCharts();
    syncCostCharts();
    syncRevUnitCharts();
    syncRevenueDrilldownCharts();
    syncProfitCharts();
    syncOtherRevenueCharts();
    syncStrRevenueCharts();
    syncGrowthRevenueCharts();
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
    syncCostCharts();
    syncRevUnitCharts();
    syncRevenueDrilldownCharts();
    syncProfitCharts();
    syncOtherRevenueCharts();
    syncStrRevenueCharts();
    syncGrowthRevenueCharts();
    renderAll();
  });
  document.getElementById("compare-scenario-select").addEventListener("change", event => {
    setCompareScenario(event.target.value);
    syncDriverCharts();
    syncCostCharts();
    syncRevUnitCharts();
    syncRevenueDrilldownCharts();
    syncProfitCharts();
    syncOtherRevenueCharts();
    syncStrRevenueCharts();
    syncGrowthRevenueCharts();
    renderAll();
  });
  document.querySelectorAll("[data-cost-lock-key]").forEach(button => {
    button.addEventListener("click", event => {
      toggleCostLock(event.currentTarget.dataset.costLockKey);
    });
  });
  document.getElementById("toggle-cost-labor-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "labor";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("toggle-cost-labor-fte-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "laborFte";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("toggle-cost-non-labor-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "nonLabor";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("toggle-cost-non-labor-category-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "nonLaborCategory";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("back-cost-labor-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("back-cost-labor-fte-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "labor";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("back-cost-non-labor-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("back-cost-non-labor-category-drilldown")?.addEventListener("click", () => {
    state.costDrilldown = "nonLabor";
    renderCostDrilldownView();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("set-cost-labor-top-to-bottom")?.addEventListener("click", setLaborTopToBottom);
  document.getElementById("set-cost-labor-bottom-to-top")?.addEventListener("click", setLaborBottomToTop);
  document.getElementById("set-cost-labor-fte-top-to-bottom")?.addEventListener("click", setLaborFteTopToBottom);
  document.getElementById("set-cost-labor-fte-bottom-to-top")?.addEventListener("click", setLaborFteBottomToTop);
  document.getElementById("set-cost-non-labor-top-to-bottom")?.addEventListener("click", setNonLaborTopToBottom);
  document.getElementById("set-cost-non-labor-bottom-to-top")?.addEventListener("click", setNonLaborBottomToTop);
  document.getElementById("set-cost-non-labor-category-top-to-bottom")?.addEventListener("click", setNonLaborCategoryTopToBottom);
  document.getElementById("set-cost-non-labor-category-bottom-to-top")?.addEventListener("click", setNonLaborCategoryBottomToTop);
  document.getElementById("toggle-labor-allocation-top-lock")?.addEventListener("click", () => {
    state.laborAllocationTopLocked = !state.laborAllocationTopLocked;
    renderCostDrilldownView();
  });
  document.getElementById("toggle-labor-allocation-others-lock")?.addEventListener("click", () => {
    state.laborAllocationOthersLocked = !state.laborAllocationOthersLocked;
    renderCostDrilldownView();
  });
  document.getElementById("labor-allocation-line-list")?.addEventListener("change", event => {
    const input = event.target.closest("input[type='checkbox']");
    if (!input) return;
    const checked = Array.from(document.querySelectorAll("#labor-allocation-line-list input[type='checkbox']:checked"))
      .map(item => item.value)
      .filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
    state.selectedLaborAllocationKeys = checked;
    renderLaborAllocationControls();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("labor-fte-line-list")?.addEventListener("change", event => {
    const input = event.target.closest("input[type='checkbox']");
    if (!input) return;
    const checked = Array.from(document.querySelectorAll("#labor-fte-line-list input[type='checkbox']:checked"))
      .map(item => item.value)
      .filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
    state.selectedLaborFteKeys = checked;
    renderLaborFteControls();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("toggle-labor-fte-top-lock")?.addEventListener("click", () => {
    state.laborFteTopLocked = !state.laborFteTopLocked;
    renderCostDrilldownView();
  });
  document.querySelectorAll("[data-labor-fte-lock]").forEach(button => {
    button.addEventListener("click", event => {
      const lockKey = event.currentTarget.dataset.laborFteLock;
      state.laborFteLockedDriver = state.laborFteLockedDriver === lockKey ? "" : lockKey;
      renderCostDrilldownView();
    });
  });
  document.getElementById("toggle-non-labor-allocation-top-lock")?.addEventListener("click", () => {
    state.nonLaborAllocationTopLocked = !state.nonLaborAllocationTopLocked;
    renderCostDrilldownView();
  });
  document.getElementById("toggle-non-labor-allocation-others-lock")?.addEventListener("click", () => {
    state.nonLaborAllocationOthersLocked = !state.nonLaborAllocationOthersLocked;
    renderCostDrilldownView();
  });
  document.getElementById("non-labor-allocation-line-list")?.addEventListener("change", event => {
    const input = event.target.closest("input[type='checkbox']");
    if (!input) return;
    const checked = Array.from(document.querySelectorAll("#non-labor-allocation-line-list input[type='checkbox']:checked"))
      .map(item => item.value)
      .filter(key => NON_LABOR_CATEGORY_KEYS.includes(key));
    state.selectedNonLaborAllocationKeys = checked;
    renderNonLaborAllocationControls();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("non-labor-category-picker")?.addEventListener("change", event => {
    const input = event.target.closest("input[type='radio']");
    if (!input) return;
    const nextKey = input.value;
    if (!NON_LABOR_CATEGORY_KEYS.includes(nextKey)) return;
    state.selectedNonLaborCategoryKey = nextKey;
    state.selectedNonLaborCategoryDepartmentKeys = [];
    renderNonLaborCategoryControls();
    syncCostCharts();
    resizeCostCharts();
  });
  document.getElementById("toggle-non-labor-category-top-lock")?.addEventListener("click", () => {
    state.nonLaborCategoryTopLocked = !state.nonLaborCategoryTopLocked;
    renderCostDrilldownView();
  });
  document.getElementById("toggle-non-labor-category-others-lock")?.addEventListener("click", () => {
    state.nonLaborCategoryOthersLocked = !state.nonLaborCategoryOthersLocked;
    renderCostDrilldownView();
  });
  document.getElementById("non-labor-category-department-line-list")?.addEventListener("change", event => {
    const input = event.target.closest("input[type='checkbox']");
    if (!input) return;
    const checked = Array.from(document.querySelectorAll("#non-labor-category-department-line-list input[type='checkbox']:checked"))
      .map(item => item.value)
      .filter(key => LABOR_DEPARTMENT_KEYS.includes(key));
    state.selectedNonLaborCategoryDepartmentKeys = checked;
    renderNonLaborCategoryControls();
    syncCostCharts();
    resizeCostCharts();
  });
  document.querySelectorAll("[data-cost-chart-mode]").forEach(group => {
    group.addEventListener("click", event => {
      const button = event.target.closest("[data-cost-chart-mode-value]");
      if (!button || !group.contains(button)) return;
      setCostChartMode(group.dataset.costChartMode, button.dataset.costChartModeValue);
    });
  });
  document.querySelectorAll("[data-cost-line-mode]").forEach(group => {
    group.addEventListener("click", event => {
      const button = event.target.closest("[data-cost-line-mode-value]");
      if (!button || !group.contains(button)) return;
      setCostChartLineMode(group.dataset.costLineMode, button.dataset.costLineModeValue);
    });
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
  document.getElementById("set-ltr-bottom-to-top")?.addEventListener("click", setLtrBottomToTop);
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
  document.querySelectorAll("[data-defense-key]").forEach(button => {
    button.addEventListener("click", event => {
      event.stopPropagation();
      openLeafDefense(event.currentTarget.dataset.defenseKey);
    });
  });
  document.getElementById("back-retention-defense").addEventListener("click", () => {
    const config = activeDefenseConfig();
    openMetricTreeTarget(config.treeAction, config.focusChart);
  });
  document.getElementById("retention-initiative-form").addEventListener("submit", event => {
    event.preventDefault();
    const input = document.getElementById("retention-initiative-input");
    addRetentionInitiative(input.value);
    input.value = "";
    input.focus();
  });
  document.getElementById("retention-initiative-list").addEventListener("click", event => {
    const removeButton = event.target.closest("[data-remove-retention-initiative]");
    if (removeButton) {
      removeRetentionInitiative(Number(removeButton.dataset.removeRetentionInitiative));
      return;
    }
    const removeTeamButton = event.target.closest("[data-remove-retention-initiative-team]");
    if (removeTeamButton) {
      removeRetentionInitiativeTeam(
        Number(removeTeamButton.dataset.removeRetentionInitiativeTeam),
        removeTeamButton.dataset.team
      );
      return;
    }
    const showTeamPickerButton = event.target.closest("[data-show-retention-initiative-team-picker]");
    if (showTeamPickerButton) {
      const picker = showTeamPickerButton.parentElement?.querySelector(
        `[data-add-retention-initiative-team="${showTeamPickerButton.dataset.showRetentionInitiativeTeamPicker}"]`
      );
      if (picker) {
        picker.hidden = false;
        picker.focus();
      }
    }
  });
  document.getElementById("retention-initiative-list").addEventListener("change", event => {
    const picker = event.target.closest("[data-add-retention-initiative-team]");
    if (!picker || !picker.value) return;
    addRetentionInitiativeTeam(Number(picker.dataset.addRetentionInitiativeTeam), picker.value);
  });
  document.getElementById("retention-impact-line").addEventListener("click", () => {
    state.retentionImpactChartType = "line";
    renderRetentionRevenueImpact();
  });
  document.getElementById("retention-impact-bar").addEventListener("click", () => {
    state.retentionImpactChartType = "bar";
    renderRetentionRevenueImpact();
  });
  document.getElementById("retention-impact-cumulative").addEventListener("click", () => {
    state.retentionImpactCumulative = !state.retentionImpactCumulative;
    renderRetentionRevenueImpact();
  });

  document.getElementById("tab-tree").addEventListener("click", () => setView("tree"));
  document.getElementById("tab-definitions").addEventListener("click", () => setView("definitions"));
  document.getElementById("tab-saas-metrics").addEventListener("click", () => setView("saasMetrics"));
  document.getElementById("tab-scenarios").addEventListener("click", () => setView("scenarios"));
  document.getElementById("tab-initiatives").addEventListener("click", () => setView("initiatives"));
  document.getElementById("tab-guided").addEventListener("click", () => setView("guided"));
  document.getElementById("tab-costs").addEventListener("click", () => setView("costs"));
  document.getElementById("tab-chart").addEventListener("click", openNapkinViewEntry);
  document.getElementById("tab-table").addEventListener("click", () => setView("table"));
  document.getElementById("tab-rev-per-unit").addEventListener("click", () => setView("revPerUnit"));
  document.getElementById("tab-raw-data").addEventListener("click", () => setView("rawData"));
  document.getElementById("back-revenue-drilldown")?.addEventListener("click", () => setView("profit"));
  document.getElementById("set-total-revenue-bottom-to-top")?.addEventListener("click", setTotalRevenueBottomToTop);
  document.getElementById("back-ltr-revenue-drilldown")?.addEventListener("click", () => {
    setView("revenueDrilldown");
    syncRevenueDrilldownCharts();
  });
  document.getElementById("back-profit-drilldown")?.addEventListener("click", () => setView("tree"));
  document.getElementById("set-profit-top-to-bottom")?.addEventListener("click", setProfitTopToBottom);
  document.getElementById("set-profit-bottom-to-top")?.addEventListener("click", setProfitBottomToTop);
  document.querySelectorAll("[data-profit-lock-key]").forEach(button => {
    button.addEventListener("click", event => toggleProfitLock(event.currentTarget.dataset.profitLockKey));
  });
  document.getElementById("back-other-revenue-drilldown")?.addEventListener("click", () => setView("tree"));
  document.getElementById("back-str-revenue-drilldown")?.addEventListener("click", () => setView("tree"));
  document.getElementById("set-str-revenue-top-to-bottom")?.addEventListener("click", setStrRevenueTopToBottom);
  document.getElementById("set-str-revenue-bottom-to-top")?.addEventListener("click", setStrRevenueBottomToTop);
  document.getElementById("back-growth-revenue-drilldown")?.addEventListener("click", () => setView("tree"));
  document.getElementById("set-growth-revenue-top-to-bottom")?.addEventListener("click", setGrowthRevenueTopToBottom);
  document.getElementById("set-growth-revenue-bottom-to-top")?.addEventListener("click", setGrowthRevenueBottomToTop);
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
  document.getElementById("kpi-reconciliation-pills")?.addEventListener("click", event => {
    const button = event.target.closest("[data-reconciliation-action]");
    if (!button) return;
    const focusTarget = button.dataset.categoryKey || button.dataset.focusChart;
    openMetricTreeTarget(button.dataset.reconciliationAction, focusTarget);
  });
  document.getElementById("initiative-team-filter")?.addEventListener("change", renderInitiativesModule);
  document.getElementById("initiative-assumption-filter")?.addEventListener("change", renderInitiativesModule);
  document.getElementById("initiatives-view")?.addEventListener("click", event => {
    const button = event.target.closest("[data-initiative-defense-key]");
    if (!button) return;
    openLeafDefense(button.dataset.initiativeDefenseKey);
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
  const napkinViews = ["chart", "profit", "revenueDrilldown", "growthRevenue", "strRevenue", "otherRevenue"];
  document.getElementById("tab-tree").classList.toggle("active", view === "tree");
  document.getElementById("tab-definitions").classList.toggle("active", view === "definitions");
  document.getElementById("tab-saas-metrics").classList.toggle("active", view === "saasMetrics");
  document.getElementById("tab-scenarios").classList.toggle("active", view === "scenarios");
  document.getElementById("tab-initiatives").classList.toggle("active", view === "initiatives");
  document.getElementById("tab-guided").classList.toggle("active", view === "guided");
  document.getElementById("tab-costs").classList.toggle("active", view === "costs");
  document.getElementById("tab-chart").classList.toggle("active", napkinViews.includes(view));
  document.getElementById("tab-table").classList.toggle("active", view === "table");
  document.getElementById("tab-rev-per-unit").classList.toggle("active", view === "revPerUnit");
  document.getElementById("tab-raw-data").classList.toggle("active", view === "rawData");
  document.getElementById("tree-view").classList.toggle("active", view === "tree");
  document.getElementById("definitions-view").classList.toggle("active", view === "definitions");
  document.getElementById("revenue-drilldown-view").classList.toggle("active", view === "revenueDrilldown");
  document.getElementById("profit-view").classList.toggle("active", view === "profit");
  document.getElementById("other-revenue-view").classList.toggle("active", view === "otherRevenue");
  document.getElementById("str-revenue-view").classList.toggle("active", view === "strRevenue");
  document.getElementById("growth-revenue-view").classList.toggle("active", view === "growthRevenue");
  document.getElementById("saas-metrics-view").classList.toggle("active", view === "saasMetrics");
  document.getElementById("scenarios-view").classList.toggle("active", view === "scenarios");
  document.getElementById("initiatives-view").classList.toggle("active", view === "initiatives");
  document.getElementById("guided-view").classList.toggle("active", view === "guided");
  document.getElementById("costs-view").classList.toggle("active", view === "costs");
  document.getElementById("chart-view").classList.toggle("active", view === "chart");
  document.getElementById("retention-defense-view").classList.toggle("active", view === "retentionDefense");
  document.getElementById("table-view").classList.toggle("active", view === "table");
  document.getElementById("new-customers-view").classList.toggle("active", view === "newCustomers");
  document.getElementById("rev-per-unit-view").classList.toggle("active", view === "revPerUnit");
  document.getElementById("raw-data-view").classList.toggle("active", view === "rawData");
  renderNewCustomerUnitDrilldownToggle();
  if (view === "tree") {
    enterMetricTreeCanvas();
  }
  refreshHeaderKpis();
  scheduleViewResize(view);
}

function init() {
  renderScenarioSelect();
  bindControls();
  initOutputCharts();
  Object.keys(driverMeta).forEach(initDriverChart);
  Object.keys(costMeta).forEach(initCostChart);
  initLaborDrilldownTopDownChart();
  initLaborAllocationControlChart();
  initLaborAllocationPctChart();
  initLaborFteSelectedCostChart();
  initLaborFtePctChart("laborFteSelectedCost", "cost-labor-fte-selected-cost-pct-chart", applyLaborFteSelectedCostPctShares);
  initLaborFteCountChart();
  initLaborFtePctChart("laborFteCount", "cost-labor-fte-count-pct-chart", applyLaborFteCountPctShares);
  initLaborFteCostPerFteChart();
  initNonLaborDrilldownTopDownChart();
  initNonLaborAllocationControlChart();
  initNonLaborCategorySelectedSpendChart();
  initNonLaborCategoryDepartmentControlChart();
  initRevenueDrilldownCharts();
  initProfitCharts();
  initOtherRevenueCharts();
  initStrRevenueCharts();
  initGrowthRevenueCharts();
  initNewCustomerCohortEditors();
  initRevUnitCharts();
  initRetentionDefenseCharts();
  updateHistoryControls();
  renderAll();
}

init();
