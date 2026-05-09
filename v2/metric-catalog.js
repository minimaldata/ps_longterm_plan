export const YEARS = [2026, 2027, 2028, 2029];

export function createBlankModel() {
  return {
    name: "Blank Model",
    dimensions: {},
    metricDefinitions: {},
    scenario: {
      id: "scenario-1",
      name: "Scenario 1",
      metrics: {},
    },
  };
}

export function createStarterModel() {
  return {
    name: "Starter Example",
    dimensions: {},
    metricDefinitions: clone(starterMetricDefinitions),
    scenario: clone(starterScenario),
  };
}

export function createManualMetric({ id, label, unit = "currency", color = "#4f7fb8" }) {
  return {
    id,
    label,
    unit,
    topDown: { type: "manualSeries", editable: true },
    bottomUp: null,
    reconciliation: { enabled: false, tolerance: 1 },
    time: {
      nativeGrain: "year",
      supportedGrains: ["year"],
      flowType: unit === "percent" ? "rate" : "flow",
      aggregateMethod: unit === "percent" ? "weightedAverage" : "sum",
      cumulativeAllowed: unit !== "percent",
      growthAllowed: true,
      seasonalityAllowed: unit !== "percent",
      dailyTargetAllowed: unit !== "percent",
    },
    presentation: { color },
  };
}

export function defaultSeries() {
  return YEARS.map(() => 0);
}

const starterMetricDefinitions = {
  profit: {
    id: "profit",
    label: "Profit",
    unit: "currency",
    topDown: { type: "manualSeries", editable: true },
    bottomUp: { type: "difference", inputs: ["revenue", "cost"] },
    reconciliation: { enabled: true, tolerance: 1 },
    time: {
      nativeGrain: "year",
      supportedGrains: ["year"],
      flowType: "flow",
      aggregateMethod: "sum",
      cumulativeAllowed: true,
      growthAllowed: true,
      seasonalityAllowed: true,
      dailyTargetAllowed: true,
    },
    presentation: { color: "#2f6f73" },
  },
  revenue: {
    id: "revenue",
    label: "Revenue",
    unit: "currency",
    topDown: { type: "manualSeries", editable: true },
    bottomUp: { type: "sum", inputs: ["productARevenue", "productBRevenue", "servicesRevenue"] },
    reconciliation: { enabled: true, tolerance: 1 },
    time: {
      nativeGrain: "year",
      supportedGrains: ["year"],
      flowType: "flow",
      aggregateMethod: "sum",
      cumulativeAllowed: true,
      growthAllowed: true,
      seasonalityAllowed: true,
      dailyTargetAllowed: true,
    },
    presentation: { color: "#4f7fb8" },
  },
  cost: {
    id: "cost",
    label: "Cost",
    unit: "currency",
    topDown: { type: "manualSeries", editable: true },
    bottomUp: { type: "sum", inputs: ["laborCost", "nonLaborCost"] },
    reconciliation: { enabled: true, tolerance: 1 },
    time: {
      nativeGrain: "year",
      supportedGrains: ["year"],
      flowType: "flow",
      aggregateMethod: "sum",
      cumulativeAllowed: true,
      growthAllowed: true,
      seasonalityAllowed: true,
      dailyTargetAllowed: true,
    },
    presentation: { color: "#6fa76b" },
  },
  productARevenue: createManualMetric({ id: "productARevenue", label: "Product A Revenue", color: "#7b6fb8" }),
  productBRevenue: createManualMetric({ id: "productBRevenue", label: "Product B Revenue", color: "#d39b2a" }),
  servicesRevenue: createManualMetric({ id: "servicesRevenue", label: "Services Revenue", color: "#4b9b8e" }),
  laborCost: createManualMetric({ id: "laborCost", label: "Labor Cost", color: "#5c8fb7" }),
  nonLaborCost: createManualMetric({ id: "nonLaborCost", label: "Non-Labor Cost", color: "#a96c50" }),
};

const starterScenario = {
  id: "scenario-1",
  name: "Scenario 1",
  metrics: {
    profit: { topDown: [14000000, 22000000, 33000000, 45000000] },
    revenue: { topDown: [68000000, 85000000, 105000000, 130000000] },
    cost: { topDown: [54000000, 63000000, 72000000, 85000000] },
    productARevenue: { topDown: [36000000, 44000000, 53000000, 64000000] },
    productBRevenue: { topDown: [18000000, 24000000, 31000000, 39000000] },
    servicesRevenue: { topDown: [11000000, 15000000, 19000000, 24000000] },
    laborCost: { topDown: [32000000, 37000000, 43000000, 50000000] },
    nonLaborCost: { topDown: [20000000, 24000000, 29000000, 34000000] },
  },
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}
