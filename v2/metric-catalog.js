export const YEARS = [2026, 2027, 2028, 2029];

export const metricDefinitions = {
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
  productARevenue: leafMetric("productARevenue", "Product A Revenue", "#7b6fb8"),
  productBRevenue: leafMetric("productBRevenue", "Product B Revenue", "#d39b2a"),
  servicesRevenue: leafMetric("servicesRevenue", "Services Revenue", "#4b9b8e"),
  laborCost: leafMetric("laborCost", "Labor Cost", "#5c8fb7"),
  nonLaborCost: leafMetric("nonLaborCost", "Non-Labor Cost", "#a96c50"),
};

export const scenario = {
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

function leafMetric(id, label, color) {
  return {
    id,
    label,
    unit: "currency",
    topDown: { type: "manualSeries", editable: true },
    bottomUp: null,
    reconciliation: { enabled: false, tolerance: 1 },
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
    presentation: { color },
  };
}
