# Metric Operating Model

This document defines the core domain model for a generalized v2 planning app. The goal is to make the app canvas-driven without hard-coding every metric, drilldown, and chart by hand.

PetScreening should be treated as the reference implementation and case study, not as the product boundary. Examples in this document are illustrative and should generalize to other operating models.

## Core Idea

A metric is not just a chart. A metric is a modeled object with:

- business meaning,
- values over time,
- optional top-down definition,
- optional bottom-up definition,
- time behavior,
- allowed views and editors,
- reconciliation rules.

The model should let users create and connect metrics on the fly while keeping the construction patterns understandable.

## Metric Definition

A `MetricDefinition` describes what a metric is.

```js
{
  id: "revenue",
  label: "Revenue",
  unit: "currency",
  category: "financial",

  topDown: {
    type: "manualSeries",
    editable: true,
    editor: "napkinLine"
  },

  bottomUp: {
    type: "sum",
    inputs: ["ltrRevenue", "growthRevenue", "strRevenue", "otherRevenue"]
  },

  reconciliation: {
    enabled: true,
    tolerance: 1,
    actions: ["setTopToBottom", "fitBottomToTop"]
  },

  time: {
    nativeGrain: "year",
    supportedGrains: ["year", "quarter", "month", "week", "day"],
    flowType: "flow",
    aggregateMethod: "sum",
    cumulativeAllowed: true,
    growthAllowed: true,
    seasonalityAllowed: true,
    dailyTargetAllowed: true
  },

  presentation: {
    defaultChart: "line",
    editableComponent: "napkinChart",
    canvasNodeSize: "normal"
  }
}
```

## Metric Instance

A `MetricInstance` is scenario-specific. It stores values, sparse control points, edits, version state, and metadata for one metric in one scenario.

```js
{
  scenarioId: "scenario-1",
  metricId: "revenue",
  topDownControlPoints: [[2026, 66845700], [2029, 180563000]],
  manualOverrides: {},
  defense: {},
  updatedAt: "..."
}
```

Definitions are reusable. Instances are scenario/version-specific.

## Top-Down And Bottom-Up

Top-down and bottom-up should be first-class definitions.

Top-down:

- usually an editable manual series,
- expresses the executive/user belief or target,
- often edited with a NapkinChart.

Bottom-up:

- calculated from components,
- may be editable indirectly through its children,
- explains whether the top-down case is defensible.

Some metrics have both definitions:

```js
profit = {
  topDown: { type: "manualSeries" },
  bottomUp: { type: "difference", inputs: ["revenue", "cost"] }
}
```

Some metrics are top-down only:

```js
otherRevenue = {
  topDown: { type: "manualSeries" },
  bottomUp: null
}
```

Some metrics are calculated only:

```js
spendPerNewUnit = {
  topDown: null,
  bottomUp: { type: "ratio", numerator: "selectedCostPool", denominator: "newUnits" }
}
```

## Composition Types

The app should support a small catalog of composition types before it supports arbitrary formulas.

| Type | Meaning | Example |
| --- | --- | --- |
| `manualSeries` | User-defined time series | Other Revenue |
| `sum` | Sum of child metrics | Total Revenue = Product A + Product B + Services |
| `difference` | One metric less another | Profit = Revenue - Cost |
| `product` | Product of two metrics | Growth Revenue = New Customers * Rev / Customer |
| `ratio` | Numerator divided by denominator | Spend / New Unit |
| `weightedAverage` | Weighted average of child ratios | Revenue / Customer |
| `cohortSum` | Sum of cohort rows over time | Customers by acquisition cohort |
| `allocation` | Values distributed by percent total | Initiative impact, cost mix |
| `proxy` | Decision aid derived from other metrics | Revenue per unit, spend per unit |

These patterns should drive available editors. Users should not need to write formulas in the first version.

## Time Behavior

Metrics need to know whether cumulative, growth, and seasonality views make sense.

The key distinction is flow vs stock vs rate/ratio.

Flow metrics can be summed through time:

- Revenue
- Cost
- Profit
- New customers
- New units
- Transactions

Stock metrics represent a point-in-time or average state:

- Active customers
- Ending properties
- FTE count

Rate/ratio metrics usually should not be cumulatively summed:

- Retention rate
- Rev / customer
- Spend / unit
- Cost / FTE

Example time behavior:

```js
newCustomers.time = {
  flowType: "flow",
  aggregateMethod: "sum",
  cumulativeAllowed: true,
  growthAllowed: true,
  seasonalityAllowed: true,
  dailyTargetAllowed: true
}

retention.time = {
  flowType: "rate",
  aggregateMethod: "weightedAverage",
  cumulativeAllowed: false,
  growthAllowed: true,
  seasonalityAllowed: false,
  dailyTargetAllowed: false
}
```

## Seasonality And Daily Targets

Daily modeling does not need to be built first, but the metric schema should allow it.

Annual planning should eventually translate into monthly, weekly, and daily operational targets.

```js
disaggregation = {
  from: "year",
  to: "day",
  method: "seasonality",
  components: {
    monthOfYear: "monthlyPattern",
    dayOfWeek: "weekdayPattern",
    holiday: "holidayAdjustment"
  },
  normalizeToAnnualTotal: true
}
```

This should be valid only for metrics where daily targeting makes sense.

## Reconciliation

For metrics with both top-down and bottom-up definitions:

- show both series,
- flag mismatches,
- allow setting top-down to bottom-up,
- allow fitting bottom-up to top-down when a safe rule exists,
- preserve locks when applying changes.

Mismatch visibility is a product feature. The app should not silently hide top-down/bottom-up disagreement.

## Editing Surfaces

Metric type should imply the editor:

| Metric Pattern | Preferred Editor |
| --- | --- |
| `manualSeries` | NapkinChart line |
| `sum` | child allocation, set top to bottom, fit bottom to top |
| `product` | edit one factor, lock one factor, fit one factor |
| `ratio` | target ratio and implied numerator/denominator |
| `cohortSum` | cohort table plus staged NapkinChart editors |
| `allocation` | NapkinChartArea pct-total |

NapkinChart should own editable series. ECharts should primarily own read-only visualization, comparison, and dense multi-series displays.
