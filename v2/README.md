# V2 Local Server

The v2 sandbox is the generalized metric operating model prototype. The current
module split is intentionally light:

- `app.js`: browser state, rendering, DOM events, scenario editing, and chart wiring.
- `engine.js`: pure calculation helpers for coordinates, series math, and cohort matrices.
- `metric-catalog.js`: starter/default model construction.
- `examples/`: JSON models used as smoke tests for modeling patterns.

Current model-level planning primitives:

- `scenarioRoles`: names the editable `working` scenario. Whole-scenario BAU / Target roles are legacy-only.
- `metricReferences`: current BAU / Target reference lines captured per metric from a reconciled working metric.
- `referenceScenarios`: legacy whole-scenario BAU / Target snapshots. New comparison behavior should use `metricReferences`.
- `assumptions`: future home for metric-level claims, evidence, owners, and initiative notes.
- `reconciliationSummary`: derived export-only summary of top-down / bottom-up mismatches for the working scenario.
- `comparisonSummary`: derived export-only summary of Working vs metric-level BAU / Target deltas and attached assumptions.

Metric references are intentionally current-state only for now:

```json
{
  "metricReferences": {
    "profit": {
      "bau": {
        "metricId": "profit",
        "series": [14000000, 16500000, 19000000, 45000000],
        "sourceScenarioId": "scenario-1",
        "sourceScenarioName": "Scenario 1",
        "setAt": "2026-05-13T12:00:00.000Z"
      }
    }
  }
}
```

History can be added later by wrapping each role in `currentVersionId` plus a
`history` array.

The Delta Review panel is a thin UI reader on top of `comparisonSummary`; it does
not create additional model state.

The Metric Workspace is an additive focused editor: it shows the selected metric
and its immediate children. For sum formulas it can set the
parent to the current children or fit children to the parent by current child mix.

Active bottom-up formula grammar:

- `manual`: no bottom-up definition; the metric is just an editable top-down series.
- `sum`: adds one or more input metric top-down series.
- `product`: multiplies two or more input metric top-down series.
- `difference`: subtracts the second input from the first input.
- `weightedAverage`: rolls a rate metric up using a selected weight metric.
- `laggedMetric`: derives a metric from another metric shifted by one or more periods, with opening values for the first visible period.
- `cohortMatrix`: generates a cohort-year matrix from starting cohorts and an age curve, then rolls it up to annual values.

Older experimental formula types such as ratio, cumulative net flow, cohort age
product, and cohort matrix sum are intentionally not exposed in the active
builder. Stock-flow patterns should be modeled with `laggedMetric` plus ordinary
sum/difference relationships.

Assumptions attach to a metric and optional coordinate scope:

```json
{
  "id": "assumption-1",
  "metricId": "new-paying-customers",
  "scope": { "coordinate": {} },
  "baselineRole": "bau",
  "comparisonRole": "target",
  "years": [2026, 2027, 2028, 2029],
  "claim": "Channel investment increases new customers.",
  "driverType": "initiative",
  "owner": null,
  "impact": {
    "mode": "pctOfDelta",
    "pct": 35
  },
  "evidence": [],
  "status": "draft"
}
```

Run the v2 sandbox from the repo root:

```bash
node v2/server.mjs
```

Then open:

```text
http://127.0.0.1:8020/
```

To use a different port:

```bash
PORT=8021 node v2/server.mjs
```
