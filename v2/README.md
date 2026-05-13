# V2 Local Server

The v2 sandbox is the generalized metric operating model prototype. The current
module split is intentionally light:

- `app.js`: browser state, rendering, DOM events, scenario editing, and chart wiring.
- `engine.js`: pure calculation helpers for coordinates, series math, and cohort matrices.
- `metric-catalog.js`: starter/default model construction.
- `examples/`: JSON models used as smoke tests for modeling patterns.

Current model-level planning primitives:

- `scenarioRoles`: names the editable `working` scenario and the current `bau` / `target` reference snapshot ids.
- `referenceScenarios`: immutable BAU / Target snapshots captured from a reconciled working scenario.
- `assumptions`: future home for metric-level claims, evidence, owners, and initiative notes.
- `reconciliationSummary`: derived export-only summary of top-down / bottom-up mismatches for the working scenario.
- `comparisonSummary`: derived export-only summary of Working vs BAU / Target deltas and attached assumptions.

The Delta Review panel is a thin UI reader on top of `comparisonSummary`; it does
not create additional model state.

The Metric Workspace is an additive focused editor: it shows the selected metric
and its immediate children. For sum formulas it can set the
parent to the current children or fit children to the parent by current child mix.

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
