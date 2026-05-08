# Next App Architecture

This document describes the intended architecture for a generalized v2 planning app. The current PetScreening app should remain a reference prototype. The next app should generalize the useful patterns into a reusable metric-tree operating model.

## Product Direction

The v2 app should be canvas-first.

Users should be able to:

- define a metric tree,
- create metrics on the fly from a catalog of construction patterns,
- compare scenarios and reference cases,
- inspect top-down vs bottom-up mismatches,
- edit appropriate metrics with NapkinChart-style controls,
- use standard charts for dense visual analysis,
- defend assumptions with initiatives, notes, and evidence,
- eventually translate annual targets into monthly, weekly, and daily operational targets.

The app should not be only a spreadsheet replacement. It should help users reason about how operating plans are constructed.

## Layered Architecture

Keep the layers strongly connected by data flow but weakly coupled in implementation.

```text
Metric Catalog
  defines what exists and how metrics are composed

Calculation Engine
  resolves values, time views, comparisons, and reconciliation

Scenario Store
  persists metric instances, edits, versions, references, and defense data

Canvas Renderer
  visualizes metric relationships and navigation

Chart Editors / Viewers
  render or edit metric series
```

## Metric Catalog

The metric catalog is the source of truth for business meaning.

It should define:

- metric id,
- label,
- unit,
- category,
- top-down definition,
- bottom-up definition,
- time behavior,
- allowed editors,
- allowed views,
- reconciliation behavior.

It should not know about DOM nodes, ECharts instances, or canvas coordinates.

## Calculation Engine

The calculation engine resolves metric values.

Example API shape:

```js
getMetricSeries({
  metricId: "revenue",
  scenarioId: "scenario-1",
  definition: "topDown", // topDown | bottomUp | effective
  grain: "year",
  view: "discrete" // discrete | cumulative | growth
})
```

The engine should own:

- formula execution,
- interpolation,
- aggregation,
- disaggregation,
- cumulative views,
- growth views,
- scenario comparison values,
- reconciliation checks.

Later, the engine can own seasonality and daily target expansion.

## Scenario Store

The scenario store should persist data, not UI implementation details.

It should store:

- scenario metadata,
- metric instances,
- sparse top-down control points,
- manual overrides,
- bottom-up input values,
- locks,
- version history,
- BAU/Target/reference snapshots,
- defense/initiative data.

It should avoid storing chart-specific state unless the state is part of the model, such as sparse control points.

## Canvas Renderer

The canvas should be generated from metric definitions and relationships.

The canvas owns:

- layout,
- zoom/pan,
- focus behavior,
- parent/child/sibling emphasis,
- mismatch badges,
- drilldown navigation,
- compact node charts.

The canvas should not calculate metrics itself. It should ask the calculation engine for series and status.

## Chart Layer

NapkinChart and ECharts should be used as components, not as business logic.

NapkinChart should handle:

- editable manual series,
- sparse control points,
- interpolation-friendly editing,
- assumption lines,
- target proxy lines,
- simple factor editing.

ECharts should handle:

- read-only trend charts,
- multi-series comparisons,
- stacked/area/bar charts,
- dense cohort displays,
- canvas node mini charts,
- reconciliation visuals.

Both should receive formatting and time behavior from shared metric presentation config.

## Metric Creation Flow

A user-created metric should begin from a constrained construction pattern, not an arbitrary formula editor.

Example flow:

1. Name the metric.
2. Choose unit.
3. Choose metric behavior: flow, stock, rate, ratio.
4. Choose construction type: manual, sum, product, ratio, allocation, cohort.
5. Choose input metrics.
6. Choose whether a top-down editable definition exists.
7. Choose allowed views: discrete, cumulative, growth, seasonality.
8. Place it in the canvas.

This provides flexibility without making the model arbitrary.

## Top-Down / Bottom-Up Contract

Top-down and bottom-up should be definitions, not chart modes.

A view can show:

- top-down only,
- bottom-up only,
- both lines,
- delta,
- reconciliation state,
- component drilldown.

An editor can:

- edit top-down,
- edit bottom-up components,
- set top-down to bottom-up,
- fit bottom-up to top-down,
- allocate a delta,
- lock specific branches.

## Time Behavior Contract

Each metric should declare whether the following are valid:

- cumulative view,
- growth view,
- monthly/weekly/daily disaggregation,
- seasonal allocation,
- daily target generation.

The app should avoid offering invalid views. For example, cumulative revenue can make sense; cumulative retention rate usually does not.

## Migration Strategy

Do not rewrite the current prototype in place.

Recommended path:

1. Keep the current app as `v1 prototype`.
2. Start a separate `v2/` implementation in this repo or a new repo.
3. Build the metric catalog and calculation engine first.
4. Add a simple canvas generated from definitions.
5. Add one or two editable metric patterns.
6. Port only the proven PetScreening patterns once the generic system can express them.

This reduces the risk of turning the current working app into a half-migrated app.
