# Prototype Lessons

This document captures lessons from the current PetScreening planning prototype. These lessons should inform the generalized v2 app, but they should not hard-code the v2 product to PetScreening-specific terminology.

## What Worked

### NapkinChart Is A Strong Starting Point

Editable lines are a good executive planning surface. They let users express a directional belief quickly without editing every spreadsheet cell.

The best use cases are:

- top-down target lines,
- direct leaf assumptions,
- simple factors,
- proxy targets,
- one selected aggregate line at a time.

### Canvas View Should Become Primary

The canvas made the model easier to discuss because it showed metric relationships, not just chart panels.

The next app should make the canvas the main organizing surface:

- parent/child metric relationships,
- drilldown navigation,
- top-down/bottom-up mismatch badges,
- comparison lines,
- focus behavior.

### Top-Down vs Bottom-Up Is Core

The strongest pattern was allowing a metric to have both:

- a top-down definition,
- a bottom-up calculated definition.

This made it possible to ask: "What do we want to believe?" and then "Can we defend it?"

### Reference Scenarios Matter

BAU and Target became more useful as reference snapshots than as editable working scenarios.

For v2, references should be first-class:

- immutable snapshots,
- version history,
- usable as compare lines,
- assignable from saved working scenarios only when rules are satisfied.

### Defense Belongs Near Assumptions

The Defend workflow is most useful around leaf assumptions or assumptions without a detailed bottom-up model.

Defense should connect:

- assumption deltas,
- initiatives,
- owners/teams,
- revenue/cost impact,
- notes or evidence.

### Proxy Metrics Are Valuable

Proxy metrics help users make better decisions when the direct causal model is too complicated.

Examples from the prototype:

- revenue per unit,
- spend per new unit,
- revenue per FTE,
- revenue per new paying customer.

The v2 app should support proxy metrics as first-class calculated or editable targets.

## What Became Too Complex

### Drilldowns Can Go Too Deep Too Fast

The new customer drilldown became powerful but overwhelming. Users need staged disclosure:

1. show the simple top-level split,
2. then offer a deeper drilldown,
3. then offer cohort/all-year detail only when needed.

### Too Many Editable Lines Are Hard To Use

Multi-line NapkinCharts can become difficult when many lines are editable. Good patterns are:

- edit one selected aggregate at a time,
- use pills or filters to select which lines are represented,
- use ECharts for dense multi-line context,
- use pct-total area charts for mix allocation.

### Manual Sync Logic Became Fragile

The current app has many hand-wired render/sync paths. Bugs often came from lifecycle issues:

- rebuilding the chart being dragged,
- losing historical points,
- materializing deleted points,
- stale tooltips,
- hidden-tab resize failures.

V2 should centralize chart lifecycle and metric patching.

### Business Definitions Drift Without A Catalog

As the model grew, the app accumulated implicit definitions in code. The next version needs a metric catalog so definitions are explicit and inspectable.

## Interaction Rules To Preserve

### Sparse Control Points

Deleted NapkinChart points should remain deleted. Interpolation should fill the visual/model value between remaining points.

Only add a deleted point back when another edit intentionally changes that exact year.

### Do Not Rewrite The Active Drag Chart

During a drag, dependent charts may update. The chart being dragged should usually not be rebuilt until the interaction finishes.

### Historical Values Should Stay Visible

Even if only forecast years are editable, historical years should remain visible in most charts.

Use edit domains to prevent historical editing instead of removing historical points.

### Mismatches Should Be Visible

Top-down and bottom-up mismatches should appear as a status, not as a hidden data problem.

The app should give users clear actions:

- set top-down to bottom-up,
- fit bottom-up to top-down,
- inspect the component branch,
- defend the gap.

## Storage Lessons

Local browser storage was useful for prototyping but confusing as the app grew.

Useful short-term bridges:

- JSON export/import,
- Excel export with formulas,
- reference scenario snapshots.

Likely v2 needs a real storage model eventually, but the first generalized version can still avoid users/databases if the workflow is consultant-led.

## Generalization Lessons

The PetScreening prototype should be treated as a demanding case study:

- multiple revenue streams,
- costs,
- proxy metrics,
- top-down and bottom-up definitions,
- reference scenarios,
- assumption defense,
- canvas navigation.

The generalized product should not assume PetScreening-specific categories. Instead, it should support the patterns that PetScreening exposed.

The reusable product is not "a PetScreening model." It is a canvas-based operating model builder where metrics can be defined, connected, reconciled, defended, and translated into operating targets.
