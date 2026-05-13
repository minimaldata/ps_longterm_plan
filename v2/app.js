import {
  YEARS as DEFAULT_YEARS,
  createBlankModel,
  createManualMetric,
  createStarterModel,
} from "./metric-catalog.js";
import {
  buildGenericCohortMatrix,
  cohortCurveValueForAge,
  cohortStartValueForYear,
  cohortValueForYear,
  cohortYoyValueForAge,
  coordinateKey,
  coordinateMatches,
  flatOpeningCohorts,
  matrixTotalSeries,
  memberCoordinateKey,
  parseCoordinateKey,
  rollupCohortMatrices,
  setEngineYears,
  sortedCohortMatrixRowKeys,
  splitOpeningCohortsByCoordinate,
  splitSeries,
  sumSeries,
} from "./engine.js";

let YEARS = [...DEFAULT_YEARS];

function defaultSeries() {
  return YEARS.map(() => 0);
}

const state = {
  model: createBlankModel(),
  selectedMetricId: null,
  selectedDimensionContext: null,
  selectedDimensionId: null,
  topDownEditorView: "chart",
  canvasViewMode: "number",
  metricWorkspaceMode: "total",
  dimensionGroupTool: "red",
  dimensionMixView: "pctTotal",
  dimensionGroupLinesShowGrouped: true,
  dimensionGroupLinesShowUngrouped: true,
  dimensionGroupLinesChartType: "line",
  topDownSheetExpandedKeys: new Set(),
  workspaceLockedKeys: new Set(),
  cohortBuilderIsDirty: false,
  jsonExplorerExpandedPaths: new Set(["", "/metricDefinitions", "/scenarios"]),
};

const currencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 0,
});
const preciseCurrencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 0,
  maximumFractionDigits: 2,
});

const dimensionPalette = ["#2f6f73", "#4f7fb8", "#7b6fb8", "#d39b2a", "#a96c50", "#6fa76b"];
const dimensionGroupDefinitions = [
  { id: "red", label: "Red", color: "#c2410c" },
  { id: "blue", label: "Blue", color: "#2563eb" },
  { id: "other", label: "Other", color: "#98a2b3" },
];
const TOTAL_COORDINATE_KEY = "";
const UNASSIGNED_MEMBER_ID = "unassigned";

function workspaceChildLockKey(parentMetricId, childId) {
  return `child::${parentMetricId || ""}::${childId || ""}`;
}

function workspaceDimensionGroupLockKey(mixContext, groupId) {
  return `dimension::${mixContext?.metricId || ""}::${dimensionMixControlKey(mixContext)}::${groupId || ""}`;
}

function workspaceIsLocked(lockKey) {
  return Boolean(lockKey && state.workspaceLockedKeys.has(lockKey));
}

function toggleWorkspaceLock(lockKey) {
  if (!lockKey) return;
  if (state.workspaceLockedKeys.has(lockKey)) {
    state.workspaceLockedKeys.delete(lockKey);
  } else {
    state.workspaceLockedKeys.add(lockKey);
  }
}

function workspaceLockButtonHtml(lockKey, { disabled = false } = {}) {
  const locked = workspaceIsLocked(lockKey);
  return `<button type="button" class="workspace-lock-button ${locked ? "locked" : ""}" data-workspace-lock-key="${escapeHtml(lockKey)}" ${disabled ? "disabled" : ""}>${locked ? "Locked" : "Lock"}</button>`;
}

const chartElement = document.getElementById("metric-detail-chart");
const chart = window.echarts
  ? echarts.init(chartElement)
  : { clear() {}, setOption() {}, resize() {} };
const cohortStartsChartElement = document.getElementById("cohort-starts-chart");
const cohortYoyChartElement = document.getElementById("cohort-yoy-chart");
const cohortStartsChart = window.echarts
  ? echarts.init(cohortStartsChartElement)
  : { clear() {}, setOption() {}, resize() {} };
const cohortYoyChart = window.echarts
  ? echarts.init(cohortYoyChartElement)
  : { clear() {}, setOption() {}, resize() {} };
const dimensionBreakdownChartElement = document.getElementById("dimension-breakdown-chart");
const dimensionBreakdownChart = window.echarts
  ? echarts.init(dimensionBreakdownChartElement)
  : { clear() {}, setOption() {}, resize() {} };
let topDownNapkinChart = null;
let topDownNapkinMetricId = null;
let topDownNapkinIsSyncing = false;
let topDownNapkinHasPendingCommit = false;
let workspaceNapkinChart = null;
let workspaceNapkinMetricId = null;
let workspaceNapkinIsSyncing = false;
let workspaceNapkinHasPendingCommit = false;
let workspaceNapkinReferenceStyleFrame = null;
let dimensionMixChart = null;
let dimensionMixMetricId = null;
let dimensionMixIsSyncing = false;
let dimensionMixHasPendingCommit = false;
let workspaceDimensionMixChart = null;
let workspaceDimensionMixMetricId = null;
let workspaceDimensionMixIsSyncing = false;
let workspaceDimensionMixHasPendingCommit = false;
let workspaceDimensionMemberCharts = new Map();
let workspaceDimensionMemberChartSyncing = new Set();
let workspaceDimensionMemberChartPendingCommits = new Set();
let workspaceChildNapkinCharts = new Map();
let workspaceChildNapkinChartSyncing = new Set();
let workspaceChildNapkinChartPendingCommits = new Set();
let workspaceLiveRefreshFrame = null;
let pendingWorkspaceLiveRefreshOptions = null;
let metricWorkspaceRenderToken = 0;
let metricWorkspaceCharts = [];
let catalogSourceIsDirty = false;
const calculationCache = new Map();

function clearCalculationCache() {
  calculationCache.clear();
}

function cachedCalculation(key, factory) {
  if (calculationCache.has(key)) return calculationCache.get(key);
  const value = factory();
  calculationCache.set(key, value);
  return value;
}

function metricDefinitions() {
  return state.model.metricDefinitions;
}

function dimensionDefinitions() {
  if (!state.model.dimensions) state.model.dimensions = {};
  return state.model.dimensions;
}

function normalizeYears(years) {
  const normalized = [...new Set((years || [])
    .map(year => Number(year))
    .filter(year => Number.isInteger(year)))]
    .sort((left, right) => left - right);
  return normalized.length ? normalized : [...DEFAULT_YEARS];
}

function modelYears() {
  return normalizeYears(state.model.years);
}

function syncActiveYears() {
  YEARS = modelYears();
  state.model.years = [...YEARS];
  setEngineYears(YEARS);
}

function yearsFromRange(startYear, endYear) {
  const start = Number(startYear);
  const end = Number(endYear);
  if (!Number.isInteger(start) || !Number.isInteger(end)) {
    throw new Error("Start year and end year must be whole years.");
  }
  if (start > end) {
    throw new Error("Start year must be before end year.");
  }
  const years = [];
  for (let year = start; year <= end; year += 1) years.push(year);
  return years;
}

function migrateSeriesByYear(series = [], oldYears = YEARS, newYears = YEARS, fillValue = 0) {
  const byYear = new Map(oldYears.map((year, index) => [Number(year), Number(series?.[index] ?? fillValue)]));
  return newYears.map(year => Number(byYear.has(Number(year)) ? byYear.get(Number(year)) : fillValue));
}

function migrateTopDownByYear(rawTopDown, oldYears, newYears) {
  if (Array.isArray(rawTopDown)) return migrateSeriesByYear(rawTopDown, oldYears, newYears);
  if (!rawTopDown || typeof rawTopDown !== "object") return { [TOTAL_COORDINATE_KEY]: newYears.map(() => 0) };
  return Object.fromEntries(Object.entries(rawTopDown)
    .filter(([_key, series]) => Array.isArray(series))
    .map(([key, series]) => [key, migrateSeriesByYear(series, oldYears, newYears)]));
}

function migratePointLineByYear(points = [], newYears = YEARS) {
  const normalized = normalizeNapkinPoints(points);
  return newYears.map(year => [year, interpolateLineValue(normalized, year)]);
}

function migratePointMapByYear(pointMap, newYears) {
  if (!pointMap || typeof pointMap !== "object") return pointMap;
  return Object.fromEntries(Object.entries(pointMap)
    .filter(([_key, points]) => Array.isArray(points))
    .map(([key, points]) => [key, migratePointLineByYear(points, newYears)]));
}

function migrateYearObjectByYear(values, newYears, fillValue = 0) {
  if (!values || typeof values !== "object" || Array.isArray(values)) return values;
  return Object.fromEntries(newYears.map(year => [
    String(year),
    Number(Object.prototype.hasOwnProperty.call(values, String(year)) ? values[String(year)] : fillValue),
  ]));
}

function migrateCohortsByYear(cohorts, newYears) {
  if (!cohorts || typeof cohorts !== "object" || Array.isArray(cohorts)) return cohorts;
  return Object.fromEntries(Object.entries(cohorts).map(([cohortKey, row]) => [
    cohortKey,
    migrateYearObjectByYear(row, newYears),
  ]));
}

function migrateMetricScenarioYears(metric = {}, oldYears, newYears) {
  const nextMetric = clone(metric);
  if (Object.prototype.hasOwnProperty.call(nextMetric, "topDown")) {
    nextMetric.topDown = migrateTopDownByYear(nextMetric.topDown, oldYears, newYears);
  }
  if (nextMetric.controlPoints) {
    nextMetric.controlPoints = migratePointMapByYear(nextMetric.controlPoints, newYears);
  }
  if (nextMetric.topDown && typeof nextMetric.topDown === "object" && !Array.isArray(nextMetric.topDown) && nextMetric.controlPoints) {
    const topDownKeys = Object.keys(nextMetric.topDown);
    Object.entries(nextMetric.controlPoints).forEach(([key, points]) => {
      const isStoredTopDownKey = Object.prototype.hasOwnProperty.call(nextMetric.topDown, key);
      const isOnlyTotal = key === TOTAL_COORDINATE_KEY && topDownKeys.length <= 1;
      if (!isStoredTopDownKey && !isOnlyTotal) return;
      nextMetric.topDown[key] = newYears.map(year => interpolateLineValue(points, year));
    });
  }
  if (Array.isArray(nextMetric.manualControlPoints)) {
    nextMetric.manualControlPoints = migratePointLineByYear(nextMetric.manualControlPoints, newYears);
    if (Array.isArray(nextMetric.topDown)) {
      nextMetric.topDown = newYears.map(year => interpolateLineValue(nextMetric.manualControlPoints, year));
    } else if (
      nextMetric.topDown
      && typeof nextMetric.topDown === "object"
      && Object.keys(nextMetric.topDown).every(key => key === TOTAL_COORDINATE_KEY)
    ) {
      nextMetric.topDown[TOTAL_COORDINATE_KEY] = newYears.map(year => interpolateLineValue(nextMetric.manualControlPoints, year));
    }
  }
  if (nextMetric.memberManualControlPoints) {
    nextMetric.memberManualControlPoints = migratePointMapByYear(nextMetric.memberManualControlPoints, newYears);
  }
  if (Array.isArray(nextMetric.dimensionShareControlPoints)) {
    nextMetric.dimensionShareControlPoints = nextMetric.dimensionShareControlPoints
      .map(points => migratePointLineByYear(points, newYears));
  }
  if (nextMetric.controlPoints && nextMetric.controlPoints[TOTAL_COORDINATE_KEY]) {
    nextMetric.manualControlPoints = nextMetric.controlPoints[TOTAL_COORDINATE_KEY];
  }
  if (nextMetric.cohortStarts) {
    nextMetric.cohortStarts = migrateYearObjectByYear(nextMetric.cohortStarts, newYears);
  }
  if (nextMetric.cohorts) {
    nextMetric.cohorts = migrateCohortsByYear(nextMetric.cohorts, newYears);
  }
  return nextMetric;
}

function migrateModelYears(newYears) {
  const oldYears = [...YEARS];
  const normalizedYears = normalizeYears(newYears);
  if (oldYears.length === normalizedYears.length && oldYears.every((year, index) => year === normalizedYears[index])) {
    return false;
  }
  Object.values(scenarioCollection()).forEach(sourceScenario => {
    sourceScenario.metrics = Object.fromEntries(Object.entries(sourceScenario.metrics || {}).map(([metricId, metric]) => [
      metricId,
      migrateMetricScenarioYears(metric, oldYears, normalizedYears),
    ]));
  });
  state.model.years = normalizedYears;
  syncActiveYears();
  state.topDownSheetExpandedKeys.clear();
  clearCalculationCache();
  return true;
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function scenarioCollection() {
  if (!state.model.scenarios || typeof state.model.scenarios !== "object") {
    const fallback = state.model.scenario && typeof state.model.scenario === "object"
      ? state.model.scenario
      : { id: "scenario-1", name: "Scenario 1", metrics: {} };
    const scenarioId = fallback.id || "scenario-1";
    state.model.scenarios = {
      [scenarioId]: {
        ...fallback,
        id: scenarioId,
        name: fallback.name || "Scenario 1",
        metrics: fallback.metrics || {},
      },
    };
    state.model.activeScenarioId = scenarioId;
    delete state.model.scenario;
  }
  if (!Object.keys(state.model.scenarios).length) {
    state.model.scenarios["scenario-1"] = { id: "scenario-1", name: "Scenario 1", metrics: {} };
    state.model.activeScenarioId = "scenario-1";
  }
  return state.model.scenarios;
}

function activeScenarioId() {
  const scenarios = scenarioCollection();
  if (!state.model.activeScenarioId || !scenarios[state.model.activeScenarioId]) {
    state.model.activeScenarioId = Object.keys(scenarios)[0];
  }
  return state.model.activeScenarioId;
}

function scenario() {
  return scenarioCollection()[activeScenarioId()];
}

function scenarioRoles() {
  if (!state.model.scenarioRoles || typeof state.model.scenarioRoles !== "object") {
    state.model.scenarioRoles = {};
  }
  state.model.scenarioRoles.working = activeScenarioId();
  if (!Object.prototype.hasOwnProperty.call(state.model.scenarioRoles, "bau")) {
    state.model.scenarioRoles.bau = null;
  }
  if (!Object.prototype.hasOwnProperty.call(state.model.scenarioRoles, "target")) {
    state.model.scenarioRoles.target = null;
  }
  return state.model.scenarioRoles;
}

function referenceScenarios() {
  if (!state.model.referenceScenarios || typeof state.model.referenceScenarios !== "object") {
    state.model.referenceScenarios = {};
  }
  if (!Object.prototype.hasOwnProperty.call(state.model.referenceScenarios, "bau")) {
    state.model.referenceScenarios.bau = null;
  }
  if (!Object.prototype.hasOwnProperty.call(state.model.referenceScenarios, "target")) {
    state.model.referenceScenarios.target = null;
  }
  return state.model.referenceScenarios;
}

function assumptions() {
  if (!Array.isArray(state.model.assumptions)) state.model.assumptions = [];
  return state.model.assumptions;
}

function filterValueList(value) {
  if (Array.isArray(value)) return value.filter(item => item !== undefined && item !== null && item !== "").map(String);
  if (value === undefined || value === null || value === "") return [];
  return [String(value)];
}

function filterKey(filters = {}) {
  return Object.entries(filters)
    .filter(([_dimensionId, value]) => filterValueList(value).length)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([dimensionId, value]) => {
      const values = filterValueList(value).sort();
      return `${encodeURIComponent(dimensionId)}=${values.map(item => encodeURIComponent(item)).join("~")}`;
    })
    .join("|");
}

function coordinateMatchesFilters(key, filters = {}) {
  const coordinate = parseCoordinateKey(key);
  return Object.entries(filters).every(([dimensionId, value]) => {
    const values = filterValueList(value);
    return values.length ? values.includes(String(coordinate[dimensionId])) : true;
  });
}

function selectedDefinition() {
  return state.selectedMetricId ? metricDefinitions()[state.selectedMetricId] : null;
}

function selectedDimensionContextForMetric(metricId = state.selectedMetricId) {
  const context = state.selectedDimensionContext;
  if (!context || context.metricId !== metricId) return null;
  const rawCoordinate = context.coordinate || (
    context.dimensionId && context.memberId
      ? { [context.dimensionId]: context.memberId }
      : {}
  );
  const coordinate = {};
  metricDimensionIds(metricId).forEach(dimensionId => {
    const memberIds = filterValueList(rawCoordinate[dimensionId])
      .filter(memberId => dimensionMembers(dimensionId).some(item => item.id === memberId));
    if (!memberIds.length) return;
    coordinate[dimensionId] = memberIds.length === 1 ? memberIds[0] : memberIds;
  });
  if (!Object.keys(coordinate).length) return null;
  const key = filterKey(coordinate);
  const entries = Object.entries(coordinate).map(([dimensionId, value]) => {
    const memberIds = filterValueList(value);
    const members = memberIds.map(memberId => dimensionMembers(dimensionId).find(item => item.id === memberId)).filter(Boolean);
    return {
      dimensionId,
      memberId: memberIds.length === 1 ? memberIds[0] : null,
      memberIds,
      member: members.length === 1 ? members[0] : null,
      members,
    };
  });
  return {
    metricId,
    coordinate,
    coordinateKey: key,
    entries,
    label: filterLabelForMetric(metricId, coordinate),
    ...(entries.length === 1 ? entries[0] : {}),
  };
}

function selectedNodeKey() {
  const context = selectedDimensionContextForMetric();
  return context?.entries?.length === 1 && context.dimensionId === primaryDimensionId(context.metricId)
    ? memberNodeKey(context.metricId, context.dimensionId, context.memberId)
    : state.selectedMetricId;
}

function selectMetric(metricId) {
  state.selectedMetricId = metricId;
  state.selectedDimensionContext = null;
}

function selectNodeKey(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  state.selectedMetricId = parsed.metricId;
  state.selectedDimensionContext = parsed.memberId
    ? { metricId: parsed.metricId, coordinate: { [parsed.dimensionId]: parsed.memberId } }
    : null;
}

function metricScenario(metricId) {
  return scenario().metrics[metricId] || {};
}

function metricDimensionIds(metricId) {
  const dimensions = metricDefinitions()[metricId]?.dimensions;
  return Array.isArray(dimensions) ? dimensions : [];
}

function metricDimensionLevelLabel(metricId, dimensionId) {
  const index = metricDimensionIds(metricId).indexOf(dimensionId);
  const dimensionLabel = dimensionDefinitions()[dimensionId]?.label || dimensionId;
  return index >= 0 ? `Level ${index + 1}: ${dimensionLabel}` : dimensionLabel;
}

function primaryDimensionId(metricId) {
  return metricDimensionIds(metricId)[0] || null;
}

function dimensionMembers(dimensionId) {
  const members = dimensionDefinitions()[dimensionId]?.members;
  if (!Array.isArray(members)) return [];
  return members.map(member => {
    if (typeof member === "string") return { id: member, label: member };
    return { id: member.id, label: member.label || member.id };
  }).filter(member => member.id);
}

function normalizeDimensionMembersFromText(text) {
  const usedIds = new Set();
  return String(text || "")
    .split(/[\n,]+/)
    .map(value => value.trim())
    .filter(Boolean)
    .map(label => {
      let id = metricSlug(label);
      let suffix = 2;
      while (usedIds.has(id)) {
        id = `${metricSlug(label)}-${suffix}`;
        suffix += 1;
      }
      usedIds.add(id);
      return { id, label };
    });
}

function dimensionMembersText(dimensionId) {
  return dimensionMembers(dimensionId).map(member => member.label).join(", ");
}

function uniqueDimensionId(label, existingId = null) {
  const base = metricSlug(label);
  let id = base;
  let suffix = 2;
  while (dimensionDefinitions()[id] && id !== existingId) {
    id = `${base}-${suffix}`;
    suffix += 1;
  }
  return id;
}

function uniqueMemberId(label, members) {
  const existingIds = new Set(members.map(member => member.id));
  const base = metricSlug(label) || "member";
  let id = base;
  let suffix = 2;
  while (existingIds.has(id)) {
    id = `${base}-${suffix}`;
    suffix += 1;
  }
  return id;
}

function normalizeLabel(value) {
  return String(value || "").trim().toLowerCase();
}

function mergeDimensionMembers(existingMembers, labels) {
  const usedExistingIds = new Set();
  const existingByLabel = new Map(existingMembers.map(member => [normalizeLabel(member.label), member]));
  return labels.reduce((nextMembers, label, index) => {
    const labelMatch = existingByLabel.get(normalizeLabel(label));
    const sameIndexMatch = labels.length === existingMembers.length ? existingMembers[index] : null;
    const existing = labelMatch && !usedExistingIds.has(labelMatch.id)
      ? labelMatch
      : sameIndexMatch && !usedExistingIds.has(sameIndexMatch.id)
        ? sameIndexMatch
        : null;
    const member = existing
      ? { ...existing, label }
      : { id: uniqueMemberId(label, [...existingMembers, ...nextMembers]), label };
    if (member.id === UNASSIGNED_MEMBER_ID) member.system = true;
    usedExistingIds.add(member.id);
    nextMembers.push(member);
    return nextMembers;
  }, []);
}

function ensureUnassignedMember(members) {
  if (members.some(member => member.id === UNASSIGNED_MEMBER_ID)) return members;
  return [
    ...members,
    { id: UNASSIGNED_MEMBER_ID, label: "Unassigned", system: true },
  ];
}

function memberNodeKey(metricId, dimensionId, memberId) {
  return `${metricId}::member::${dimensionId}::${memberId}`;
}

function parseNodeKey(nodeKey) {
  const parts = String(nodeKey || "").split("::member::");
  if (parts.length !== 2) return { metricId: nodeKey, dimensionId: null, memberId: null };
  const [metricId, dimensionPart] = parts;
  const [dimensionId, memberId] = dimensionPart.split("::");
  return { metricId, dimensionId, memberId };
}

function isMemberNodeKey(nodeKey) {
  return Boolean(parseNodeKey(nodeKey).memberId);
}

function dimensionMemberNodeKeys(metricId) {
  const definition = metricDefinitions()[metricId];
  if (!definition || metricIsRate(definition)) return [];
  const dimensionId = primaryDimensionId(metricId);
  return dimensionMembers(dimensionId).map(member => memberNodeKey(metricId, dimensionId, member.id));
}

function metricUsesDimension(metricId, dimensionId) {
  return Boolean(dimensionId && metricDimensionIds(metricId).includes(dimensionId));
}

function metricCoordinateKeys(metricId) {
  const dimensions = metricDimensionIds(metricId);
  if (!dimensions.length) return [TOTAL_COORDINATE_KEY];

  let coordinates = [{}];
  dimensions.forEach(dimensionId => {
    const members = dimensionMembers(dimensionId);
    const memberIds = members.length ? members.map(member => member.id) : [UNASSIGNED_MEMBER_ID];
    coordinates = coordinates.flatMap(coordinate => memberIds.map(memberId => ({
      ...coordinate,
      [dimensionId]: memberId,
    })));
  });
  return coordinates.map(coordinateKey);
}

function normalizeRawCoordinateKey(definition, rawKey) {
  if (!rawKey) return TOTAL_COORDINATE_KEY;
  if (String(rawKey).includes("=")) return coordinateKey(parseCoordinateKey(rawKey));
  const [dimensionId] = Array.isArray(definition.dimensions) ? definition.dimensions : [];
  return dimensionId ? memberCoordinateKey(dimensionId, rawKey) : TOTAL_COORDINATE_KEY;
}

function normalizeSeriesMap(definition, rawTopDown) {
  const map = {};
  if (Array.isArray(rawTopDown)) {
    map[TOTAL_COORDINATE_KEY] = [...rawTopDown];
  } else if (rawTopDown && typeof rawTopDown === "object") {
    Object.entries(rawTopDown).forEach(([rawKey, series]) => {
      if (!Array.isArray(series)) return;
      const key = normalizeRawCoordinateKey(definition, rawKey);
      map[key] = map[key] ? sumSeries(map[key], series) : [...series];
    });
  }

  if (!Object.keys(map).length) {
    map[TOTAL_COORDINATE_KEY] = defaultSeries();
  }
  return map;
}

function normalizeControlPointMap(definition, rawMetric = {}) {
  const map = {};
  if (rawMetric.controlPoints && typeof rawMetric.controlPoints === "object") {
    Object.entries(rawMetric.controlPoints).forEach(([rawKey, points]) => {
      if (!Array.isArray(points)) return;
      map[normalizeRawCoordinateKey(definition, rawKey)] = normalizeNapkinPoints(points);
    });
  }
  if (Array.isArray(rawMetric.manualControlPoints)) {
    map[TOTAL_COORDINATE_KEY] = normalizeNapkinPoints(rawMetric.manualControlPoints);
  }
  if (rawMetric.memberManualControlPoints && typeof rawMetric.memberManualControlPoints === "object") {
    Object.entries(rawMetric.memberManualControlPoints).forEach(([memberId, points]) => {
      if (!Array.isArray(points)) return;
      map[normalizeRawCoordinateKey(definition, memberId)] = normalizeNapkinPoints(points);
    });
  }
  return map;
}

function metricSeriesMap(metricId) {
  const topDown = metricScenario(metricId).topDown;
  return topDown && typeof topDown === "object" && !Array.isArray(topDown)
    ? topDown
    : normalizeSeriesMap(metricDefinitions()[metricId] || {}, topDown);
}

function aggregateSeriesMap(metricId, filters = {}) {
  const definition = metricDefinitions()[metricId];
  const entries = Object.entries(metricSeriesMap(metricId))
    .filter(([key]) => coordinateMatchesFilters(key, filters));
  if (!entries.length) return defaultSeries();
  return aggregateMemberSeries(metricId, Object.fromEntries(entries));
}

function aggregateMemberSeries(metricId, seriesByMember) {
  const definition = metricDefinitions()[metricId];
  const values = Object.values(seriesByMember || {});
  if (!values.length) return defaultSeries();
  const isRate = metricIsRate(definition);
  return YEARS.map((_year, index) => {
    const numbers = values.map(series => Number(series?.[index] || 0));
    const total = numbers.reduce((sum, value) => sum + value, 0);
    return isRate ? total / numbers.length : total;
  });
}

function metricIsRate(definition) {
  return definition?.unit === "percent" || definition?.time?.flowType === "rate";
}

function metricIsLagged(definition) {
  return definition?.bottomUp?.type === "laggedMetric";
}

function laggedTopDownDefinition() {
  return { type: "derivedLaggedSeries", editable: false };
}

function metricTopDownMemberSeries(metricId, memberId) {
  return cachedCalculation(`topDownMember:${metricId}:${memberId}`, () => {
    const dimensionId = primaryDimensionId(metricId);
    if (!dimensionId) return defaultSeries();
    return aggregateSeriesMap(metricId, { [dimensionId]: memberId });
  });
}

function filtersSelectSingleMetricCoordinate(metricId, filters = {}) {
  const dimensionIds = metricDimensionIds(metricId);
  return dimensionIds.length > 0 && dimensionIds.every(dimensionId => filterValueList(filters[dimensionId]).length === 1);
}

function metricTopDownFilteredSeries(metricId, filters = {}) {
  return cachedCalculation(
    `topDownFiltered:${metricId}:${filterKey(filters)}`,
    () => {
      const definition = metricDefinitions()[metricId];
      if (
        definition?.bottomUp?.type === "weightedAverage"
        && !weightedAverageConfig(definition).valueMetricId
        && metricDimensionIds(metricId).length
        && !filtersSelectSingleMetricCoordinate(metricId, filters)
      ) {
        return weightedAverageSeriesForFilters(definition, filters) || aggregateSeriesMap(metricId, filters);
      }
      return aggregateSeriesMap(metricId, filters);
    }
  );
}

function metricMemberControlPoints(metricId, memberId) {
  const dimensionId = primaryDimensionId(metricId);
  const coordinate = memberCoordinateKey(dimensionId, memberId);
  const stored = metricScenario(metricId).controlPoints?.[coordinate];
  if (Array.isArray(stored) && stored.length) return normalizeNapkinPoints(stored);
  return YEARS.map((year, index) => [year, Number(metricTopDownMemberSeries(metricId, memberId)[index] || 0)]);
}

function setMetricTopDownCoordinateSeries(metricId, coordinate, series) {
  const metric = ensureMetricScenario(metricId);
  metric.topDown = normalizeSeriesMap(metricDefinitions()[metricId] || {}, metric.topDown);
  metric.topDown[coordinate] = [...series];
  delete metric.controlPoints?.[TOTAL_COORDINATE_KEY];
  delete metric.dimensionShareControlPoints;
  clearCalculationCache();
}

function setMetricTopDownMemberSeries(metricId, memberId, series) {
  const dimensionId = primaryDimensionId(metricId);
  setMetricTopDownCoordinateSeries(metricId, memberCoordinateKey(dimensionId, memberId), series);
}

function setMetricTopDownCoordinateControlPoints(metricId, coordinate, series) {
  const metric = ensureMetricScenario(metricId);
  if (!metric.controlPoints) metric.controlPoints = {};
  metric.controlPoints[coordinate] = YEARS.map((year, index) => [year, Number(series[index] || 0)]);
  setMetricTopDownCoordinateSeries(metricId, coordinate, series);
}

function setMetricTopDownFilteredSeries(metricId, filters = {}, series) {
  const metric = ensureMetricScenario(metricId);
  const filterKeyValue = filterKey(filters);
  const coordinateKeys = metricCoordinateKeys(metricId)
    .filter(key => coordinateMatchesFilters(key, filters));

  if (!filterKeyValue || !Object.keys(filters).length) {
    if (!metric.controlPoints) metric.controlPoints = {};
    metric.controlPoints[TOTAL_COORDINATE_KEY] = YEARS.map((year, index) => [year, Number(series[index] || 0)]);
    setMetricTopDownSeries(metricId, series);
    return;
  }

  if (coordinateKeys.length === 1 && coordinateKeys[0] === filterKeyValue) {
    setMetricTopDownCoordinateControlPoints(metricId, filterKeyValue, series);
    return;
  }

  const existingByCoordinate = Object.fromEntries(coordinateKeys.map(key => [
    key,
    metricTopDownCoordinateSeries(metricId, key),
  ]));
  const sharesByYear = YEARS.map((_year, index) => {
    const directShare = sharesForKeys(coordinateKeys, existingByCoordinate, index);
    if (directShare) return directShare;
    const nearestIndex = nearestIndexWithCoordinateMix(coordinateKeys, existingByCoordinate, index);
    if (nearestIndex !== null) return sharesForKeys(coordinateKeys, existingByCoordinate, nearestIndex);
    const evenShare = coordinateKeys.length ? 1 / coordinateKeys.length : 1;
    return Object.fromEntries(coordinateKeys.map(key => [key, evenShare]));
  });

  metric.topDown = normalizeSeriesMap(metricDefinitions()[metricId] || {}, metric.topDown);
  const definition = metricDefinitions()[metricId];
  coordinateKeys.forEach(key => {
    metric.topDown[key] = metricIsRate(definition)
      ? [...series]
      : YEARS.map((_year, index) => Number(series[index] || 0) * Number(sharesByYear[index]?.[key] || 0));
  });
  if (!metric.controlPoints) metric.controlPoints = {};
  metric.controlPoints[filterKeyValue] = YEARS.map((year, index) => [year, Number(series[index] || 0)]);
  delete metric.dimensionShareControlPoints;
  clearCalculationCache();
}

function preserveTotalControlPoints(metric) {
  const totalControlPoints = metric.controlPoints?.[TOTAL_COORDINATE_KEY];
  if (Array.isArray(totalControlPoints)) {
    metric.controlPoints = { [TOTAL_COORDINATE_KEY]: normalizeNapkinPoints(totalControlPoints) };
  } else {
    delete metric.controlPoints;
  }
}

function metricTopDownSeries(metricId) {
  return cachedCalculation(`topDown:${metricId}`, () => {
    const definition = metricDefinitions()[metricId];
    if (
      definition?.bottomUp?.type === "weightedAverage"
      && !weightedAverageConfig(definition).valueMetricId
      && metricDimensionIds(metricId).length
    ) {
      return weightedAverageSeriesForFilters(definition) || aggregateSeriesMap(metricId);
    }
    return aggregateSeriesMap(metricId);
  });
}

function setMetricTopDownSeries(metricId, series) {
  if (allocateDimensionTopDownSeries(metricId, series)) return;
  const metric = ensureMetricScenario(metricId);
  metric.topDown = { [TOTAL_COORDINATE_KEY]: [...series] };
  clearCalculationCache();
}

function allocateDimensionTopDownSeries(metricId, series) {
  const definition = metricDefinitions()[metricId];
  if (!metricDimensionIds(metricId).length) return false;

  const coordinateKeys = metricCoordinateKeys(metricId);
  if (!coordinateKeys.length) return false;
  const existingSeriesByCoordinate = Object.fromEntries(coordinateKeys.map(key => [
    key,
    metricTopDownCoordinateSeries(metricId, key),
  ]));
  const coordinateSharesByYear = YEARS.map((_year, index) => {
    const directShare = sharesForKeys(coordinateKeys, existingSeriesByCoordinate, index);
    if (directShare) return directShare;

    const nearestIndex = nearestIndexWithCoordinateMix(coordinateKeys, existingSeriesByCoordinate, index);
    if (nearestIndex !== null) return sharesForKeys(coordinateKeys, existingSeriesByCoordinate, nearestIndex);

    const evenShare = coordinateKeys.length ? 1 / coordinateKeys.length : 1;
    return Object.fromEntries(coordinateKeys.map(key => [key, evenShare]));
  });

  const allocatedTopDown = Object.fromEntries(coordinateKeys.map(key => [
    key,
    metricIsRate(definition)
      ? [...series]
      : YEARS.map((_year, index) => Number(series[index] || 0) * Number(coordinateSharesByYear[index]?.[key] || 0)),
  ]));

  const metric = ensureMetricScenario(metricId);
  metric.topDown = allocatedTopDown;
  preserveTotalControlPoints(metric);
  clearCalculationCache();
  return true;
}

function metricTopDownCoordinateSeries(metricId, coordinate) {
  const map = metricSeriesMap(metricId);
  return Array.isArray(map[coordinate]) ? map[coordinate] : defaultSeries();
}

function coordinateLabel(key) {
  if (!key) return "Total";
  const coordinate = parseCoordinateKey(key);
  const orderedDimensionIds = Object.keys(dimensionDefinitions());
  return Object.entries(coordinate)
    .sort(([leftDimensionId], [rightDimensionId]) => {
      const leftIndex = orderedDimensionIds.indexOf(leftDimensionId);
      const rightIndex = orderedDimensionIds.indexOf(rightDimensionId);
      return (leftIndex < 0 ? 999 : leftIndex) - (rightIndex < 0 ? 999 : rightIndex);
    })
    .map(([dimensionId, memberId]) => {
      const dimension = dimensionDefinitions()[dimensionId];
      const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
      return `${dimension?.label || dimensionId}: ${member?.label || memberId}`;
    })
    .join(" | ");
}

function coordinateLabelForMetric(metricId, key) {
  if (!key) return "Total";
  const coordinate = parseCoordinateKey(key);
  return filterLabelForMetric(metricId, coordinate);
}

function filterLabelForMetric(metricId, coordinate = {}) {
  if (!Object.keys(coordinate || {}).length) return "Total";
  const orderedDimensionIds = [
    ...metricDimensionIds(metricId),
    ...Object.keys(coordinate).filter(dimensionId => !metricDimensionIds(metricId).includes(dimensionId)),
  ];
  return orderedDimensionIds
    .filter(dimensionId => coordinate[dimensionId])
    .map(dimensionId => {
      const memberIds = filterValueList(coordinate[dimensionId]);
      const memberLabels = memberIds.map(memberId => {
        const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
        return member?.label || memberId;
      });
      const levelIndex = metricDimensionIds(metricId).indexOf(dimensionId);
      const levelPrefix = levelIndex >= 0 ? `L${levelIndex + 1} ` : "";
      const label = memberLabels.length > 2
        ? `${memberLabels.slice(0, 2).join(", ")} +${memberLabels.length - 2}`
        : memberLabels.join(", ");
      return `${levelPrefix}${dimensionDefinitions()[dimensionId]?.label || dimensionId}: ${label}`;
    })
    .join(" | ");
}

function sharesForKeys(keys, seriesByKey, index) {
  const values = keys.map(key => Number(seriesByKey[key]?.[index] || 0));
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!Number.isFinite(total) || total === 0) return null;
  return Object.fromEntries(keys.map((key, keyIndex) => [key, values[keyIndex] / total]));
}

function nearestIndexWithCoordinateMix(keys, seriesByKey, targetIndex) {
  let bestIndex = null;
  let bestDistance = Infinity;
  YEARS.forEach((_year, index) => {
    if (!sharesForKeys(keys, seriesByKey, index)) return;
    const distance = Math.abs(index - targetIndex);
    if (distance < bestDistance || (distance === bestDistance && index < targetIndex)) {
      bestIndex = index;
      bestDistance = distance;
    }
  });
  return bestIndex;
}

function sharesForIndex(members, seriesByMember, index) {
  const values = members.map(member => Number(seriesByMember[member.id]?.[index] || 0));
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!Number.isFinite(total) || total === 0) return null;
  return Object.fromEntries(members.map((member, memberIndex) => [member.id, values[memberIndex] / total]));
}

function nearestIndexWithMemberMix(members, seriesByMember, targetIndex) {
  let bestIndex = null;
  let bestDistance = Infinity;
  YEARS.forEach((_year, index) => {
    if (!sharesForIndex(members, seriesByMember, index)) return;
    const distance = Math.abs(index - targetIndex);
    if (distance < bestDistance || (distance === bestDistance && index < targetIndex)) {
      bestIndex = index;
      bestDistance = distance;
    }
  });
  return bestIndex;
}

function dimensionTopDownSeriesByMember(metricId, members) {
  return Object.fromEntries(members.map(member => [
    member.id,
    metricTopDownMemberSeries(metricId, member.id),
  ]));
}

function dimensionSharesByYear(metricId, members) {
  const seriesByMember = dimensionTopDownSeriesByMember(metricId, members);
  return YEARS.map((_year, index) => {
    const directShare = sharesForIndex(members, seriesByMember, index);
    if (directShare) return directShare;

    const nearestIndex = nearestIndexWithMemberMix(members, seriesByMember, index);
    if (nearestIndex !== null) return sharesForIndex(members, seriesByMember, nearestIndex);

    const evenShare = 1 / members.length;
    return Object.fromEntries(members.map(member => [member.id, evenShare]));
  });
}

function dimensionChartMembers(members) {
  return [...members].reverse();
}

function dimensionMemberColor(members, memberId) {
  const index = Math.max(0, members.findIndex(member => member.id === memberId));
  return dimensionPalette[index % dimensionPalette.length];
}

function dimensionMixContext(definition) {
  const metricId = definition?.id;
  const dimensionIds = metricDimensionIds(metricId);
  if (!metricId || !dimensionIds.length) return null;
  const selectedContext = selectedDimensionContextForMetric(metricId);
  const baseFilters = selectedContext?.coordinate || {};
  const dimensionId = dimensionIds.find(id => !filterValueList(baseFilters[id]).length);
  if (!dimensionId) return null;
  const members = dimensionMembers(dimensionId);
  return {
    metricId,
    baseFilters,
    baseKey: filterKey(baseFilters),
    dimensionId,
    dimensionLabel: metricDimensionLevelLabel(metricId, dimensionId),
    members,
  };
}

function dimensionMixControlKey(mixContext) {
  return `${mixContext.baseKey || TOTAL_COORDINATE_KEY}::${mixContext.dimensionId}`;
}

function dimensionGroupingStore(metricId) {
  const metric = ensureMetricScenario(metricId);
  if (!metric.dimensionGroupsByContext || typeof metric.dimensionGroupsByContext !== "object") {
    metric.dimensionGroupsByContext = {};
  }
  return metric.dimensionGroupsByContext;
}

function dimensionGroupingForContext(mixContext) {
  const store = dimensionGroupingStore(mixContext.metricId);
  const key = dimensionMixControlKey(mixContext);
  if (!store[key] || typeof store[key] !== "object") {
    store[key] = { red: [], blue: [] };
  }
  store[key].red = Array.isArray(store[key].red) ? store[key].red.filter(memberId => mixContext.members.some(member => member.id === memberId)) : [];
  store[key].blue = Array.isArray(store[key].blue) ? store[key].blue.filter(memberId => mixContext.members.some(member => member.id === memberId)) : [];
  store[key].blue = store[key].blue.filter(memberId => !store[key].red.includes(memberId));
  return store[key];
}

function dimensionGroupColor(groupId) {
  return dimensionGroupDefinitions.find(group => group.id === groupId)?.color || "#98a2b3";
}

function dimensionGroupsForContext(mixContext, { includeEmpty = false } = {}) {
  const grouping = dimensionGroupingForContext(mixContext);
  const assigned = new Set([...grouping.red, ...grouping.blue]);
  const groups = dimensionGroupDefinitions.map(group => {
    const memberIds = group.id === "other"
      ? mixContext.members.map(member => member.id).filter(memberId => !assigned.has(memberId))
      : grouping[group.id] || [];
    return {
      ...group,
      memberIds,
      members: memberIds
        .map(memberId => mixContext.members.find(member => member.id === memberId))
        .filter(Boolean),
    };
  });
  return includeEmpty ? groups : groups.filter(group => group.members.length > 0);
}

function dimensionGroupSeries(mixContext, group) {
  return group.members.reduce((total, member) => {
    const memberSeries = metricTopDownFilteredSeries(
      mixContext.metricId,
      { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id }
    );
    return sumSeries(total, memberSeries);
  }, defaultSeries());
}

function dimensionGroupBottomUpSeries(mixContext, group) {
  const series = group.members
    .map(member => metricBottomUpFilteredSeries(
      mixContext.metricId,
      { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id }
    ))
    .filter(Boolean);
  if (!series.length) return null;
  return series.reduce((total, item) => sumSeries(total, item), defaultSeries());
}

function dimensionGroupSharesByYear(mixContext, groups) {
  const seriesByGroup = Object.fromEntries(groups.map(group => [group.id, dimensionGroupSeries(mixContext, group)]));
  return YEARS.map((_year, index) => {
    const total = groups.reduce((sum, group) => sum + Number(seriesByGroup[group.id]?.[index] || 0), 0);
    if (total) {
      return Object.fromEntries(groups.map(group => [group.id, Number(seriesByGroup[group.id]?.[index] || 0) / total]));
    }
    const evenShare = groups.length ? 1 / groups.length : 1;
    return Object.fromEntries(groups.map(group => [group.id, evenShare]));
  });
}

function dimensionGroupLineObject(group, groupIndex, points) {
  if (!group) return null;
  return {
    name: group.label,
    memberId: group.id,
    groupId: group.id,
    boundaryIndex: groupIndex,
    color: group.color,
    data: normalizeNapkinPoints(points),
    editable: true,
  };
}

function dimensionGroupShareBoundaryLines(mixContext) {
  const groups = dimensionGroupsForContext(mixContext);
  if (groups.length <= 1) return [];
  const chartGroups = dimensionChartMembers(groups);
  const sharesByYear = dimensionGroupSharesByYear(mixContext, groups);
  let cumulative = YEARS.map(() => 0);
  return chartGroups.slice(0, -1).map((group, groupIndex) => {
    cumulative = cumulative.map((value, yearIndex) => value + Number(sharesByYear[yearIndex]?.[group.id] || 0));
    const points = YEARS.map((year, yearIndex) => [year, cumulative[yearIndex] * 100]);
    return dimensionGroupLineObject(group, groupIndex, points);
  }).filter(Boolean).reverse();
}

function dimensionSharesFromGroupLines(lines, groups) {
  if (groups.length <= 1) {
    return YEARS.map(() => Object.fromEntries(groups.map(group => [group.id, 1])));
  }
  const chartGroups = dimensionChartMembers(groups);
  const orderedLines = [...lines].sort((left, right) => Number(left.boundaryIndex || 0) - Number(right.boundaryIndex || 0));
  return YEARS.map(year => {
    const boundaries = orderedLines
      .slice(0, groups.length - 1)
      .map(line => Math.max(0, Math.min(100, interpolateLineValue(line.data, year))))
      .sort((left, right) => left - right);
    const shares = {};
    let previousBoundary = 0;
    chartGroups.forEach((group, index) => {
      const boundary = index < chartGroups.length - 1 ? boundaries[index] : 100;
      shares[group.id] = Math.max(0, (boundary - previousBoundary) / 100);
      previousBoundary = boundary;
    });
    return shares;
  });
}

function applyDimensionGroupShares(mixContext, lines) {
  const groups = dimensionGroupsForContext(mixContext);
  if (!groups.length) return;
  const totals = metricTopDownFilteredSeries(mixContext.metricId, mixContext.baseFilters);
  const sharesByYear = dimensionSharesFromGroupLines(lines, groups);
  const lockedGroups = groups.filter(group => workspaceIsLocked(workspaceDimensionGroupLockKey(mixContext, group.id)));
  const editableGroups = groups.filter(group => !workspaceIsLocked(workspaceDimensionGroupLockKey(mixContext, group.id)));
  if (!editableGroups.length) {
    setSourceStatus("All dimension groups are locked. Unlock at least one group before editing the mix.", "error");
    return;
  }
  const currentSeriesByGroup = Object.fromEntries(groups.map(group => [group.id, dimensionGroupSeries(mixContext, group)]));
  editableGroups.forEach(group => {
    const targetSeries = YEARS.map((_year, index) => {
      const lockedTotal = lockedGroups.reduce((sum, lockedGroup) => {
        return sum + Number(currentSeriesByGroup[lockedGroup.id]?.[index] || 0);
      }, 0);
      const remainingTotal = Number(totals[index] || 0) - lockedTotal;
      const editableShareTotal = editableGroups.reduce((sum, editableGroup) => {
        return sum + Number(sharesByYear[index]?.[editableGroup.id] || 0);
      }, 0);
      const share = editableShareTotal
        ? Number(sharesByYear[index]?.[group.id] || 0) / editableShareTotal
        : 1 / editableGroups.length;
      return remainingTotal * share;
    });
    setDimensionGroupSeries(mixContext, group, targetSeries);
  });
  clearCalculationCache();
}

function setDimensionGroupSeries(mixContext, group, targetSeries) {
  if (!group?.members?.length) return;
  const definition = metricDefinitions()[mixContext.metricId];
  const currentSeriesByMember = Object.fromEntries(group.members.map(member => [
    member.id,
    metricTopDownFilteredSeries(mixContext.metricId, { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id }),
  ]));
  if (definition?.bottomUp?.type === "weightedAverage" && !weightedAverageConfig(definition).valueMetricId) {
    const currentGroupSeries = dimensionGroupSeries(mixContext, group);
    group.members.forEach(member => {
      const memberSeries = YEARS.map((_year, index) => {
        const currentMemberValue = Number(currentSeriesByMember[member.id]?.[index] || 0);
        const currentGroupValue = Number(currentGroupSeries[index] || 0);
        const targetValue = Number(targetSeries[index] || 0);
        return currentGroupValue ? currentMemberValue * (targetValue / currentGroupValue) : targetValue;
      });
      setMetricTopDownFilteredSeries(mixContext.metricId, { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id }, memberSeries);
    });
    return;
  }
  const sharesByYear = YEARS.map((_year, index) => {
    const total = group.members.reduce((sum, member) => sum + Number(currentSeriesByMember[member.id]?.[index] || 0), 0);
    if (total) {
      return Object.fromEntries(group.members.map(member => [member.id, Number(currentSeriesByMember[member.id]?.[index] || 0) / total]));
    }
    const evenShare = 1 / group.members.length;
    return Object.fromEntries(group.members.map(member => [member.id, evenShare]));
  });
  group.members.forEach(member => {
    const memberSeries = YEARS.map((_year, index) => Number(targetSeries[index] || 0) * Number(sharesByYear[index]?.[member.id] || 0));
    setMetricTopDownFilteredSeries(mixContext.metricId, { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id }, memberSeries);
  });
}

function setDimensionMemberGroup(mixContext, memberId, groupId) {
  const grouping = dimensionGroupingForContext(mixContext);
  grouping.red = (grouping.red || []).filter(id => id !== memberId);
  grouping.blue = (grouping.blue || []).filter(id => id !== memberId);
  if (groupId === "red" || groupId === "blue") {
    grouping[groupId].push(memberId);
  }
  clearCalculationCache();
}

function seriesByDimensionMember(metricId, dimensionId, members, baseFilters = {}) {
  return Object.fromEntries(members.map(member => [
    member.id,
    metricTopDownFilteredSeries(metricId, { ...baseFilters, [dimensionId]: member.id }),
  ]));
}

function dimensionSharesByYearForContext(metricId, dimensionId, members, baseFilters = {}) {
  const seriesByMember = seriesByDimensionMember(metricId, dimensionId, members, baseFilters);
  return YEARS.map((_year, index) => {
    const directShare = sharesForIndex(members, seriesByMember, index);
    if (directShare) return directShare;

    const nearestIndex = nearestIndexWithMemberMix(members, seriesByMember, index);
    if (nearestIndex !== null) return sharesForIndex(members, seriesByMember, nearestIndex);

    const evenShare = 1 / members.length;
    return Object.fromEntries(members.map(member => [member.id, evenShare]));
  });
}

function dimensionShareBoundaryLines(metricId, members) {
  const metric = metricScenario(metricId);
  const chartMembers = dimensionChartMembers(members);
  if (Array.isArray(metric.dimensionShareControlPoints) && metric.dimensionShareControlPoints.length) {
    return metric.dimensionShareControlPoints
      .map((lineData, index) => dimensionShareLineObject(chartMembers[index], index, lineData, members))
      .filter(Boolean)
      .reverse();
  }

  const sharesByYear = dimensionSharesByYear(metricId, members);
  let cumulative = YEARS.map(() => 0);
  return chartMembers.slice(0, -1).map((member, memberIndex) => {
    cumulative = cumulative.map((value, yearIndex) => value + Number(sharesByYear[yearIndex]?.[member.id] || 0));
    const points = YEARS.map((year, yearIndex) => [year, cumulative[yearIndex] * 100]);
    return dimensionShareLineObject(member, memberIndex, points, members);
  }).reverse();
}

function dimensionShareLineObject(member, memberIndex, points, members) {
  if (!member) return null;
  return {
    name: member.label,
    memberId: member.id,
    boundaryIndex: memberIndex,
    color: dimensionMemberColor(members, member.id),
    data: normalizeNapkinPoints(points),
    editable: true,
  };
}

function dimensionShareBoundaryLinesForContext(mixContext) {
  const metric = metricScenario(mixContext.metricId);
  const chartMembers = dimensionChartMembers(mixContext.members);
  const controlKey = dimensionMixControlKey(mixContext);
  const storedLines = metric.dimensionShareControlPointsByContext?.[controlKey];
  if (Array.isArray(storedLines) && storedLines.length) {
    return storedLines
      .map((lineData, index) => dimensionShareLineObject(chartMembers[index], index, lineData, mixContext.members))
      .filter(Boolean)
      .reverse();
  }

  if (
    !mixContext.baseKey
    && mixContext.dimensionId === primaryDimensionId(mixContext.metricId)
    && metricDimensionIds(mixContext.metricId).length === 1
  ) {
    return dimensionShareBoundaryLines(mixContext.metricId, mixContext.members);
  }

  const sharesByYear = dimensionSharesByYearForContext(
    mixContext.metricId,
    mixContext.dimensionId,
    mixContext.members,
    mixContext.baseFilters
  );
  let cumulative = YEARS.map(() => 0);
  return chartMembers.slice(0, -1).map((member, memberIndex) => {
    cumulative = cumulative.map((value, yearIndex) => value + Number(sharesByYear[yearIndex]?.[member.id] || 0));
    const points = YEARS.map((year, yearIndex) => [year, cumulative[yearIndex] * 100]);
    return dimensionShareLineObject(member, memberIndex, points, mixContext.members);
  }).reverse();
}

function dimensionSharesFromBoundaryLines(lines, members) {
  if (members.length <= 1) {
    return YEARS.map(() => Object.fromEntries(members.map(member => [member.id, 1])));
  }

  const chartMembers = dimensionChartMembers(members);
  const orderedLines = [...lines].sort((left, right) => Number(left.boundaryIndex || 0) - Number(right.boundaryIndex || 0));
  return YEARS.map(year => {
    const boundaries = orderedLines
      .slice(0, members.length - 1)
      .map(line => Math.max(0, Math.min(100, interpolateLineValue(line.data, year))))
      .sort((left, right) => left - right);

    const shares = {};
    let previousBoundary = 0;
    chartMembers.forEach((member, index) => {
      const boundary = index < chartMembers.length - 1 ? boundaries[index] : 100;
      shares[member.id] = Math.max(0, (boundary - previousBoundary) / 100);
      previousBoundary = boundary;
    });
    return shares;
  });
}

function applyDimensionShares(metricId, members, lines) {
  const totals = metricTopDownSeries(metricId);
  const sharesByYear = dimensionSharesFromBoundaryLines(lines, members);
  const metric = ensureMetricScenario(metricId);
  metric.dimensionShareControlPoints = [...lines]
    .sort((left, right) => Number(left.boundaryIndex || 0) - Number(right.boundaryIndex || 0))
    .map(line => normalizeNapkinPoints(line.data));
  const dimensionId = primaryDimensionId(metricId);
  metric.topDown = Object.fromEntries(members.map(member => [
    memberCoordinateKey(dimensionId, member.id),
    YEARS.map((_year, index) => Number(totals[index] || 0) * Number(sharesByYear[index]?.[member.id] || 0)),
  ]));
  preserveTotalControlPoints(metric);
  clearCalculationCache();
}

function coordinateShareByYear(coordinateKeys, seriesByCoordinate) {
  return YEARS.map((_year, index) => {
    const directShare = sharesForKeys(coordinateKeys, seriesByCoordinate, index);
    if (directShare) return directShare;
    const nearestIndex = nearestIndexWithCoordinateMix(coordinateKeys, seriesByCoordinate, index);
    if (nearestIndex !== null) return sharesForKeys(coordinateKeys, seriesByCoordinate, nearestIndex);
    const evenShare = coordinateKeys.length ? 1 / coordinateKeys.length : 1;
    return Object.fromEntries(coordinateKeys.map(key => [key, evenShare]));
  });
}

function applyDimensionSharesForContext(mixContext, lines) {
  if (
    !mixContext.baseKey
    && mixContext.dimensionId === primaryDimensionId(mixContext.metricId)
    && metricDimensionIds(mixContext.metricId).length === 1
  ) {
    applyDimensionShares(mixContext.metricId, mixContext.members, lines);
    return;
  }

  const metric = ensureMetricScenario(mixContext.metricId);
  if (!metric.dimensionShareControlPointsByContext) metric.dimensionShareControlPointsByContext = {};
  metric.dimensionShareControlPointsByContext[dimensionMixControlKey(mixContext)] = [...lines]
    .sort((left, right) => Number(left.boundaryIndex || 0) - Number(right.boundaryIndex || 0))
    .map(line => normalizeNapkinPoints(line.data));

  const totals = metricTopDownFilteredSeries(mixContext.metricId, mixContext.baseFilters);
  const sharesByYear = dimensionSharesFromBoundaryLines(lines, mixContext.members);
  const definition = metricDefinitions()[mixContext.metricId];
  const allCoordinateKeys = metricCoordinateKeys(mixContext.metricId);
  metric.topDown = normalizeSeriesMap(definition || {}, metric.topDown);

  mixContext.members.forEach(member => {
    const memberFilters = { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id };
    const memberKeys = allCoordinateKeys.filter(key => coordinateMatchesFilters(key, memberFilters));
    const targetKeys = memberKeys.length ? memberKeys : [coordinateKey(memberFilters)];
    const currentByCoordinate = Object.fromEntries(targetKeys.map(key => [key, metricTopDownCoordinateSeries(mixContext.metricId, key)]));
    const deeperSharesByYear = coordinateShareByYear(targetKeys, currentByCoordinate);
    const memberTotals = YEARS.map((_year, index) => Number(totals[index] || 0) * Number(sharesByYear[index]?.[member.id] || 0));
    targetKeys.forEach(key => {
      metric.topDown[key] = metricIsRate(definition)
        ? [...memberTotals]
        : YEARS.map((_year, index) => Number(memberTotals[index] || 0) * Number(deeperSharesByYear[index]?.[key] || 0));
    });
  });
  clearCalculationCache();
}

function ensureMetricScenario(metricId) {
  if (!scenario().metrics[metricId]) {
    scenario().metrics[metricId] = { topDown: { [TOTAL_COORDINATE_KEY]: defaultSeries() } };
  }
  return scenario().metrics[metricId];
}

function ensureMetricScenarioFor(sourceScenario, metricId) {
  if (!sourceScenario.metrics) sourceScenario.metrics = {};
  if (!sourceScenario.metrics[metricId]) {
    sourceScenario.metrics[metricId] = { topDown: { [TOTAL_COORDINATE_KEY]: defaultSeries() } };
  }
  return sourceScenario.metrics[metricId];
}

function assignPrimaryDimension(metricId, dimensionId) {
  assignMetricDimensions(metricId, dimensionId ? [dimensionId] : []);
}

function assignMetricDimensions(metricId, dimensionIds) {
  const definition = metricDefinitions()[metricId];
  if (!definition) return;
  const totalSeries = metricTopDownSeries(metricId);
  const metric = ensureMetricScenario(metricId);
  const nextDimensionIds = [...new Set(dimensionIds)].filter(dimensionId => dimensionDefinitions()[dimensionId]);

  if (!nextDimensionIds.length) {
    definition.dimensions = [];
    metric.topDown = { [TOTAL_COORDINATE_KEY]: [...totalSeries] };
    metric.controlPoints = {
      [TOTAL_COORDINATE_KEY]: YEARS.map((year, index) => [year, Number(totalSeries[index] || 0)]),
    };
    delete metric.dimensionShareControlPoints;
    if (state.selectedMetricId === metricId) state.selectedDimensionContext = null;
    clearCalculationCache();
    return;
  }

  definition.dimensions = nextDimensionIds;
  allocateDimensionValuesFromTotal(metricId, totalSeries);
  if (state.selectedMetricId === metricId) state.selectedDimensionContext = null;
}

function allocateDimensionValuesFromTotal(metricId, totalSeries) {
  const coordinateKeys = metricCoordinateKeys(metricId);
  const metric = ensureMetricScenario(metricId);
  const existingSeriesByCoordinate = Object.fromEntries(coordinateKeys.map(key => [
    key,
    metricTopDownCoordinateSeries(metricId, key),
  ]));
  const definition = metricDefinitions()[metricId];
  const coordinateSharesByYear = YEARS.map((_year, index) => {
    const directShare = sharesForKeys(coordinateKeys, existingSeriesByCoordinate, index);
    if (directShare) return directShare;

    const nearestIndex = nearestIndexWithCoordinateMix(coordinateKeys, existingSeriesByCoordinate, index);
    if (nearestIndex !== null) return sharesForKeys(coordinateKeys, existingSeriesByCoordinate, nearestIndex);

    const evenShare = coordinateKeys.length ? 1 / coordinateKeys.length : 1;
    return Object.fromEntries(coordinateKeys.map(key => [key, evenShare]));
  });

  metric.topDown = Object.fromEntries(coordinateKeys.map(key => [
    key,
    metricIsRate(definition)
      ? [...totalSeries]
      : YEARS.map((_year, index) => Number(totalSeries[index] || 0) * Number(coordinateSharesByYear[index]?.[key] || 0)),
  ]));
  delete metric.controlPoints;
  delete metric.dimensionShareControlPoints;
  clearCalculationCache();
}

function reconcileMetricsForDimensionChange(dimensionId) {
  Object.values(metricDefinitions()).forEach(definition => {
    if (!metricUsesDimension(definition.id, dimensionId)) return;
    ensureMetricCoordinateCoverage(definition.id);
  });
}

function ensureMetricCoordinateCoverage(metricId) {
  const keys = metricCoordinateKeys(metricId);
  Object.values(scenarioCollection()).forEach(sourceScenario => {
    const metric = ensureMetricScenarioFor(sourceScenario, metricId);
    const existing = normalizeSeriesMap(metricDefinitions()[metricId] || {}, metric.topDown);
    metric.topDown = Object.fromEntries(keys.map(key => [
      key,
      Array.isArray(existing[key])
        ? existing[key]
        : defaultSeries(),
    ]));
  });
  clearCalculationCache();
}

function remapDeletedDimensionMembers(dimensionId, deletedMemberIds) {
  if (!deletedMemberIds.length) return;
  const deletedSet = new Set(deletedMemberIds);
  Object.values(metricDefinitions()).forEach(definition => {
    if (!metricUsesDimension(definition.id, dimensionId)) return;
    Object.values(scenarioCollection()).forEach(sourceScenario => {
      const metric = ensureMetricScenarioFor(sourceScenario, definition.id);
      metric.topDown = remapSeriesCoordinates(normalizeSeriesMap(definition, metric.topDown), dimensionId, deletedSet);
      if (metric.openingValue && typeof metric.openingValue === "object" && !Array.isArray(metric.openingValue)) {
        const openingValue = { ...metric.openingValue };
        deletedSet.forEach(memberId => {
          if (!Object.prototype.hasOwnProperty.call(openingValue, memberId)) return;
          openingValue[UNASSIGNED_MEMBER_ID] = Number(openingValue[UNASSIGNED_MEMBER_ID] || 0) + Number(openingValue[memberId] || 0);
          delete openingValue[memberId];
        });
        metric.openingValue = openingValue;
      }
      delete metric.controlPoints;
      delete metric.dimensionShareControlPoints;
    });
  });
  clearCalculationCache();
}

function remapSeriesCoordinates(seriesMap, dimensionId, deletedSet) {
  return Object.entries(seriesMap).reduce((nextMap, [key, series]) => {
    const coordinate = parseCoordinateKey(key);
    if (deletedSet.has(coordinate[dimensionId])) {
      coordinate[dimensionId] = UNASSIGNED_MEMBER_ID;
    }
    const nextKey = coordinateKey(coordinate);
    nextMap[nextKey] = nextMap[nextKey] ? sumSeries(nextMap[nextKey], series) : [...series];
    return nextMap;
  }, {});
}

function saveDimensionDefinition() {
  const nameInput = document.getElementById("dimension-name-input");
  const membersInput = document.getElementById("dimension-members-input");
  const label = nameInput.value.trim();
  const members = normalizeDimensionMembersFromText(membersInput.value);
  if (!label) {
    setSourceStatus("Dimension name is required.", "error");
    return;
  }
  if (!members.length) {
    setSourceStatus("Add at least one dimension member.", "error");
    return;
  }

  const existingId = state.selectedDimensionId && dimensionDefinitions()[state.selectedDimensionId]
    ? state.selectedDimensionId
    : null;
  const dimensionId = existingId || uniqueDimensionId(label);
  const existingMembers = existingId ? dimensionMembers(existingId) : [];
  const mergedMembers = existingId
    ? mergeDimensionMembers(existingMembers, members.map(member => member.label))
    : members;
  const nextMemberIds = new Set(mergedMembers.map(member => member.id));
  const deletedMemberIds = existingMembers
    .map(member => member.id)
    .filter(memberId => memberId !== UNASSIGNED_MEMBER_ID && !nextMemberIds.has(memberId));
  const nextMembers = deletedMemberIds.length
    ? ensureUnassignedMember(mergedMembers)
    : mergedMembers;

  dimensionDefinitions()[dimensionId] = { id: dimensionId, label, members: nextMembers };
  state.selectedDimensionId = dimensionId;
  remapDeletedDimensionMembers(dimensionId, deletedMemberIds);
  reconcileMetricsForDimensionChange(dimensionId);
  catalogSourceIsDirty = false;
  setSourceStatus(`Saved dimension ${label}.`, "success");
  render();
}

function newDimensionDraft() {
  state.selectedDimensionId = null;
  document.getElementById("dimension-name-input").value = "";
  document.getElementById("dimension-members-input").value = "";
  document.getElementById("dimension-name-input").focus();
  renderDimensionCatalog();
}

function sanitizeLaggedMetricScenario(metricId) {
  const existing = ensureMetricScenario(metricId);
  scenario().metrics[metricId] = laggedMetricScenarioSnapshot(existing);
  clearCalculationCache();
  return scenario().metrics[metricId];
}

function metricCohorts(metricId) {
  return metricScenario(metricId).cohorts || null;
}

function metricAgeCurve(metricId) {
  return metricScenario(metricId).ageCurve || null;
}

function metricCohortStartSourceId(metricId) {
  return metricScenario(metricId).cohortStartSourceId || "";
}

function metricCohortStarts(metricId) {
  const sourceId = metricCohortStartSourceId(metricId);
  if (sourceId && sourceId !== metricId && metricDefinitions()[sourceId]) {
    const sourceSeries = metricEffectiveSeries(sourceId);
    return Object.fromEntries(YEARS.map((year, index) => [String(year), Number(sourceSeries[index] || 0)]));
  }
  return metricScenario(metricId).cohortStarts || {};
}

function metricCohortAgeYoy(metricId) {
  return metricScenario(metricId).cohortAgeYoy || {};
}

function metricOpeningValue(metricId, memberId = null) {
  const openingValue = metricScenario(metricId).openingValue;
  if (openingValue && typeof openingValue === "object" && !Array.isArray(openingValue)) {
    if (memberId) return Number(openingValue[memberId] ?? 0);
    return Object.values(openingValue).reduce((sum, value) => sum + Number(value || 0), 0);
  }
  return Number(openingValue ?? 0);
}

function metricHasOpeningValue(metricId, memberId = null) {
  const metric = metricScenario(metricId);
  if (!Object.prototype.hasOwnProperty.call(metric, "openingValue")) return false;
  const openingValue = metric.openingValue;
  if (memberId && openingValue && typeof openingValue === "object" && !Array.isArray(openingValue)) {
    return Object.prototype.hasOwnProperty.call(openingValue, memberId);
  }
  return true;
}

function manualControlPoints(metricId) {
  const stored = metricScenario(metricId).controlPoints?.[TOTAL_COORDINATE_KEY];
  if (Array.isArray(stored) && stored.length) return normalizeNapkinPoints(stored);
  return YEARS.map((year, index) => [year, Number(metricTopDownSeries(metricId)[index] || 0)]);
}

function activeTopDownControlPoints(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  if (!context) return manualControlPoints(definition.id);
  const stored = metricScenario(definition.id).controlPoints?.[context.coordinateKey];
  if (Array.isArray(stored) && stored.length) return normalizeNapkinPoints(stored);
  return YEARS.map((year, index) => [year, Number(metricTopDownFilteredSeries(definition.id, context.coordinate)[index] || 0)]);
}

function activeTopDownSeries(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  return context
    ? metricTopDownFilteredSeries(definition.id, context.coordinate)
    : metricTopDownSeries(definition.id);
}

function activeBottomUpSeries(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  return context
    ? metricBottomUpFilteredSeries(definition.id, context.coordinate)
    : metricBottomUpSeries(definition.id);
}

function setActiveTopDownSeries(definition, points) {
  const series = seriesFromControlPoints(points);
  const metric = ensureMetricScenario(definition.id);
  const context = selectedDimensionContextForMetric(definition.id);

  if (context) {
    setMetricTopDownFilteredSeries(definition.id, context.coordinate, series);
    return;
  }

  if (!metric.controlPoints) metric.controlPoints = {};
  metric.controlPoints[TOTAL_COORDINATE_KEY] = points;
  setMetricTopDownSeries(definition.id, series);
}

function normalizeNapkinPoints(points) {
  const byX = new Map();
  (points || []).forEach(point => {
    const x = Number(point?.[0]);
    const y = Number(point?.[1]);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    byX.set(x, y);
  });
  return Array.from(byX.entries()).sort((left, right) => left[0] - right[0]);
}

function ageCurveValue(ageCurve, age) {
  if (!Array.isArray(ageCurve) || !ageCurve.length) return 0;
  if (ageCurve[age] !== undefined) return Number(ageCurve[age] || 0);
  return Number(ageCurve[ageCurve.length - 1] || 0);
}

function interpolateLineValue(points, x) {
  const sorted = [...points].sort((left, right) => Number(left[0]) - Number(right[0]));
  if (!sorted.length) return 0;
  if (x <= sorted[0][0]) return Number(sorted[0][1] || 0);
  if (x >= sorted[sorted.length - 1][0]) return Number(sorted[sorted.length - 1][1] || 0);
  for (let index = 1; index < sorted.length; index += 1) {
    const [leftX, leftY] = sorted[index - 1];
    const [rightX, rightY] = sorted[index];
    if (x >= leftX && x <= rightX) {
      const span = rightX - leftX || 1;
      const progress = (x - leftX) / span;
      return Number(leftY || 0) + (Number(rightY || 0) - Number(leftY || 0)) * progress;
    }
  }
  return Number(sorted[sorted.length - 1][1] || 0);
}

function seriesFromControlPoints(points) {
  return YEARS.map(year => interpolateLineValue(points, year));
}

function cohortTotalSeries(metricId) {
  const cohorts = metricCohorts(metricId);
  if (!cohorts) return null;
  return YEARS.map(year => {
    return Object.values(cohorts).reduce((sum, row) => sum + cohortValueForYear(row, year), 0);
  });
}

function generatedCohortMatrix(metricId) {
  const starts = metricCohortStarts(metricId);
  const ageYoy = metricCohortAgeYoy(metricId);
  const matrix = {};

  YEARS.forEach(cohortYear => {
    const startValue = cohortStartValueForYear(starts, cohortYear);
    if (!startValue) return;
    matrix[String(cohortYear)] = {};
    YEARS.forEach(year => {
      if (year < cohortYear) return;
      if (year === cohortYear) {
        matrix[String(cohortYear)][String(year)] = startValue;
        return;
      }
      const prior = matrix[String(cohortYear)][String(year - 1)] || 0;
      const age = year - cohortYear;
      matrix[String(cohortYear)][String(year)] = prior * (1 + cohortYoyValueForAge(ageYoy, age));
    });
  });

  return matrix;
}

function generatedCohortMatrixTotalSeries(metricId) {
  const matrix = generatedCohortMatrix(metricId);
  return YEARS.map(year => {
    return Object.values(matrix).reduce((sum, row) => sum + cohortValueForYear(row, year), 0);
  });
}

function cohortOpeningCohorts(metricId) {
  const metric = metricScenario(metricId);
  return metric.openingCohorts || {};
}

function cohortAgeCurve(metricId) {
  const metric = metricScenario(metricId);
  return metric.cohortAgeCurve || metric.retentionByAge || metric.cohortAgeYoy || {};
}

function cohortCurveType(definition) {
  return definition?.bottomUp?.curveType || "survivalRate";
}

function cohortStartTiming(definition) {
  return definition?.bottomUp?.startTiming || "currentPeriod";
}

function sourceSeriesForCohortMetric(metricId, filters = {}) {
  const definition = metricDefinitions()[metricId];
  const [sourceMetricId] = metricInputIds(definition);
  if (!sourceMetricId || !metricDefinitions()[sourceMetricId]) {
    const starts = metricCohortStarts(metricId);
    return YEARS.map(year => cohortStartValueForYear(starts, year));
  }
  const sourceFilters = relevantCoordinateFilters(sourceMetricId, { coordinate: filters });
  return Object.keys(sourceFilters).length
    ? metricTopDownFilteredSeries(sourceMetricId, sourceFilters)
    : metricTopDownSeries(sourceMetricId);
}

function cohortMatrixSourceMetricId(metricId) {
  const definition = metricDefinitions()[metricId];
  const [sourceMetricId] = metricInputIds(definition);
  return sourceMetricId && metricDefinitions()[sourceMetricId] ? sourceMetricId : null;
}

function cohortMatrixCoordinateKeys(metricId) {
  const metricKeys = metricDimensionIds(metricId).length ? metricCoordinateKeys(metricId) : [];
  if (metricKeys.length) return metricKeys;
  const sourceMetricId = cohortMatrixSourceMetricId(metricId);
  if (sourceMetricId && metricDimensionIds(sourceMetricId).length) return metricCoordinateKeys(sourceMetricId);
  return [TOTAL_COORDINATE_KEY];
}

function rollupOpeningCohorts(metricId, filters = {}) {
  const rawOpeningCohorts = cohortOpeningCohorts(metricId);
  const hasCoordinateSpecificValues = Object.values(rawOpeningCohorts || {})
    .some(value => value && typeof value === "object" && !Array.isArray(value));
  if (!hasCoordinateSpecificValues) return flatOpeningCohorts(rawOpeningCohorts);

  return Object.entries(rawOpeningCohorts || {})
    .filter(([key, value]) => value && typeof value === "object" && !Array.isArray(value) && coordinateMatchesFilters(key, filters))
    .reduce((rollup, [_key, values]) => {
      Object.entries(values).forEach(([ageKey, value]) => {
        rollup[ageKey] = Number(rollup[ageKey] || 0) + Number(value || 0);
      });
      return rollup;
    }, {});
}

function openingCohortShares(metricId, coordinateKeys) {
  const startSeriesByCoordinate = cohortMatrixStartSeriesByCoordinate(metricId, coordinateKeys);
  const firstYearValues = coordinateKeys.map(key => Number(startSeriesByCoordinate[key]?.[0] || 0));
  const firstYearTotal = firstYearValues.reduce((sum, value) => sum + value, 0);
  const evenShare = coordinateKeys.length ? 1 / coordinateKeys.length : 1;
  return Object.fromEntries(coordinateKeys.map((key, index) => [
    key,
    firstYearTotal ? firstYearValues[index] / firstYearTotal : evenShare,
  ]));
}

function serializeOpeningCohortsForSave(metricId, flatOpenings, filters = {}) {
  const coordinateKeys = cohortMatrixCoordinateKeys(metricId).filter(Boolean);
  if (!coordinateKeys.length) return flatOpenings;

  const rawOpeningCohorts = cohortOpeningCohorts(metricId);
  const hasCoordinateSpecificValues = Object.values(rawOpeningCohorts || {})
    .some(value => value && typeof value === "object" && !Array.isArray(value));
  const next = hasCoordinateSpecificValues ? clone(rawOpeningCohorts) : {};
  const matchingKeys = coordinateKeys.filter(key => coordinateMatchesFilters(key, filters));
  const targetKeys = matchingKeys.length ? matchingKeys : coordinateKeys;
  const shares = openingCohortShares(metricId, targetKeys);

  targetKeys.forEach(key => {
    next[key] = Object.fromEntries(
      Object.entries(flatOpenings || {}).map(([ageKey, value]) => [ageKey, Number(value || 0) * Number(shares[key] || 0)])
    );
  });
  return next;
}

function cohortMatrixStartSeriesByCoordinateForSource(metricId, coordinateKeys, sourceMetricId) {
  if (!sourceMetricId) {
    const starts = metricCohortStarts(metricId);
    const manualSeries = YEARS.map(year => cohortStartValueForYear(starts, year));
    const share = coordinateKeys.length > 1 ? 1 / coordinateKeys.length : 1;
    return Object.fromEntries(coordinateKeys.map(key => [key, splitSeries(manualSeries, share)]));
  }

  const groups = new Map();
  coordinateKeys.forEach(key => {
    const filters = relevantCoordinateFilters(sourceMetricId, { coordinate: parseCoordinateKey(key) });
    const sourceKey = coordinateKey(filters);
    if (!groups.has(sourceKey)) groups.set(sourceKey, { filters, keys: [] });
    groups.get(sourceKey).keys.push(key);
  });

  const seriesByCoordinate = {};
  groups.forEach(group => {
    const sourceSeries = Object.keys(group.filters).length
      ? metricTopDownFilteredSeries(sourceMetricId, group.filters)
      : metricTopDownSeries(sourceMetricId);
    const share = group.keys.length > 1 ? 1 / group.keys.length : 1;
    group.keys.forEach(key => {
      seriesByCoordinate[key] = splitSeries(sourceSeries, share);
    });
  });
  return seriesByCoordinate;
}

function cohortMatrixStartSeriesByCoordinate(metricId, coordinateKeys) {
  return cohortMatrixStartSeriesByCoordinateForSource(metricId, coordinateKeys, cohortMatrixSourceMetricId(metricId));
}

function cohortOpeningCohortsByCoordinate(metricId, coordinateKeys, startSeriesByCoordinate) {
  return splitOpeningCohortsByCoordinate(cohortOpeningCohorts(metricId), coordinateKeys, startSeriesByCoordinate);
}

function cohortAgeCurveForCoordinate(metricId, coordinateKeyValue) {
  const curve = cohortAgeCurve(metricId);
  const direct = curve?.[coordinateKeyValue];
  return direct && typeof direct === "object" && !Array.isArray(direct) ? direct : curve;
}

function generatedGenericCohortMatrixByCoordinate(metricId) {
  const definition = metricDefinitions()[metricId];
  const coordinateKeys = cohortMatrixCoordinateKeys(metricId);
  const startSeriesByCoordinate = cohortMatrixStartSeriesByCoordinate(metricId, coordinateKeys);
  const openingCohortsByCoordinate = cohortOpeningCohortsByCoordinate(metricId, coordinateKeys, startSeriesByCoordinate);
  return Object.fromEntries(coordinateKeys.map(key => [
    key,
    buildGenericCohortMatrix({
      startSeries: startSeriesByCoordinate[key] || defaultSeries(),
      openingCohorts: openingCohortsByCoordinate[key] || {},
      curve: cohortAgeCurveForCoordinate(metricId, key),
      curveType: cohortCurveType(definition),
      startTiming: cohortStartTiming(definition),
    }),
  ]));
}

function generatedGenericCohortMatrix(metricId, filters = {}) {
  return rollupCohortMatrices(generatedGenericCohortMatrixByCoordinate(metricId), filters);
}

function genericCohortMatrixTotalSeries(metricId, filters = {}) {
  return matrixTotalSeries(generatedGenericCohortMatrix(metricId, filters));
}

function cohortMatrixForDefinition(definition) {
  if (!definition) return null;
  const draftMatrix = cohortMatrixDraftForDefinition(definition);
  if (draftMatrix) return draftMatrix;
  if (definition.bottomUp?.type === "cohortMatrix") {
    const context = selectedDimensionContextForMetric(definition.id);
    return generatedGenericCohortMatrix(definition.id, context?.coordinate || {});
  }
  if (selectedFormulaType() === "cohortMatrix") {
    const context = selectedDimensionContextForMetric(definition.id);
    return generatedGenericCohortMatrix(definition.id, context?.coordinate || {});
  }
  if (definition.bottomUp?.type === "cohortMatrixFromStartsAndAgeYoy") {
    return generatedCohortMatrix(definition.id);
  }
  return null;
}

function cohortMatrixDraftForDefinition(definition) {
  const form = document.getElementById("cohort-matrix-builder-form");
  if (!form || form.classList.contains("is-hidden")) return null;
  const selectedType = selectedFormulaType() || definition?.bottomUp?.type;
  if (selectedType !== "cohortMatrix") return null;

  const sourceId = document.getElementById("cohort-start-source-input")?.value || "";
  const context = selectedDimensionContextForMetric(definition.id);
  const sourceFilters = sourceId && metricDefinitions()[sourceId]
    ? relevantCoordinateFilters(sourceId, context)
    : {};
  const startSeries = sourceId && metricDefinitions()[sourceId]
    ? Object.keys(sourceFilters).length
      ? metricTopDownFilteredSeries(sourceId, sourceFilters)
      : metricTopDownSeries(sourceId)
    : YEARS.map(year => Number(document.querySelector(`[data-cohort-start-year="${year}"]`)?.value || 0));
  const curve = {};
  document.querySelectorAll("[data-cohort-yoy-age]").forEach(input => {
    curve[input.dataset.cohortYoyAge] = Number(input.value || 0) / 100;
  });
  const openingCohorts = {};
  document.querySelectorAll("[data-cohort-opening-age]").forEach(input => {
    openingCohorts[`age_${input.dataset.cohortOpeningAge}`] = Number(input.value || 0);
  });

  return buildGenericCohortMatrix({
    startSeries,
    openingCohorts,
    curve,
    curveType: document.getElementById("cohort-curve-type-input")?.value || "survivalRate",
    startTiming: document.getElementById("cohort-start-timing-input")?.value || "currentPeriod",
  });
}

function cohortMatrixRowLabel(rowKey) {
  if (String(rowKey).startsWith("opening_age_")) {
    return `Opening Age ${String(rowKey).replace("opening_age_", "")}`;
  }
  if (String(rowKey).startsWith("opening_")) {
    return `Opening ${String(rowKey).replace("opening_", "").replace(/_/g, " ")}`;
  }
  return `Cohort ${rowKey}`;
}

function cohortMatrixCellApplies(rowKey, year) {
  if (String(rowKey).startsWith("opening_")) return true;
  const cohortYear = Number(rowKey);
  return Number.isFinite(cohortYear) && year >= cohortYear;
}

function coordinateMemberLabel(dimensionId, memberId) {
  const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
  return member?.label || memberId || "All";
}

function cohortPreviewDimensionIds(definition, coordinateKeys) {
  const dimensionIds = new Set(metricDimensionIds(definition?.id));
  coordinateKeys.forEach(key => {
    Object.keys(parseCoordinateKey(key)).forEach(dimensionId => dimensionIds.add(dimensionId));
  });
  return [...dimensionIds];
}

function cohortPreviewEntries(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  const draftMatrix = state.cohortBuilderIsDirty ? cohortMatrixDraftForDefinition(definition) : null;
  if (draftMatrix) {
    const coordinateKeys = cohortMatrixCoordinateKeys(definition.id);
    if (coordinateKeys.length <= 1 && !coordinateKeys[0]) return [[coordinateKey(context?.coordinate || {}), draftMatrix]];

    const sourceId = document.getElementById("cohort-start-source-input")?.value || "";
    const manualSeries = YEARS.map(year => Number(document.querySelector(`[data-cohort-start-year="${year}"]`)?.value || 0));
    const startSeriesByCoordinate = sourceId && metricDefinitions()[sourceId]
      ? cohortMatrixStartSeriesByCoordinateForSource(definition.id, coordinateKeys, sourceId)
      : Object.fromEntries(coordinateKeys.map(key => [key, splitSeries(manualSeries, coordinateKeys.length > 1 ? 1 / coordinateKeys.length : 1)]));
    const curve = {};
    document.querySelectorAll("[data-cohort-yoy-age]").forEach(input => {
      curve[input.dataset.cohortYoyAge] = Number(input.value || 0) / 100;
    });
    const openingCohorts = {};
    document.querySelectorAll("[data-cohort-opening-age]").forEach(input => {
      openingCohorts[`age_${input.dataset.cohortOpeningAge}`] = Number(input.value || 0);
    });
    const openingByCoordinate = splitOpeningCohortsByCoordinate(openingCohorts, coordinateKeys, startSeriesByCoordinate);
    const filters = context?.coordinate || {};
    return coordinateKeys
      .filter(key => coordinateMatchesFilters(key, filters))
      .map(key => [key, buildGenericCohortMatrix({
        startSeries: startSeriesByCoordinate[key] || defaultSeries(),
        openingCohorts: openingByCoordinate[key] || {},
        curve,
        curveType: document.getElementById("cohort-curve-type-input")?.value || "survivalRate",
        startTiming: document.getElementById("cohort-start-timing-input")?.value || "currentPeriod",
      })]);
  }
  if (definition?.bottomUp?.type === "cohortMatrix") {
    const result = metricResult(definition.id);
    const filters = context?.coordinate || {};
    return Object.entries(result.matrixByCoordinate || {})
      .filter(([key]) => coordinateMatchesFilters(key, filters))
      .sort(([left], [right]) => coordinateLabelForMetric(definition.id, left).localeCompare(coordinateLabelForMetric(definition.id, right)));
  }
  const matrix = cohortMatrixForDefinition(definition);
  return matrix ? [[TOTAL_COORDINATE_KEY, matrix]] : [];
}

function metricResult(metricId, context = null) {
  const definition = metricDefinitions()[metricId];
  const filters = relevantCoordinateFilters(metricId, context);
  if (definition?.bottomUp?.type === "cohortMatrix") {
    const matrixByCoordinate = generatedGenericCohortMatrixByCoordinate(metricId);
    const matrix = rollupCohortMatrices(matrixByCoordinate, filters);
    return {
      shape: "cohortMatrix",
      annualSeries: matrixTotalSeries(matrix),
      matrix,
      matrixByCoordinate,
    };
  }
  const annualSeries = Object.keys(filters).length
    ? metricBottomUpFilteredSeries(metricId, filters) || metricTopDownFilteredSeries(metricId, filters)
    : metricBottomUpSeries(metricId) || metricTopDownSeries(metricId);
  return {
    shape: "series",
    annualSeries,
    valuesByCoordinate: metricSeriesMap(metricId),
  };
}

function metricInputIds(definition) {
  const rawInputs = (definition?.bottomUp?.type === "ratio" || definition?.bottomUp?.type === "weightedAverage") && (!definition.bottomUp.inputs || !definition.bottomUp.inputs.length)
    ? [definition.bottomUp.numerator || definition.bottomUp.valueMetricId, definition.bottomUp.denominator || definition.bottomUp.weightMetricId]
    : definition?.bottomUp?.inputs || [];
  return rawInputs.map(input => {
    if (typeof input === "string") return input;
    return input?.metricId;
  }).filter(Boolean);
}

function weightedAverageConfig(definition) {
  const inputs = metricInputIds(definition);
  const hasExplicitValueMetric = Boolean(definition?.bottomUp?.valueMetricId || definition?.bottomUp?.numerator);
  const [firstInput = "", secondInput = ""] = inputs;
  return {
    valueMetricId: definition?.bottomUp?.valueMetricId || definition?.bottomUp?.numerator || (hasExplicitValueMetric ? firstInput : ""),
    weightMetricId: definition?.bottomUp?.weightMetricId || definition?.bottomUp?.denominator || (hasExplicitValueMetric ? secondInput : firstInput),
  };
}

function lagForDefinition(definition) {
  const rawLag = Number(definition?.bottomUp?.lag ?? 1);
  return Number.isFinite(rawLag) && rawLag > 0 ? Math.floor(rawLag) : 1;
}

function laggedOpeningValue(definition, sourceMetricId, memberId = null) {
  if (metricHasOpeningValue(definition.id, memberId)) {
    return metricOpeningValue(definition.id, memberId);
  }
  return metricOpeningValue(sourceMetricId, memberId);
}

function laggedSeriesFromSource(definition, sourceMetricId, sourceSeries, memberId = null) {
  const lag = lagForDefinition(definition);
  return YEARS.map((_year, index) => {
    const sourceIndex = index - lag;
    if (sourceIndex < 0) return laggedOpeningValue(definition, sourceMetricId, memberId);
    return Number(sourceSeries?.[sourceIndex] || 0);
  });
}

function topDownSeriesForContext(metricId, context = null) {
  const filters = relevantCoordinateFilters(metricId, context);
  return Object.keys(filters).length
    ? metricTopDownFilteredSeries(metricId, filters)
    : metricTopDownSeries(metricId);
}

function bottomUpSeriesForContext(metricId, context = null) {
  const filters = relevantCoordinateFilters(metricId, context);
  return Object.keys(filters).length
    ? metricBottomUpFilteredSeries(metricId, filters)
    : metricBottomUpSeries(metricId);
}

function laggedComparisonLines(definition, context = null) {
  const [sourceMetricId] = metricInputIds(definition);
  if (!sourceMetricId || !metricDefinitions()[sourceMetricId]) return null;
  const sourceDefinition = metricDefinitions()[sourceMetricId];
  const sourceFilters = relevantCoordinateFilters(sourceMetricId, context);
  const sourceContext = Object.keys(sourceFilters).length ? { coordinate: sourceFilters } : null;
  const memberId = Object.keys(sourceFilters).length === 1 ? Object.values(sourceFilters)[0] : null;
  const topDown = laggedSeriesFromSource(
    definition,
    sourceMetricId,
    topDownSeriesForContext(sourceMetricId, sourceContext),
    memberId
  );
  const bottomUpSource = bottomUpSeriesForContext(sourceMetricId, sourceContext);
  const bottomUp = bottomUpSource
    ? laggedSeriesFromSource(definition, sourceMetricId, bottomUpSource, memberId)
    : null;
  const sourceLabel = context?.member
    ? `${sourceDefinition.label}: ${context.member.label}`
    : context?.label
      ? `${sourceDefinition.label}: ${context.label}`
    : sourceDefinition.label;
  return { sourceLabel, topDown, bottomUp };
}

function canvasNodeSeries(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  const definition = metricDefinitions()[parsed.metricId];
  if (!definition) return defaultSeries();
  const context = contextForNodeKey(nodeKey);
  if (metricIsLagged(definition)) {
    const laggedLines = laggedComparisonLines(definition, context);
    if (laggedLines?.topDown) return laggedLines.topDown;
  }
  return context
    ? metricTopDownFilteredSeries(parsed.metricId, context.coordinate)
    : metricTopDownSeries(parsed.metricId);
}

function contextForNodeKey(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  if (!parsed.memberId) return null;
  return {
    metricId: parsed.metricId,
    coordinate: { [parsed.dimensionId]: parsed.memberId },
    dimensionId: parsed.dimensionId,
    memberId: parsed.memberId,
    member: dimensionMembers(parsed.dimensionId).find(item => item.id === parsed.memberId),
  };
}

function canvasNodeComparisonSeries(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  const definition = metricDefinitions()[parsed.metricId];
  if (!definition) return { topDown: defaultSeries(), bottomUp: null };
  const context = contextForNodeKey(nodeKey);
  if (metricIsLagged(definition)) {
    const laggedLines = laggedComparisonLines(definition, context);
    return {
      topDown: laggedLines?.topDown || defaultSeries(),
      bottomUp: laggedLines?.bottomUp || null,
    };
  }
  return {
    topDown: context
      ? metricTopDownFilteredSeries(parsed.metricId, context.coordinate)
      : metricTopDownSeries(parsed.metricId),
    bottomUp: context
      ? metricBottomUpFilteredSeries(parsed.metricId, context.coordinate)
      : metricBottomUpSeries(parsed.metricId),
  };
}

function relevantCoordinateFilters(metricId, context = null) {
  const coordinate = context?.coordinate || (
    context?.dimensionId && context?.memberId
      ? { [context.dimensionId]: context.memberId }
      : {}
  );
  return Object.fromEntries(
    Object.entries(coordinate).filter(([dimensionId]) => metricUsesDimension(metricId, dimensionId))
  );
}

function formulaInputDimensionIds(definition) {
  const ids = new Set();
  metricInputIds(definition).forEach(inputId => {
    metricDimensionIds(inputId).forEach(dimensionId => ids.add(dimensionId));
  });
  return [...ids];
}

function coordinateKeysForDimensionIds(dimensionIds = []) {
  if (!dimensionIds.length) return [TOTAL_COORDINATE_KEY];
  const build = (index, coordinate) => {
    if (index >= dimensionIds.length) return [coordinateKey(coordinate)];
    const dimensionId = dimensionIds[index];
    const members = dimensionMembers(dimensionId);
    if (!members.length) return build(index + 1, coordinate);
    return members.flatMap(member => build(index + 1, { ...coordinate, [dimensionId]: member.id }));
  };
  return build(0, {});
}

function unsupportedFormulaFilterDimensions(definition, filters = {}) {
  const formulaDimensionIds = formulaInputDimensionIds(definition);
  return Object.keys(filters).filter(dimensionId => !formulaDimensionIds.includes(dimensionId));
}

function inputCanUseFilters(inputId, filters = {}) {
  const inputDimensionIds = metricDimensionIds(inputId);
  return Object.keys(filters).every(dimensionId => inputDimensionIds.includes(dimensionId));
}

function formulaCoordinateKeys(definition, filters = {}) {
  const formulaDimensionIds = formulaInputDimensionIds(definition);
  if (unsupportedFormulaFilterDimensions(definition, filters).length) return null;
  return coordinateKeysForDimensionIds(formulaDimensionIds)
    .filter(key => coordinateMatchesFilters(key, filters));
}

function weightedAverageRateCoordinateKeys(definition, weightMetricId, filters = {}) {
  const rateDimensionIds = metricDimensionIds(definition.id);
  const weightDimensionIds = metricDimensionIds(weightMetricId);
  const unsupportedFilters = Object.keys(filters).filter(dimensionId => (
    !rateDimensionIds.includes(dimensionId) || !weightDimensionIds.includes(dimensionId)
  ));
  if (unsupportedFilters.length) return null;
  if (!rateDimensionIds.every(dimensionId => weightDimensionIds.includes(dimensionId))) return null;
  return (rateDimensionIds.length ? coordinateKeysForDimensionIds(rateDimensionIds) : [TOTAL_COORDINATE_KEY])
    .filter(key => coordinateMatchesFilters(key, filters));
}

function weightedAverageTotalsAtIndex(definition, index, filters = {}) {
  const { valueMetricId, weightMetricId } = weightedAverageConfig(definition);
  if (!weightMetricId || !metricDefinitions()[weightMetricId]) return null;
  if (valueMetricId) {
    if (![valueMetricId, weightMetricId].every(inputId => inputCanUseFilters(inputId, filters))) return null;
    const coordinateKeys = formulaCoordinateKeys(definition, filters);
    if (!coordinateKeys) return null;
    return coordinateKeys.reduce((acc, key) => {
      const coordinate = parseCoordinateKey(key);
      const valueSeries = metricTopDownContextSeries(valueMetricId, { coordinate });
      const weightSeries = metricTopDownContextSeries(weightMetricId, { coordinate });
      acc.value += Number(valueSeries[index] || 0);
      acc.weight += Number(weightSeries[index] || 0);
      return acc;
    }, { value: 0, weight: 0 });
  }

  const coordinateKeys = weightedAverageRateCoordinateKeys(definition, weightMetricId, filters);
  if (!coordinateKeys) return null;
  return coordinateKeys.reduce((acc, key) => {
    const coordinate = parseCoordinateKey(key);
    const rateSeries = key
      ? metricTopDownFilteredSeries(definition.id, coordinate)
      : metricTopDownSeries(definition.id);
    const weightSeries = metricTopDownContextSeries(weightMetricId, { coordinate });
    const weight = Number(weightSeries[index] || 0);
    acc.value += Number(rateSeries[index] || 0) * weight;
    acc.weight += weight;
    return acc;
  }, { value: 0, weight: 0 });
}

function metricValueAtIndex(metricId, index, stack = new Set(), memberContext = null) {
  const definition = metricDefinitions()[metricId];
  const filters = relevantCoordinateFilters(metricId, memberContext);
  const filterKey = coordinateKey(filters);
  const memberId = Object.keys(filters).length === 1 ? Object.values(filters)[0] : null;
  if (!definition?.bottomUp) {
    const series = filterKey ? metricTopDownFilteredSeries(metricId, filters) : metricTopDownSeries(metricId);
    return Number(series[index] || 0);
  }

  const stackKey = `${metricId}:${filterKey || "total"}:${index}`;
  if (stack.has(stackKey)) {
    const series = filterKey ? metricTopDownFilteredSeries(metricId, filters) : metricTopDownSeries(metricId);
    return Number(series[index] || 0);
  }

  const nextStack = new Set(stack);
  nextStack.add(stackKey);
  const inputs = metricInputIds(definition);
  const inputValue = (inputId, currentIndex = index, coordinateFilters = filters) => {
    const series = metricTopDownContextSeries(inputId, { coordinate: coordinateFilters });
    return Number(series[currentIndex] || 0);
  };

  if (definition.bottomUp.type === "sum") {
    if (!inputs.every(inputId => inputCanUseFilters(inputId, filters))) return null;
    return inputs.reduce((sum, inputId) => sum + inputValue(inputId), 0);
  }

  if (definition.bottomUp.type === "product") {
    const coordinateKeys = formulaCoordinateKeys(definition, filters);
    if (!coordinateKeys) return null;
    return coordinateKeys.reduce((sum, key) => {
      const coordinate = parseCoordinateKey(key);
      const value = inputs.reduce((product, inputId) => product * inputValue(inputId, index, coordinate), 1);
      return sum + value;
    }, 0);
  }

  if (definition.bottomUp.type === "ratio") {
    const [numeratorId, denominatorId] = inputs;
    if (!inputs.every(inputId => inputCanUseFilters(inputId, filters))) return null;
    const numerator = inputValue(numeratorId);
    const denominator = inputValue(denominatorId);
    if (!denominator) return null;
    return numerator / denominator;
  }

  if (definition.bottomUp.type === "weightedAverage") {
    const totals = weightedAverageTotalsAtIndex(definition, index, filters);
    if (!totals) return null;
    if (!totals.weight) return null;
    return totals.value / totals.weight;
  }

  if (definition.bottomUp.type === "difference") {
    if (!inputs.every(inputId => inputCanUseFilters(inputId, filters))) return null;
    const [leftId, rightId] = inputs;
    return inputValue(leftId) - inputValue(rightId);
  }

  if (definition.bottomUp.type === "cumulativeNetFlow") {
    if (!inputs.every(inputId => inputCanUseFilters(inputId, filters))) return null;
    const [inflowId, outflowId] = inputs;
    let runningValue = metricOpeningValue(metricId, memberId);
    for (let currentIndex = 0; currentIndex <= index; currentIndex += 1) {
      runningValue += inputValue(inflowId, currentIndex) - inputValue(outflowId, currentIndex);
    }
    return runningValue;
  }

  if (definition.bottomUp.type === "laggedMetric") {
    const [sourceMetricId] = inputs;
    if (!inputCanUseFilters(sourceMetricId, filters)) return null;
    const sourceIndex = index - lagForDefinition(definition);
    if (sourceIndex < 0) return laggedOpeningValue(definition, sourceMetricId, memberId);
    return inputValue(sourceMetricId, sourceIndex);
  }

  if (definition.bottomUp.type === "laggedRetentionFlow") {
    const [stockMetricId, retentionMetricId] = inputs;
    if (!inputs.every(inputId => inputCanUseFilters(inputId, filters))) return null;
    const priorStock = index <= 0
      ? metricOpeningValue(metricId, memberId) || metricOpeningValue(stockMetricId, memberId)
      : inputValue(stockMetricId, index - 1);
    return priorStock * inputValue(retentionMetricId);
  }

  if (definition.bottomUp.type === "cohortMatrix") {
    return Number(genericCohortMatrixTotalSeries(metricId, filters)?.[index] || 0);
  }

  const series = filterKey ? metricTopDownFilteredSeries(metricId, filters) : metricTopDownSeries(metricId);
  return Number(series[index] || 0);
}

function metricBottomUpSeries(metricId) {
  return cachedCalculation(`bottomUp:${metricId}`, () => {
    const definition = metricDefinitions()[metricId];
    if (!definition?.bottomUp) return null;

    if (definition.bottomUp.type === "cohortAgeProduct") {
      const [cohortMetricId, ageCurveMetricId] = metricInputIds(definition);
      const cohorts = metricCohorts(cohortMetricId);
      const ageCurve = metricAgeCurve(ageCurveMetricId);
      if (!cohorts || !ageCurve) return null;
      return YEARS.map(year => {
        return Object.entries(cohorts).reduce((sum, [cohortYear, row]) => {
          const age = year - Number(cohortYear);
          if (age < 0) return sum;
          return sum + cohortValueForYear(row, year) * ageCurveValue(ageCurve, age);
        }, 0);
      });
    }

    if (definition.bottomUp.type === "cohortMatrixFromStartsAndAgeYoy") {
      return generatedCohortMatrixTotalSeries(metricId);
    }

    if (definition.bottomUp.type === "cohortMatrix") {
      return metricResult(metricId).annualSeries;
    }

    const series = YEARS.map((_year, index) => metricValueAtIndex(metricId, index));
    return series.some(value => value === null || value === undefined) ? null : series;
  });
}

function metricBottomUpMemberSeries(metricId, dimensionId, memberId) {
  return cachedCalculation(`bottomUpMember:${metricId}:${dimensionId}:${memberId}`, () => {
    const definition = metricDefinitions()[metricId];
    if (!definition?.bottomUp || !metricUsesDimension(metricId, dimensionId)) return null;
    if (definition.bottomUp.type === "cohortMatrix") {
      return metricResult(metricId, { coordinate: { [dimensionId]: memberId } }).annualSeries;
    }
    const series = YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { coordinate: { [dimensionId]: memberId } }));
    return series.some(value => value === null || value === undefined) ? null : series;
  });
}

function metricBottomUpFilteredSeries(metricId, filters = {}) {
  return cachedCalculation(`bottomUpFiltered:${metricId}:${coordinateKey(filters)}`, () => {
    const definition = metricDefinitions()[metricId];
    if (!definition?.bottomUp) return null;
    if (definition.bottomUp.type === "cohortMatrix") {
      return metricResult(metricId, { coordinate: filters }).annualSeries;
    }
    const series = YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { coordinate: filters }));
    return series.some(value => value === null || value === undefined) ? null : series;
  });
}

function metricTopDownContextSeries(metricId, context = null) {
  const filters = relevantCoordinateFilters(metricId, context);
  return Object.keys(filters).length
    ? metricTopDownFilteredSeries(metricId, filters)
    : metricTopDownSeries(metricId);
}

function metricDirectBottomUpSeries(metricId, context = null) {
  const definition = metricDefinitions()[metricId];
  if (!definition?.bottomUp) return null;
  const filters = relevantCoordinateFilters(metricId, context);
  const series = YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { coordinate: filters }));
  return series.some(value => value === null || value === undefined) ? null : series;
}

function activeWorkspaceBottomUpSeries(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  return metricDirectBottomUpSeries(definition?.id, context);
}

function metricEffectiveSeries(metricId) {
  return cachedCalculation(`effective:${metricId}`, () => metricBottomUpSeries(metricId) || cohortTotalSeries(metricId) || metricTopDownSeries(metricId));
}

function metricEffectiveMemberSeries(metricId, dimensionId, memberId) {
  return cachedCalculation(
    `effectiveMember:${metricId}:${dimensionId}:${memberId}`,
    () => metricBottomUpMemberSeries(metricId, dimensionId, memberId) || metricTopDownMemberSeries(metricId, memberId)
  );
}

function metricReconciliation(metricId) {
  const definition = metricDefinitions()[metricId];
  if (!definition?.reconciliation?.enabled) return { enabled: false, matched: true, maxDelta: 0 };
  const laggedLines = metricIsLagged(definition) ? laggedComparisonLines(definition) : null;
  const topDown = laggedLines?.topDown || metricTopDownSeries(metricId);
  const bottomUp = laggedLines ? laggedLines.bottomUp : metricBottomUpSeries(metricId);
  if (!bottomUp) return { enabled: false, matched: true, maxDelta: 0 };
  const maxDelta = Math.max(...YEARS.map((_year, index) => Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0))));
  return {
    enabled: true,
    matched: maxDelta <= Number(definition.reconciliation.tolerance || 0),
    maxDelta,
  };
}

function metricMemberReconciliation(metricId, dimensionId, memberId) {
  const definition = metricDefinitions()[metricId];
  if (!definition?.reconciliation?.enabled || !metricUsesDimension(metricId, dimensionId)) {
    return { enabled: false, matched: true, maxDelta: 0 };
  }
  const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
  const context = { metricId, dimensionId, memberId, member };
  const laggedLines = metricIsLagged(definition) ? laggedComparisonLines(definition, context) : null;
  const topDown = laggedLines?.topDown || metricTopDownMemberSeries(metricId, memberId);
  const bottomUp = laggedLines ? laggedLines.bottomUp : metricBottomUpMemberSeries(metricId, dimensionId, memberId);
  if (!bottomUp) return { enabled: false, matched: true, maxDelta: 0 };
  const maxDelta = Math.max(...YEARS.map((_year, index) => Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0))));
  return {
    enabled: true,
    matched: maxDelta <= Number(definition.reconciliation.tolerance || 0),
    maxDelta,
  };
}

function metricFilteredReconciliation(metricId, context) {
  const definition = metricDefinitions()[metricId];
  if (!definition?.reconciliation?.enabled || !context) {
    return { enabled: false, matched: true, maxDelta: 0 };
  }
  const laggedLines = metricIsLagged(definition) ? laggedComparisonLines(definition, context) : null;
  const topDown = laggedLines?.topDown || metricTopDownFilteredSeries(metricId, context.coordinate);
  const bottomUp = laggedLines ? laggedLines.bottomUp : metricBottomUpFilteredSeries(metricId, context.coordinate);
  if (!bottomUp) return { enabled: false, matched: true, maxDelta: 0 };
  const maxDelta = Math.max(...YEARS.map((_year, index) => Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0))));
  return {
    enabled: true,
    matched: maxDelta <= Number(definition.reconciliation.tolerance || 0),
    maxDelta,
  };
}

function activeScenarioReconciliationSummary() {
  const items = Object.values(metricDefinitions())
    .map(definition => ({
      metricId: definition.id,
      label: definition.label,
      ...metricReconciliation(definition.id),
    }))
    .filter(item => item.enabled);
  const openItems = items
    .filter(item => !item.matched)
    .sort((left, right) => Number(right.maxDelta || 0) - Number(left.maxDelta || 0));
  return {
    scenarioId: activeScenarioId(),
    checkedMetricCount: items.length,
    openMetricCount: openItems.length,
    matched: openItems.length === 0,
    openItems,
  };
}

function childIds(metricId) {
  const definition = metricDefinitions()[metricId];
  if (definition?.bottomUp?.type === "weightedAverage") return [];
  if (definition?.bottomUp?.type === "laggedMetric") return [];
  if (definition?.bottomUp?.type === "laggedRetentionFlow") return Array.from(new Set(metricInputIds(definition).slice(1)));
  const ids = metricInputIds(definition);
  if (definition?.bottomUp?.type === "cohortMatrixFromStartsAndAgeYoy") {
    const startSourceId = metricCohortStartSourceId(metricId);
    if (startSourceId && metricDefinitions()[startSourceId]) ids.push(startSourceId);
  }
  if (definition?.bottomUp?.type === "cohortMatrix") {
    const [startSourceId] = metricInputIds(definition);
    if (startSourceId && metricDefinitions()[startSourceId]) ids.push(startSourceId);
  }
  return Array.from(new Set(ids));
}

function nodeChildKeys(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  if (parsed.memberId) return [];
  return childIds(parsed.metricId);
}

function parentIds(metricId) {
  return Object.values(metricDefinitions())
    .filter(definition => childIds(definition.id).includes(metricId))
    .map(definition => definition.id);
}

function descendantIds(metricId, seen = new Set()) {
  childIds(metricId).forEach(childId => {
    if (seen.has(childId)) return;
    seen.add(childId);
    descendantIds(childId, seen);
  });
  return seen;
}

function rootMetricIds() {
  const definitions = Object.values(metricDefinitions());
  const childSet = new Set(definitions.flatMap(definition => childIds(definition.id)));
  return definitions
    .filter(definition => !childSet.has(definition.id))
    .map(definition => definition.id);
}

function canvasLevels() {
  const roots = rootMetricIds();
  const levels = [];
  const visited = new Set();
  let current = roots.length ? roots : Object.keys(metricDefinitions());
  while (current.length) {
    const level = Array.from(new Set(current)).filter(nodeKey => {
      const parsed = parseNodeKey(nodeKey);
      return metricDefinitions()[parsed.metricId] && !visited.has(nodeKey);
    });
    if (!level.length) break;
    levels.push(level);
    level.forEach(nodeKey => visited.add(nodeKey));
    current = level.flatMap(nodeChildKeys);
  }
  const unvisited = Object.keys(metricDefinitions()).filter(metricId => !visited.has(metricId));
  if (unvisited.length) levels.push(unvisited);
  return levels;
}

function formatCurrency(value) {
  return currencyFormatter.format(Number(value || 0));
}

function formatCurrencyPrecise(value) {
  return preciseCurrencyFormatter.format(Number(value || 0));
}

function trimFixed(value, decimals) {
  return Number(value).toFixed(decimals).replace(/\.0+$|(\.\d*[1-9])0+$/, "$1");
}

function compactMagnitude(value, divisor, suffix, precisionOffset = 0) {
  const scaled = value / divisor;
  const absScaled = Math.abs(scaled);
  const baseDecimals = absScaled < 10 && !Number.isInteger(scaled)
    ? 1
    : absScaled < 100 && Math.abs(scaled - Math.round(scaled)) >= 0.1
      ? 1
      : 0;
  const decimals = Math.min(3, baseDecimals + precisionOffset);
  return `${trimFixed(scaled, decimals)}${suffix}`;
}

function compactCurrency(value) {
  const numericValue = Number(value || 0);
  const abs = Math.abs(numericValue);
  const sign = numericValue < 0 ? "-" : "";
  if (abs >= 1000000) return `${sign}$${compactMagnitude(abs, 1000000, "M")}`;
  if (abs >= 1000) return `${sign}$${compactMagnitude(abs, 1000, "k")}`;
  return formatCurrency(numericValue);
}

function compactCurrencyTooltip(value) {
  const numericValue = Number(value || 0);
  const abs = Math.abs(numericValue);
  const sign = numericValue < 0 ? "-" : "";
  if (abs >= 1000000) return `${sign}$${compactMagnitude(abs, 1000000, "M", 1)}`;
  if (abs >= 1000) return `${sign}$${compactMagnitude(abs, 1000, "k", 1)}`;
  return formatCurrencyPrecise(numericValue);
}

function formatMetricValue(definition, value, { precise = false } = {}) {
  const numericValue = Number(value || 0);
  if (definition?.unit === "percent") {
    return `${trimFixed(numericValue * 100, precise ? 2 : 1)}%`;
  }
  if (definition?.unit === "count") {
    return new Intl.NumberFormat("en-US", {
      maximumFractionDigits: precise ? 2 : 0,
    }).format(numericValue);
  }
  return precise ? compactCurrencyTooltip(numericValue) : compactCurrency(numericValue);
}

function napkinSnapStep(maxY) {
  if (!Number.isFinite(maxY) || maxY <= 0) return 1;
  return 5 * (10 ** (Math.floor(Math.log10(maxY)) - 2));
}

function snapSafeNapkinYMax(rawMax) {
  let yMax = Number(rawMax || 0);
  if (!Number.isFinite(yMax) || yMax <= 0) return 10;
  for (let index = 0; index < 3; index += 1) {
    const step = napkinSnapStep(yMax);
    if (!Number.isFinite(step) || step <= 0) break;
    yMax = Math.ceil(yMax / step) * step;
  }
  return yMax;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formulaText(metricId) {
  const definition = metricDefinitions()[metricId];
  if (!definition?.bottomUp) return "No bottom-up definition";
  const inputIds = metricInputIds(definition);
  const labels = inputIds.map(inputId => metricDefinitions()[inputId]?.label || inputId);
  if (definition.bottomUp.type === "sum") return labels.length ? labels.join(" + ") : "Empty sum";
  if (definition.bottomUp.type === "product") return labels.length ? labels.join(" × ") : "Empty product";
  if (definition.bottomUp.type === "ratio") return labels.length === 2 ? `${labels[0]} / ${labels[1]}` : "Empty ratio";
  if (definition.bottomUp.type === "weightedAverage") {
    const { valueMetricId, weightMetricId } = weightedAverageConfig(definition);
    const weightLabel = metricDefinitions()[weightMetricId]?.label || "weight";
    if (!valueMetricId) return `weighted average by ${weightLabel}`;
    const valueLabel = metricDefinitions()[valueMetricId]?.label || "weighted value";
    return `weighted average: ${valueLabel} / ${weightLabel}`;
  }
  if (definition.bottomUp.type === "difference") return labels.length ? labels.join(" - ") : "Empty difference";
  if (definition.bottomUp.type === "cohortAgeProduct") {
    return labels.length === 2 ? `${labels[0]} cohorts × ${labels[1]} age curve` : "Cohort stock × age curve";
  }
  if (definition.bottomUp.type === "cohortMatrixFromStartsAndAgeYoy") {
    const sourceId = metricCohortStartSourceId(metricId);
    const sourceLabel = metricDefinitions()[sourceId]?.label;
    return sourceLabel
      ? `${sourceLabel} as starting cohorts × YoY change by cohort age`
      : "starting cohort values × YoY change by cohort age";
  }
  if (definition.bottomUp.type === "cohortMatrix") {
    const [sourceId] = metricInputIds(definition);
    const sourceLabel = metricDefinitions()[sourceId]?.label || "starting cohorts";
    const curveLabel = cohortCurveType(definition) === "survivalRate" ? "survival by cohort age" : "YoY change by cohort age";
    return `${sourceLabel} × ${curveLabel}`;
  }
  if (definition.bottomUp.type === "cumulativeNetFlow") {
    return labels.length === 2 ? `opening value + ${labels[0]} - ${labels[1]}` : "opening value + additions - losses";
  }
  if (definition.bottomUp.type === "laggedRetentionFlow") {
    return labels.length === 2 ? `prior ${labels[0]} × ${labels[1]}` : "prior stock × retention ratio";
  }
  if (definition.bottomUp.type === "laggedMetric") {
    return labels.length === 1 ? `${labels[0]} t-${lagForDefinition(definition)}` : "source metric t-1";
  }
  return definition.bottomUp.type;
}

function metricSlug(label) {
  const base = String(label)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return base || "metric";
}

function uniqueMetricId(label) {
  const base = metricSlug(label);
  let id = base;
  let suffix = 2;
  while (metricDefinitions()[id]) {
    id = `${base}-${suffix}`;
    suffix += 1;
  }
  return id;
}

function uniqueScenarioId() {
  let index = Object.keys(scenarioCollection()).length + 1;
  let id = `scenario-${index}`;
  while (scenarioCollection()[id]) {
    index += 1;
    id = `scenario-${index}`;
  }
  return id;
}

function uniqueScenarioName() {
  const usedNames = new Set(Object.values(scenarioCollection()).map(item => item.name));
  let index = Object.keys(scenarioCollection()).length + 1;
  let name = `Scenario ${index}`;
  while (usedNames.has(name)) {
    index += 1;
    name = `Scenario ${index}`;
  }
  return name;
}

function sourceSnapshot() {
  return JSON.stringify({
    name: state.model.name,
    years: YEARS,
    selectedMetricId: state.selectedMetricId,
    selectedDimensionContext: state.selectedDimensionContext,
    activeScenarioId: activeScenarioId(),
    scenarioRoles: scenarioRoles(),
    referenceScenarios: referenceScenarios(),
    dimensions: dimensionDefinitions(),
    metricDefinitions: metricDefinitions(),
    assumptions: assumptions(),
    reconciliationSummary: activeScenarioReconciliationSummary(),
    comparisonSummary: comparisonSummary(),
    scenarios: scenariosSnapshot(),
  }, null, 2);
}

function jsonExplorerPath(parentPath, key) {
  return `${parentPath}/${String(key).replace(/~/g, "~0").replace(/\//g, "~1")}`;
}

function jsonValueType(value) {
  if (Array.isArray(value)) return "array";
  if (value === null) return "null";
  return typeof value;
}

function jsonNodeSummary(value) {
  const type = jsonValueType(value);
  if (type === "array") return `${value.length} item${value.length === 1 ? "" : "s"}`;
  if (type === "object") {
    const count = Object.keys(value).length;
    return `${count} key${count === 1 ? "" : "s"}`;
  }
  if (type === "string") return JSON.stringify(value);
  return String(value);
}

function renderJsonExplorerNode(key, value, path, depth = 0) {
  const type = jsonValueType(value);
  const isContainer = type === "object" || type === "array";
  const isExpanded = state.jsonExplorerExpandedPaths.has(path);
  const label = path ? key : "root";
  const children = isContainer
    ? Object.entries(value).map(([childKey, childValue]) => {
      return renderJsonExplorerNode(childKey, childValue, jsonExplorerPath(path, childKey), depth + 1);
    }).join("")
    : "";

  return `
    <div class="json-explorer-row" style="--depth: ${depth}">
      ${isContainer
        ? `<button type="button" class="json-toggle" data-json-explorer-path="${escapeHtml(path)}" aria-label="${isExpanded ? "Collapse" : "Expand"} ${escapeHtml(label)}">${isExpanded ? "-" : "+"}</button>`
        : `<span class="json-toggle-spacer"></span>`}
      <span class="json-key">${escapeHtml(label)}</span>
      <span class="json-type">${escapeHtml(type)}</span>
      <span class="json-summary">${escapeHtml(jsonNodeSummary(value))}</span>
    </div>
    ${isContainer && isExpanded ? `<div class="json-children">${children}</div>` : ""}
  `;
}

function renderJsonExplorerFromText(text) {
  const explorer = document.getElementById("catalog-json-explorer");
  if (!explorer) return;
  try {
    const parsed = JSON.parse(text || "{}");
    explorer.innerHTML = renderJsonExplorerNode("", parsed, "");
  } catch (error) {
    explorer.innerHTML = `
      <div class="json-explorer-empty">
        <strong>Invalid JSON</strong>
        <span>${escapeHtml(error.message)}</span>
      </div>
    `;
  }
}

function laggedMetricScenarioSnapshot(metric = {}) {
  const snapshot = {};
  if (Object.prototype.hasOwnProperty.call(metric, "openingValue")) {
    snapshot.openingValue = metric.openingValue;
  }
  return snapshot;
}

function normalizeScenarioMetric(definition, rawMetric = {}) {
  if (metricIsLagged(definition)) return laggedMetricScenarioSnapshot(rawMetric);
  const controlPoints = normalizeControlPointMap(definition, rawMetric);
  const normalized = {
    ...rawMetric,
    topDown: normalizeSeriesMap(definition, rawMetric.topDown),
    ...(Object.keys(controlPoints).length ? { controlPoints } : {}),
  };
  delete normalized.manualControlPoints;
  delete normalized.memberManualControlPoints;
  delete normalized.topDownTotal;
  return normalized;
}

function scenarioSnapshot(sourceScenario = scenario()) {
  const metrics = {};
  Object.entries(sourceScenario.metrics || {}).forEach(([id, metric]) => {
    const definition = metricDefinitions()[id];
    metrics[id] = metricIsLagged(definition)
      ? laggedMetricScenarioSnapshot(metric)
      : normalizeScenarioMetric(definition || {}, metric);
  });

  return {
    ...sourceScenario,
    metrics,
  };
}

function scenariosSnapshot() {
  return Object.fromEntries(Object.entries(scenarioCollection()).map(([id, item]) => [
    id,
    scenarioSnapshot(item),
  ]));
}

function uniqueAssumptionId() {
  let index = assumptions().length + 1;
  let id = `assumption-${index}`;
  const usedIds = new Set(assumptions().map(item => item.id));
  while (usedIds.has(id)) {
    index += 1;
    id = `assumption-${index}`;
  }
  return id;
}

function assumptionScopeForSelection(metricId = state.selectedMetricId) {
  const context = selectedDimensionContextForMetric(metricId);
  return {
    coordinate: context?.coordinate || {},
  };
}

function normalizeAssumption(rawAssumption) {
  const id = rawAssumption?.id || uniqueAssumptionId();
  const rawImpact = rawAssumption?.impact && typeof rawAssumption.impact === "object"
    ? rawAssumption.impact
    : {};
  const pct = Number(rawImpact.pct ?? 0);
  return {
    id,
    metricId: rawAssumption?.metricId || null,
    scope: rawAssumption?.scope && typeof rawAssumption.scope === "object"
      ? { coordinate: rawAssumption.scope.coordinate || {} }
      : { coordinate: {} },
    baselineRole: rawAssumption?.baselineRole || "bau",
    comparisonRole: rawAssumption?.comparisonRole || "target",
    years: Array.isArray(rawAssumption?.years) ? rawAssumption.years.map(Number).filter(Number.isFinite) : [...YEARS],
    claim: rawAssumption?.claim || "",
    driverType: rawAssumption?.driverType || "other",
    owner: rawAssumption?.owner || null,
    impact: {
      mode: rawImpact.mode || "pctOfDelta",
      pct: Number.isFinite(pct) ? pct : 0,
    },
    evidence: Array.isArray(rawAssumption?.evidence) ? rawAssumption.evidence : [],
    status: rawAssumption?.status || "draft",
    createdAt: rawAssumption?.createdAt || null,
    updatedAt: rawAssumption?.updatedAt || null,
  };
}

function assumptionScopeKey(scope = {}) {
  return coordinateKey(scope.coordinate || {});
}

function assumptionsForSelection(metricId = state.selectedMetricId) {
  const scopeKey = assumptionScopeKey(assumptionScopeForSelection(metricId));
  return assumptions().filter(item => (
    item.metricId === metricId
    && assumptionScopeKey(item.scope || {}) === scopeKey
  ));
}

function referenceMetricSeries(role, metricId, filters = {}) {
  const reference = referenceScenarios()[role];
  const definition = metricDefinitions()[metricId];
  const metric = reference?.snapshot?.metrics?.[metricId];
  if (!definition || !metric || metricIsLagged(definition)) return null;
  const seriesMap = normalizeSeriesMap(definition, metric.topDown);
  const entries = Object.entries(seriesMap).filter(([key]) => coordinateMatchesFilters(key, filters));
  if (!entries.length) return null;
  return aggregateMemberSeries(metricId, Object.fromEntries(entries));
}

function seriesDeltaSummary(left = [], right = []) {
  const deltaByYear = YEARS.map((_year, index) => Number(left?.[index] || 0) - Number(right?.[index] || 0));
  return {
    annual: deltaByYear,
    total: deltaByYear.reduce((sum, value) => sum + value, 0),
  };
}

function totalSeriesValue(series = []) {
  return YEARS.reduce((sum, _year, index) => sum + Number(series?.[index] || 0), 0);
}

function assumptionsForMetric(metricId) {
  return assumptions().filter(item => item.metricId === metricId);
}

function assumptionCoverage(assumptionItems = []) {
  const pct = assumptionItems.reduce((sum, item) => {
    if ((item.impact?.mode || "pctOfDelta") !== "pctOfDelta") return sum;
    return sum + Number(item.impact?.pct || 0);
  }, 0);
  return {
    pct,
    status: pct > 100.000001
      ? "over"
      : pct >= 99.999999
        ? "full"
        : pct > 0
          ? "partial"
          : "none",
  };
}

function explainedDeltaValue(totalDelta, coveragePct) {
  return Number(totalDelta || 0) * (Number(coveragePct || 0) / 100);
}

function comparisonMetricSummary(role, metricId) {
  const definition = metricDefinitions()[metricId];
  if (!definition) return null;
  const referenceSeries = referenceMetricSeries(role, metricId);
  if (!referenceSeries) return null;
  const workingSeries = metricTopDownSeries(metricId);
  const delta = seriesDeltaSummary(workingSeries, referenceSeries);
  const assumptionItems = assumptionsForMetric(metricId);
  const coverage = assumptionCoverage(assumptionItems);
  const absoluteTotalDelta = Math.abs(delta.total);
  const annualAbsDelta = delta.annual.reduce((sum, value) => sum + Math.abs(value), 0);
  const hasMaterialDelta = absoluteTotalDelta > 0.000001 || annualAbsDelta > 0.000001;

  return {
    metricId,
    label: definition.label,
    unit: definition.unit,
    workingTotal: totalSeriesValue(workingSeries),
    referenceTotal: totalSeriesValue(referenceSeries),
    totalDelta: delta.total,
    absoluteTotalDelta,
    annualAbsoluteDelta: annualAbsDelta,
    annualDelta: delta.annual,
    annualDeltaByYear: Object.fromEntries(YEARS.map((year, index) => [year, delta.annual[index]])),
    assumptionCount: assumptionItems.length,
    assumptionIds: assumptionItems.map(item => item.id),
    coveragePct: coverage.pct,
    coverageStatus: coverage.status,
    explainedTotalDelta: explainedDeltaValue(delta.total, coverage.pct),
    unexplainedTotalDelta: hasMaterialDelta ? delta.total - explainedDeltaValue(delta.total, coverage.pct) : 0,
    hasExplanation: !hasMaterialDelta || coverage.pct >= 99.999999,
  };
}

function comparisonSummaryForRole(role) {
  const reference = referenceScenarios()[role];
  if (!reference?.snapshot) {
    return {
      available: false,
      role,
      referenceId: null,
      metrics: {},
      largestDeltas: [],
      unexplainedDeltas: [],
    };
  }

  const metricSummaries = Object.fromEntries(Object.keys(metricDefinitions())
    .map(metricId => [metricId, comparisonMetricSummary(role, metricId)])
    .filter(([_metricId, summary]) => summary));
  const sortedDeltas = Object.values(metricSummaries)
    .filter(summary => summary.absoluteTotalDelta > 0.000001 || summary.annualAbsoluteDelta > 0.000001)
    .sort((left, right) => {
      const rightMagnitude = Math.max(right.absoluteTotalDelta, right.annualAbsoluteDelta);
      const leftMagnitude = Math.max(left.absoluteTotalDelta, left.annualAbsoluteDelta);
      return rightMagnitude - leftMagnitude;
    });

  return {
    available: true,
    role,
    referenceId: reference.id,
    sourceScenarioId: reference.sourceScenarioId,
    sourceScenarioName: reference.sourceScenarioName,
    assignedAt: reference.assignedAt,
    metrics: metricSummaries,
    largestDeltas: sortedDeltas.slice(0, 10).map(summary => ({
      metricId: summary.metricId,
      label: summary.label,
      totalDelta: summary.totalDelta,
      annualAbsoluteDelta: summary.annualAbsoluteDelta,
      assumptionCount: summary.assumptionCount,
      coveragePct: summary.coveragePct,
      coverageStatus: summary.coverageStatus,
      explainedTotalDelta: summary.explainedTotalDelta,
      unexplainedTotalDelta: summary.unexplainedTotalDelta,
      hasExplanation: summary.hasExplanation,
    })),
    unexplainedDeltas: sortedDeltas
      .filter(summary => !summary.hasExplanation)
      .map(summary => ({
        metricId: summary.metricId,
        label: summary.label,
        totalDelta: summary.totalDelta,
        annualAbsoluteDelta: summary.annualAbsoluteDelta,
      })),
  };
}

function comparisonSummary() {
  return {
    workingVsBau: comparisonSummaryForRole("bau"),
    workingVsTarget: comparisonSummaryForRole("target"),
  };
}

function deltaReviewRows(summary) {
  if (!summary.available) {
    return `
      <div class="delta-review-empty">
        <strong>${summary.role.toUpperCase()} not set</strong>
        <span>Set ${summary.role.toUpperCase()} from a reconciled scenario to review deltas.</span>
      </div>
    `;
  }
  const rows = summary.largestDeltas.slice(0, 5);
  if (!rows.length) {
    return `
      <div class="delta-review-empty">
        <strong>No deltas</strong>
        <span>Working matches ${summary.role.toUpperCase()} across available metrics.</span>
      </div>
    `;
  }
  return rows.map(row => {
    const definition = metricDefinitions()[row.metricId] || {};
    return `
      <button class="delta-review-row ${row.hasExplanation ? "" : "unexplained"}" data-delta-metric-id="${escapeHtml(row.metricId)}" type="button">
        <span>
          <strong>${escapeHtml(row.label)}</strong>
          <small>${row.assumptionCount} assumption${row.assumptionCount === 1 ? "" : "s"} | ${trimFixed(row.coveragePct || 0, 1)}% covered${row.coverageStatus === "over" ? " | over-explained" : row.hasExplanation ? "" : " | unexplained"}</small>
        </span>
        <b>${formatMetricValue(definition, row.totalDelta, { precise: true })}</b>
      </button>
    `;
  }).join("");
}

function renderDeltaReviewPanel() {
  const panel = document.getElementById("delta-review-panel");
  if (!panel) return;
  const summary = comparisonSummary();
  const reconciliation = activeScenarioReconciliationSummary();
  panel.innerHTML = `
    <div class="delta-review-heading">
      <div>
        <h2>Delta Review</h2>
        <p>Largest changes from reference scenarios and assumptions attached to them.</p>
      </div>
      <span class="scenario-role-status ${reconciliation.matched ? "matched" : "open"}">
        ${reconciliation.matched ? "Top-down / bottom-up matched" : `${reconciliation.openMetricCount} reconciliation mismatch${reconciliation.openMetricCount === 1 ? "" : "es"}`}
      </span>
    </div>
    <div class="delta-review-grid">
      <section>
        <div class="delta-review-section-title">
          <strong>Working vs BAU</strong>
          <span>${summary.workingVsBau.available ? escapeHtml(summary.workingVsBau.sourceScenarioName || "BAU") : "Unset"}</span>
        </div>
        <div class="delta-review-list">${deltaReviewRows(summary.workingVsBau)}</div>
      </section>
      <section>
        <div class="delta-review-section-title">
          <strong>Working vs Target</strong>
          <span>${summary.workingVsTarget.available ? escapeHtml(summary.workingVsTarget.sourceScenarioName || "Target") : "Unset"}</span>
        </div>
        <div class="delta-review-list">${deltaReviewRows(summary.workingVsTarget)}</div>
      </section>
    </div>
  `;
}

function disposeMetricWorkspaceCharts() {
  metricWorkspaceCharts.forEach(item => item?.dispose?.());
  metricWorkspaceCharts = [];
}

function disposeWorkspaceNapkinChart() {
  if (workspaceNapkinReferenceStyleFrame) {
    cancelAnimationFrame(workspaceNapkinReferenceStyleFrame);
    workspaceNapkinReferenceStyleFrame = null;
  }
  if (workspaceNapkinChart?.chart) {
    workspaceNapkinChart.chart.dispose();
  }
  workspaceNapkinChart = null;
  workspaceNapkinMetricId = null;
  workspaceNapkinHasPendingCommit = false;
}

function disposeWorkspaceDimensionMixChart() {
  if (workspaceDimensionMixChart?.chart) {
    workspaceDimensionMixChart.chart.dispose();
  }
  workspaceDimensionMixChart = null;
  workspaceDimensionMixMetricId = null;
  workspaceDimensionMixHasPendingCommit = false;
}

function disposeWorkspaceDimensionMemberCharts() {
  workspaceDimensionMemberCharts.forEach(item => item?.chart?.dispose?.());
  workspaceDimensionMemberCharts = new Map();
  workspaceDimensionMemberChartSyncing = new Set();
  workspaceDimensionMemberChartPendingCommits = new Set();
}

function disposeWorkspaceChildNapkinCharts() {
  workspaceChildNapkinCharts.forEach(item => item?.chart?.dispose?.());
  workspaceChildNapkinCharts = new Map();
  workspaceChildNapkinChartSyncing = new Set();
  workspaceChildNapkinChartPendingCommits = new Set();
}

function cancelWorkspaceLiveRefresh() {
  if (workspaceLiveRefreshFrame) {
    cancelAnimationFrame(workspaceLiveRefreshFrame);
    workspaceLiveRefreshFrame = null;
  }
  pendingWorkspaceLiveRefreshOptions = null;
}

function workspaceChartOption(definition, seriesList, { compact = false } = {}) {
  const values = seriesList.flatMap(item => item.data || []);
  const maxValue = Math.max(1, ...values.map(value => Math.abs(Number(value || 0))));
  return {
    animation: false,
    tooltip: {
      trigger: "axis",
      valueFormatter: value => formatMetricValue(definition, value, { precise: true }),
    },
    legend: compact ? { show: false } : { top: 0 },
    grid: { left: 10, right: 10, top: compact ? 10 : 34, bottom: 22, containLabel: true },
    xAxis: {
      type: "category",
      data: YEARS.map(String),
      axisTick: { alignWithLabel: true },
    },
    yAxis: {
      type: "value",
      min: value => Math.min(0, value.min),
      max: value => Math.max(maxValue, value.max),
      axisLabel: { formatter: value => formatMetricValue(definition, value) },
    },
    series: seriesList.map(item => ({
      name: item.name,
      type: "line",
      data: item.data,
      symbolSize: compact ? 4 : 6,
      lineStyle: { width: item.width || (compact ? 2 : 2.5), type: item.type || "solid" },
      itemStyle: { color: item.color },
    })),
  };
}

function renderWorkspaceChart(elementId, definition, seriesList, options = {}) {
  const element = document.getElementById(elementId);
  if (!element || !window.echarts) return;
  const instance = echarts.init(element);
  instance.setOption(workspaceChartOption(definition, seriesList, options), true);
  metricWorkspaceCharts.push(instance);
}

function workspaceAvailableModes(definition) {
  if (!definition) return [];
  const modes = [];
  const isWeightedAverage = definition.bottomUp?.type === "weightedAverage";
  if (isWeightedAverage) modes.push({ id: "weightedAverage", label: "Weighted Avg" });
  if (childIds(definition.id).length) modes.push({ id: "children", label: "Children" });
  if (!isWeightedAverage && dimensionMixContext(definition)) modes.push({ id: "dimensions", label: "Dimensions" });
  if (["cohortMatrix", "cohortMatrixFromStartsAndAgeYoy"].includes(definition.bottomUp?.type)) {
    modes.push({ id: "cohort", label: "Cohort" });
  }
  return modes;
}

function activeWorkspaceMode(definition) {
  const modes = workspaceAvailableModes(definition);
  if (!modes.some(mode => mode.id === state.metricWorkspaceMode)) {
    state.metricWorkspaceMode = modes[0]?.id || "";
  }
  return state.metricWorkspaceMode;
}

function workspaceModeTabs(definition) {
  const modes = workspaceAvailableModes(definition);
  if (!modes.length) return "";
  const activeMode = activeWorkspaceMode(definition);
  return `
    <div class="metric-workspace-lenses">
      <span>Explain by</span>
      <div class="metric-workspace-tabs">
        ${modes.map(mode => `
          <button class="${activeMode === mode.id ? "active" : ""}" type="button" data-workspace-mode="${mode.id}">
            ${mode.label}
          </button>
        `).join("")}
      </div>
    </div>
  `;
}

function workspaceTotalModeHtml(definition, bottomUp) {
  const topDown = activeTopDownSeries(definition);
  const summary = bottomUp ? seriesDeltaSummary(topDown, bottomUp) : null;
  const breadcrumb = workspaceDimensionBreadcrumbHtml(definition);
  const summaryCards = `
    <div class="metric-workspace-summary-grid">
      <div>
        <span>Top-Down Total</span>
        <strong data-workspace-summary-value="top">${formatMetricValue(definition, totalSeriesValue(topDown), { precise: true })}</strong>
      </div>
      <div>
        <span>Bottom-Up Total</span>
        <strong data-workspace-summary-value="bottom">${bottomUp ? formatMetricValue(definition, totalSeriesValue(bottomUp), { precise: true }) : "-"}</strong>
      </div>
      <div data-workspace-summary-card="gap" class="${summary && Math.abs(summary.total) > (definition.reconciliation?.tolerance || 1) ? "open" : "matched"}">
        <span>Gap</span>
        <strong data-workspace-summary-value="gap">${summary ? formatMetricValue(definition, summary.total, { precise: true }) : "-"}</strong>
      </div>
    </div>
  `;
  if (metricIsLagged(definition)) {
    return `
      <section class="metric-workspace-selected">
        ${breadcrumb}
        <h3>${escapeHtml(definition.label)}: Derived Lagged Series</h3>
        ${summaryCards}
        <div id="metric-workspace-selected-chart" class="metric-workspace-chart"></div>
      </section>
    `;
  }
  return `
    <section class="metric-workspace-selected">
      ${breadcrumb}
      <div class="metric-workspace-section-title">
        <div>
          <h3>${escapeHtml(definition.label)} Top-Down</h3>
          <span>${bottomUp ? "Editable top-down line. Use the gap summary to compare against the current bottom-up definition." : "Editable independent top-down line."}</span>
        </div>
      </div>
      ${summaryCards}
      <div id="metric-workspace-napkin-chart" class="metric-workspace-napkin-chart"></div>
    </section>
  `;
}

function workspaceChildrenModeHtml(definition, children) {
  const canFitChildren = definition.bottomUp?.type === "sum" && children.length > 0;
  return `
    <section class="metric-workspace-children">
      <div class="metric-workspace-section-title">
        <div>
          <h3>${escapeHtml(definition.label)} Children</h3>
          <span>${children.length ? "Edit or inspect the direct formula inputs that explain the selected total above." : "No children. Use the formula builder below to define this metric."}</span>
        </div>
        <div class="metric-workspace-inline-actions">
          <button type="button" data-workspace-set-parent="${escapeHtml(definition.id)}" ${metricDirectBottomUpSeries(definition.id) ? "" : "disabled"}>Set Parent to Children</button>
          <button type="button" data-workspace-fit-children="${escapeHtml(definition.id)}" ${canFitChildren ? "" : "disabled"}>Fit Children to Parent</button>
        </div>
      </div>
      <div class="metric-workspace-child-grid">
        ${children.length ? children.map((childId, index) => {
          const child = metricDefinitions()[childId];
          const bottomUp = metricDirectBottomUpSeries(childId);
          const lockKey = workspaceChildLockKey(definition.id, childId);
          const locked = workspaceIsLocked(lockKey);
          return `
            <div class="metric-workspace-child-card metric-workspace-child-tile ${locked ? "locked" : ""}">
              <div class="metric-workspace-card-heading">
                <div>
                  <span>${escapeHtml(child?.label || childId)}</span>
                  <strong data-workspace-child-value="${escapeHtml(childId)}">${formatMetricValue(child, totalSeriesValue(metricTopDownSeries(childId)), { precise: true })}</strong>
                  <small>${locked ? "Locked" : bottomUp ? "Top-down with bottom-up reference" : "Manual top-down series"}</small>
                </div>
                <div class="metric-workspace-card-actions">
                  ${workspaceLockButtonHtml(lockKey)}
                  <button type="button" data-metric-id="${escapeHtml(childId)}">Open metric</button>
                </div>
              </div>
              <div id="metric-workspace-child-chart-${index}" class="metric-workspace-child-chart"></div>
            </div>
          `;
        }).join("") : `
          <div class="metric-workspace-empty">
            <strong>Leaf metric</strong>
            <span>Use the total mode and assumptions to shape and defend this value.</span>
          </div>
        `}
      </div>
    </section>
  `;
}

function formulaInputSeriesForFilters(metricId, filters = {}) {
  if (!metricId || !metricDefinitions()[metricId] || !inputCanUseFilters(metricId, filters)) return null;
  return metricTopDownContextSeries(metricId, { coordinate: filters });
}

function weightedAverageWeightedValueSeries(definition, filters = {}) {
  const series = YEARS.map((_year, index) => {
    const totals = weightedAverageTotalsAtIndex(definition, index, filters);
    return totals ? totals.value : null;
  });
  return series.some(value => value === null || value === undefined) ? null : series;
}

function weightedAverageSeriesForFilters(definition, filters = {}) {
  const { valueMetricId, weightMetricId } = weightedAverageConfig(definition);
  const weightSeries = formulaInputSeriesForFilters(weightMetricId, filters);
  const valueSeries = valueMetricId
    ? formulaInputSeriesForFilters(valueMetricId, filters)
    : weightedAverageWeightedValueSeries(definition, filters);
  if (!valueSeries || !weightSeries) return null;
  const series = YEARS.map((_year, index) => {
    const weight = Number(weightSeries[index] || 0);
    return weight ? Number(valueSeries[index] || 0) / weight : null;
  });
  return series.some(value => value === null || value === undefined) ? null : series;
}

function weightedAverageSummaryValue(definition, filters = {}) {
  const { valueMetricId, weightMetricId } = weightedAverageConfig(definition);
  const weightSeries = formulaInputSeriesForFilters(weightMetricId, filters);
  const valueSeries = valueMetricId
    ? formulaInputSeriesForFilters(valueMetricId, filters)
    : weightedAverageWeightedValueSeries(definition, filters);
  if (!valueSeries || !weightSeries) return null;
  const totalWeight = totalSeriesValue(weightSeries);
  if (!totalWeight) return null;
  return totalSeriesValue(valueSeries) / totalWeight;
}

function workspaceWeightedAverageRows(definition) {
  const mixContext = dimensionMixContext(definition);
  const baseFilters = mixContext?.baseFilters || selectedDimensionContextForMetric(definition.id)?.coordinate || {};
  const baseRow = {
    id: "total",
    label: mixContext?.baseKey ? coordinateLabelForMetric(definition.id, mixContext.baseKey) : `${definition.label} Total`,
    filters: baseFilters,
  };
  if (!mixContext) return [baseRow];
  return [
    baseRow,
    ...dimensionGroupsForContext(mixContext).map(group => ({
      id: group.id,
      label: group.label,
      filters: { ...mixContext.baseFilters, [mixContext.dimensionId]: group.memberIds },
      members: group.members,
    })),
  ];
}

function workspaceWeightedAverageModeHtml(definition) {
  const { valueMetricId, weightMetricId } = weightedAverageConfig(definition);
  const valueDefinition = metricDefinitions()[valueMetricId];
  const weightDefinition = metricDefinitions()[weightMetricId];
  const valueLabel = valueDefinition?.label || `Implied ${definition.label} × ${weightDefinition?.label || "Weight"}`;
  const weightLabel = weightDefinition?.label || "Weight";
  const weightedSeries = activeWorkspaceBottomUpSeries(definition);
  const activeFilters = selectedDimensionContextForMetric(definition.id)?.coordinate || {};
  const valueSeries = valueMetricId
    ? formulaInputSeriesForFilters(valueMetricId, activeFilters)
    : weightedAverageWeightedValueSeries(definition, activeFilters);
  const weightSeries = formulaInputSeriesForFilters(weightMetricId, selectedDimensionContextForMetric(definition.id)?.coordinate || {});
  const rows = workspaceWeightedAverageRows(definition);
  const mixContext = dimensionMixContext(definition);
  const dimensionTitle = mixContext
    ? `${mixContext.dimensionLabel || "Dimension"} Weighted Mix`
    : "Weighted Mix";
  return `
    <section class="metric-workspace-children metric-workspace-weighted-average">
      <div class="metric-workspace-section-title">
        <div>
          <h3>Weighted Average View</h3>
          <span>${escapeHtml(definition.label)} is explained by ${escapeHtml(valueLabel)} divided by ${escapeHtml(weightLabel)}.</span>
        </div>
      </div>
      <div class="weighted-average-equation">
        <div>
          <span>Weighted Avg</span>
          <strong>${escapeHtml(definition.label)}</strong>
        </div>
        <i>=</i>
        <div>
          <span>Weighted Value</span>
          <strong>${escapeHtml(valueLabel)}</strong>
        </div>
        <i>/</i>
        <div>
          <span>Weight</span>
          <strong>${escapeHtml(weightLabel)}</strong>
        </div>
      </div>
      <div class="weighted-average-grid">
        <div class="weighted-average-card rate-card">
          <div class="metric-workspace-card-heading">
            <div>
              <span>${escapeHtml(definition.label)}</span>
              <strong>${weightedSeries ? formatMetricValue(definition, weightedAverageSummaryValue(definition, selectedDimensionContextForMetric(definition.id)?.coordinate || {}), { precise: true }) : "-"}</strong>
              <small>Weighted average rate</small>
            </div>
          </div>
          <div id="metric-workspace-weighted-rate-chart" class="weighted-average-chart"></div>
        </div>
        <div class="weighted-average-card">
          <div class="metric-workspace-card-heading">
            <div>
              <span>${escapeHtml(valueLabel)}</span>
              <strong>${valueSeries ? formatMetricValue(valueDefinition || definition, totalSeriesValue(valueSeries), { precise: true }) : "-"}</strong>
              <small>${valueMetricId ? "Numerator / weighted value" : "Derived from rate × weight"}</small>
            </div>
          </div>
          <div id="metric-workspace-weighted-value-chart" class="weighted-average-chart"></div>
        </div>
        <div class="weighted-average-card">
          <div class="metric-workspace-card-heading">
            <div>
              <span>${escapeHtml(weightLabel)}</span>
              <strong>${weightSeries ? formatMetricValue(weightDefinition, totalSeriesValue(weightSeries), { precise: true }) : "-"}</strong>
              <small>Denominator / weight</small>
            </div>
          </div>
          <div id="metric-workspace-weighted-weight-chart" class="weighted-average-chart"></div>
        </div>
      </div>
      <div class="metric-workspace-section-title weighted-average-mix-heading">
        <div>
          <h3>${escapeHtml(dimensionTitle)}</h3>
          <span>${mixContext ? "Dimension slices are evaluated through weight share, weighted value, and contribution to the rate." : "This metric has no dimensional slices yet."}</span>
        </div>
      </div>
      ${mixContext ? `
        <div class="metric-workspace-dimension-mix-layout weighted-average-dimension-layout">
          <div class="metric-workspace-dimension-chart-tile">
            <div class="metric-workspace-card-heading">
              <div>
                <span>Editable Rate Slices</span>
                <strong>${escapeHtml(mixContext.dimensionLabel || "Dimension")}</strong>
                <small>Edit grouped dimension rates. The weighted average updates from each slice's weight.</small>
              </div>
            </div>
            <div class="metric-workspace-dimension-grid weighted-average-edit-grid">
              ${dimensionGroupsForContext(mixContext).map((group, index) => {
                const series = dimensionGroupSeries(mixContext, group);
                const drillCoordinate = { ...mixContext.baseFilters, [mixContext.dimensionId]: group.memberIds };
                const lockKey = workspaceDimensionGroupLockKey(mixContext, group.id);
                const locked = workspaceIsLocked(lockKey);
                return `
                  <div class="metric-workspace-child-card metric-workspace-dimension-card ${locked ? "locked" : ""}">
                    <div class="metric-workspace-card-heading">
                      <div>
                        <span>${escapeHtml(group.label)}</span>
                        <strong data-workspace-dimension-group-value="${escapeHtml(group.id)}">${formatMetricValue(definition, weightedAverageSummaryValue(definition, drillCoordinate) ?? totalSeriesValue(series), { precise: true })}</strong>
                        <small>${locked ? "Locked" : group.members.map(member => member.label).join(", ") || "No members"}</small>
                      </div>
                      <div class="metric-workspace-card-actions">
                        ${workspaceLockButtonHtml(lockKey, { disabled: !group.members.length })}
                        <button type="button" ${group.members.length ? "" : "disabled"} data-workspace-coordinate-metric="${escapeHtml(definition.id)}" data-workspace-coordinate-json="${escapeHtml(JSON.stringify(drillCoordinate))}">Open group</button>
                      </div>
                    </div>
                    <div id="metric-workspace-dimension-chart-${index}" class="metric-workspace-child-chart metric-workspace-member-napkin"></div>
                  </div>
                `;
              }).join("")}
            </div>
          </div>
          ${workspaceDimensionGroupingPanelHtml(mixContext)}
        </div>
      ` : ""}
      <div class="weighted-average-table-wrap">
        <table class="weighted-average-table">
          <thead>
            <tr>
              <th>Slice</th>
              <th>${escapeHtml(weightLabel)}</th>
              <th>${escapeHtml(valueLabel)}</th>
              <th>${escapeHtml(definition.label)}</th>
              <th>Weight Share</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(row => {
              const rate = weightedAverageSeriesForFilters(definition, row.filters);
              const value = valueMetricId
                ? formulaInputSeriesForFilters(valueMetricId, row.filters)
                : weightedAverageWeightedValueSeries(definition, row.filters);
              const weight = formulaInputSeriesForFilters(weightMetricId, row.filters);
              const totalWeight = weight ? totalSeriesValue(weight) : null;
              const parentWeight = weightSeries ? totalSeriesValue(weightSeries) : null;
              return `
                <tr>
                  <th>${escapeHtml(row.label)}</th>
                  <td>${weight ? formatMetricValue(weightDefinition, totalWeight, { precise: true }) : "-"}</td>
                  <td>${value ? formatMetricValue(valueDefinition || definition, totalSeriesValue(value), { precise: true }) : "-"}</td>
                  <td>${rate ? formatMetricValue(definition, weightedAverageSummaryValue(definition, row.filters), { precise: true }) : "-"}</td>
                  <td>${parentWeight ? `${trimFixed((Number(totalWeight || 0) / parentWeight) * 100, 1)}%` : "-"}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function workspaceDimensionsModeHtml(definition) {
  const mixContext = dimensionMixContext(definition);
  if (!mixContext) return workspaceTotalModeHtml(definition, activeWorkspaceBottomUpSeries(definition));
  const { dimensionLabel } = mixContext;
  const groups = dimensionGroupsForContext(mixContext);
  const contextLabel = mixContext.baseKey ? ` inside ${coordinateLabelForMetric(definition.id, mixContext.baseKey)}` : "";
  const hasPctTotalView = groups.length >= 2;
  const activeChartView = hasPctTotalView && state.dimensionMixView === "pctTotal" ? "pctTotal" : "total";
  const anyGroupHasBottomUp = groups.some(group => dimensionGroupBottomUpSeries(mixContext, group));
  const bottomUpLensText = definition.bottomUp && !anyGroupHasBottomUp
    ? `This dimension is allocation-only for bottom-up because ${dimensionLabel || "this dimension"} is not present in the formula inputs.`
    : groups.length >= 2
      ? `Group members, edit group mix as pct total, or edit grouped value slices${contextLabel}.`
      : "Add at least two members to edit pct total.";
  return `
    <section class="metric-workspace-children">
      <div class="metric-workspace-section-title">
        <div>
          <h3>${escapeHtml(dimensionLabel || "Dimension")} Mix</h3>
          <span>${escapeHtml(bottomUpLensText)}</span>
        </div>
      </div>
      <div class="metric-workspace-dimension-mix-layout">
        <div class="metric-workspace-dimension-chart-tile">
          <div class="metric-workspace-card-heading">
            <div>
              <span>${activeChartView === "pctTotal" ? "% Total" : "Total Lines"}</span>
              <strong>${escapeHtml(dimensionLabel || "Dimension")}</strong>
              <small>${activeChartView === "pctTotal" ? "Edit group share of total." : "Inspect grouped and raw member totals."}</small>
            </div>
            <div class="metric-workspace-chart-view-toggle">
              <button type="button" class="${activeChartView === "total" ? "active" : ""}" data-dimension-mix-view="total">Total</button>
              <button type="button" class="${activeChartView === "pctTotal" ? "active" : ""}" data-dimension-mix-view="pctTotal" ${hasPctTotalView ? "" : "disabled"}>% Total</button>
            </div>
          </div>
          ${activeChartView === "pctTotal" ? `
            <div id="metric-workspace-dimension-mix-chart" class="metric-workspace-dimension-mix-chart"></div>
          ` : `
            <div class="metric-workspace-group-line-controls">
              <div>
                <button type="button" class="${state.dimensionGroupLinesShowGrouped ? "active" : ""}" data-dimension-group-line-layer="grouped">Grouped</button>
                <button type="button" class="${state.dimensionGroupLinesShowUngrouped ? "active" : ""}" data-dimension-group-line-layer="ungrouped">Ungrouped</button>
              </div>
              <div>
                <button type="button" class="${state.dimensionGroupLinesChartType === "line" ? "active" : ""}" data-dimension-group-line-chart-type="line">Line</button>
                <button type="button" class="${state.dimensionGroupLinesChartType === "bar" ? "active" : ""}" data-dimension-group-line-chart-type="bar">Bar</button>
              </div>
            </div>
            <div id="metric-workspace-dimension-group-lines-chart" class="metric-workspace-group-lines-chart"></div>
          `}
        </div>
        ${workspaceDimensionGroupingPanelHtml(mixContext)}
      </div>
      <div class="metric-workspace-dimension-grid">
        ${groups.map((group, index) => {
          const series = dimensionGroupSeries(mixContext, group);
          const drillCoordinate = { ...mixContext.baseFilters, [mixContext.dimensionId]: group.memberIds };
          const lockKey = workspaceDimensionGroupLockKey(mixContext, group.id);
          const locked = workspaceIsLocked(lockKey);
          return `
            <div class="metric-workspace-child-card metric-workspace-dimension-card ${locked ? "locked" : ""}">
              <div class="metric-workspace-card-heading">
                <div>
                  <span>${escapeHtml(group.label)}</span>
                  <strong data-workspace-dimension-group-value="${escapeHtml(group.id)}">${formatMetricValue(definition, totalSeriesValue(series), { precise: true })}</strong>
                  <small>${locked ? "Locked" : group.members.map(member => member.label).join(", ") || "No members"}</small>
                </div>
                <div class="metric-workspace-card-actions">
                  ${workspaceLockButtonHtml(lockKey, { disabled: !group.members.length })}
                  <button type="button" ${group.members.length ? "" : "disabled"} data-workspace-coordinate-metric="${escapeHtml(definition.id)}" data-workspace-coordinate-json="${escapeHtml(JSON.stringify(drillCoordinate))}">Open group</button>
                </div>
              </div>
              <div id="metric-workspace-dimension-chart-${index}" class="metric-workspace-child-chart metric-workspace-member-napkin"></div>
            </div>
          `;
        }).join("")}
      </div>
    </section>
  `;
}

function workspaceDimensionGroupingPanelHtml(mixContext) {
  const grouping = dimensionGroupingForContext(mixContext);
  const activeTool = state.dimensionGroupTool === "blue" ? "blue" : "red";
  return `
    <div class="metric-workspace-grouping-panel">
      <div class="metric-workspace-grouping-header">
        <div>
          <strong>Dynamic Groups</strong>
          <span>Assign members into editable rollups.</span>
        </div>
      </div>
      <div class="metric-workspace-group-tools">
        <button type="button" class="red ${activeTool === "red" ? "active" : ""}" data-dimension-group-tool="red">Red</button>
        <button type="button" class="blue ${activeTool === "blue" ? "active" : ""}" data-dimension-group-tool="blue">Blue</button>
      </div>
      <div class="metric-workspace-group-pills">
        ${mixContext.members.map(member => {
          const groupId = grouping.red.includes(member.id) ? "red" : grouping.blue.includes(member.id) ? "blue" : "other";
          return `
            <button type="button" class="${escapeHtml(groupId)}" data-dimension-group-member="${escapeHtml(member.id)}">
              ${escapeHtml(member.label)}
            </button>
          `;
        }).join("")}
      </div>
    </div>
  `;
}

function workspaceDimensionBreadcrumbHtml(definition) {
  const dimensionIds = metricDimensionIds(definition.id);
  if (!dimensionIds.length) return "";
  const context = selectedDimensionContextForMetric(definition.id);
  const coordinate = context?.coordinate || {};
  const crumbs = [{
    label: `${definition.label} Total`,
    coordinate: {},
    active: !Object.keys(coordinate).length,
  }];

  const accumulated = {};
  dimensionIds.forEach((dimensionId, index) => {
    const memberIds = filterValueList(coordinate[dimensionId]);
    if (!memberIds.length) return;
    accumulated[dimensionId] = memberIds.length === 1 ? memberIds[0] : memberIds;
    const label = filterLabelForMetric(definition.id, { [dimensionId]: accumulated[dimensionId] });
    crumbs.push({
      label,
      coordinate: { ...accumulated },
      active: index === Object.keys(coordinate).length - 1,
    });
  });

  const nextDimensionId = dimensionIds.find(dimensionId => !filterValueList(coordinate[dimensionId]).length);
  return `
    <div class="metric-workspace-breadcrumb">
      <span>Dimension Path</span>
      <div>
        ${crumbs.map((crumb, index) => `
          <button class="${crumb.active ? "active" : ""}" type="button" data-workspace-coordinate-metric="${escapeHtml(definition.id)}" data-workspace-coordinate-json="${escapeHtml(JSON.stringify(crumb.coordinate))}">
            ${escapeHtml(crumb.label)}
          </button>
          ${index < crumbs.length - 1 ? `<i>/</i>` : ""}
        `).join("")}
        ${nextDimensionId ? `<em>Next: ${escapeHtml(metricDimensionLevelLabel(definition.id, nextDimensionId))}</em>` : `<em>Deepest level</em>`}
      </div>
    </div>
  `;
}

function cohortMatrixPreviewMarkup(definition) {
  const matrixEntries = cohortPreviewEntries(definition);
  const displayDefinition = metricDefinitions()[definition?.id] || definition;
  if (!matrixEntries.length || !matrixEntries.some(([_key, matrix]) => Object.keys(matrix || {}).length)) {
    return `
      <div class="empty-state">
        <strong>No matrix yet</strong>
        <span>Choose a starting source or enter starting cohort values, then apply the matrix inputs.</span>
      </div>
    `;
  }

  const dimensionIds = cohortPreviewDimensionIds(definition, matrixEntries.map(([key]) => key));
  const hasCoordinateRows = dimensionIds.length > 0;
  const totals = YEARS.map(year => matrixEntries.reduce((sum, [_key, matrix]) => {
    return sum + sortedCohortMatrixRowKeys(matrix).reduce((rowSum, rowKey) => rowSum + cohortValueForYear(matrix[rowKey], year), 0);
  }, 0));
  const showCoordinateSubtotals = hasCoordinateRows && matrixEntries.length > 1;

  const rowsHtml = matrixEntries.map(([key, matrix]) => {
    const coordinate = parseCoordinateKey(key);
    const rowKeys = sortedCohortMatrixRowKeys(matrix);
    const subtotal = YEARS.map(year => rowKeys.reduce((sum, rowKey) => sum + cohortValueForYear(matrix[rowKey], year), 0));
    return `
      ${rowKeys.map(rowKey => `
        <tr>
          ${dimensionIds.map(dimensionId => `<td class="dimension-cell">${escapeHtml(coordinateMemberLabel(dimensionId, coordinate[dimensionId]))}</td>`).join("")}
          <td>${escapeHtml(cohortMatrixRowLabel(rowKey))}</td>
          ${YEARS.map(year => {
            const applies = cohortMatrixCellApplies(rowKey, year);
            const value = cohortValueForYear(matrix[rowKey], year);
            return `<td class="${applies ? "" : "not-applicable"}">${applies ? formatMetricValue(displayDefinition, value, { precise: true }) : "-"}</td>`;
          }).join("")}
        </tr>
      `).join("")}
      ${showCoordinateSubtotals ? `
        <tr class="coordinate-rollup-row">
          <th colspan="${dimensionIds.length + 1}">${escapeHtml(coordinateLabelForMetric(definition.id, key))}</th>
          ${subtotal.map(value => `<td>${formatMetricValue(displayDefinition, value, { precise: true })}</td>`).join("")}
        </tr>
      ` : ""}
    `;
  }).join("");

  return `
    <div class="cohort-matrix-preview">
      <table class="cohort-matrix-table">
        <thead>
          <tr>
            ${dimensionIds.map((dimensionId, index) => `<th>${escapeHtml(metricDimensionLevelLabel(definition.id, dimensionId) || `Level ${index + 1}`)}</th>`).join("")}
            <th>Cohort</th>
            ${YEARS.map(year => `<th>${year}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
          <tr class="rollup-row">
            <th colspan="${dimensionIds.length + 1}">Total</th>
            ${totals.map(value => `<td>${formatMetricValue(displayDefinition, value, { precise: true })}</td>`).join("")}
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

function workspaceCohortModeHtml(definition) {
  return `
    <section class="metric-workspace-children">
      <div class="metric-workspace-section-title">
        <div>
          <h3>Cohort Matrix</h3>
          <span>Preview of cohort-year cells. The active cohort controls are directly below in Metric Detail while we finish moving this editor.</span>
        </div>
        <div class="metric-workspace-inline-actions">
          <button type="button" data-focus-cohort-builder>Go to Cohort Builder</button>
        </div>
      </div>
      ${cohortMatrixPreviewMarkup(definition)}
    </section>
  `;
}

function setMetricControlPointsFromSeries(metricId, series) {
  const metric = ensureMetricScenario(metricId);
  if (!metric.controlPoints) metric.controlPoints = {};
  metric.controlPoints[TOTAL_COORDINATE_KEY] = YEARS.map((year, index) => [year, Number(series[index] || 0)]);
  setMetricTopDownSeries(metricId, series);
}

function fitChildrenToParent(metricId) {
  const definition = metricDefinitions()[metricId];
  const children = metricInputIds(definition);
  if (definition?.bottomUp?.type !== "sum" || !children.length) {
    setSourceStatus("Fit Children to Parent is currently available for sum formulas.", "error");
    return;
  }
  const parentSeries = metricTopDownSeries(metricId);
  const childSeriesById = Object.fromEntries(children.map(childId => [childId, metricTopDownSeries(childId)]));
  const lockedChildren = children.filter(childId => workspaceIsLocked(workspaceChildLockKey(metricId, childId)));
  const editableChildren = children.filter(childId => !workspaceIsLocked(workspaceChildLockKey(metricId, childId)));
  if (!editableChildren.length) {
    setSourceStatus("All children are locked. Unlock at least one child before fitting children to the parent.", "error");
    return;
  }
  const sharesByYear = YEARS.map((_year, index) => {
    const total = editableChildren.reduce((sum, childId) => sum + Number(childSeriesById[childId]?.[index] || 0), 0);
    if (total) {
      return Object.fromEntries(editableChildren.map(childId => [childId, Number(childSeriesById[childId]?.[index] || 0) / total]));
    }
    const evenShare = 1 / editableChildren.length;
    return Object.fromEntries(editableChildren.map(childId => [childId, evenShare]));
  });

  editableChildren.forEach(childId => {
    const nextSeries = YEARS.map((_year, index) => {
      const lockedTotal = lockedChildren.reduce((sum, lockedChildId) => {
        return sum + Number(childSeriesById[lockedChildId]?.[index] || 0);
      }, 0);
      return (Number(parentSeries[index] || 0) - lockedTotal) * Number(sharesByYear[index]?.[childId] || 0);
    });
    setMetricControlPointsFromSeries(childId, nextSeries);
  });
  setSourceStatus(`Fit ${definition.label} children to parent using current unlocked child mix.`, "success");
  render();
}

function syncWorkspaceNapkinChart() {
  if (!workspaceNapkinChart || !workspaceNapkinMetricId) return;
  const parsed = parseNodeKey(workspaceNapkinMetricId);
  const definition = metricDefinitions()[parsed.metricId];
  if (!definition) return;
  const points = activeTopDownControlPoints(definition);
  const context = selectedDimensionContextForMetric(definition.id);
  const label = context
    ? `${definition.label}: ${context.label}`
    : parsed.memberId
      ? `${definition.label}: ${dimensionMembers(parsed.dimensionId).find(member => member.id === parsed.memberId)?.label || parsed.memberId}`
      : definition.label;
  workspaceNapkinIsSyncing = true;
  const lines = [{
    name: label,
    color: "#111827",
    data: points,
    editable: true,
  }];
  const bottomUp = activeWorkspaceBottomUpSeries(definition);
  if (bottomUp) {
    lines.push({
      name: "Bottom-Up",
      color: "#4f7fb8",
      data: YEARS.map((year, index) => [year, Number(bottomUp[index] || 0)]),
      editable: false,
    });
  }
  workspaceNapkinChart.lines = lines;
  workspaceNapkinChart.globalMaxX = YEARS[YEARS.length - 1];
  workspaceNapkinChart.windowStartX = YEARS[0];
  workspaceNapkinChart.windowEndX = YEARS[YEARS.length - 1];
  workspaceNapkinChart._refreshChart();
  scheduleWorkspaceNapkinReferenceLineStyle();
  workspaceNapkinIsSyncing = false;
}

function scheduleWorkspaceNapkinReferenceLineStyle() {
  if (workspaceNapkinReferenceStyleFrame) return;
  workspaceNapkinReferenceStyleFrame = requestAnimationFrame(() => {
    workspaceNapkinReferenceStyleFrame = null;
    styleWorkspaceNapkinReferenceLines();
  });
}

function styleNapkinBottomUpReference(chart) {
  if (!chart?.chart) return;
  const option = chart.chart.getOption();
  const series = (option.series || []).map(seriesItem => {
    if (seriesItem.name !== "Bottom-Up") {
      return { ...seriesItem, z: 4 };
    }
    return {
      ...seriesItem,
      silent: true,
      showSymbol: false,
      symbolSize: 0,
      z: 3,
      lineStyle: {
        ...(seriesItem.lineStyle || {}),
        width: 2,
        type: "dashed",
        color: "#4f7fb8",
      },
    };
  });
  chart.chart.setOption({ series }, false);
}

function styleWorkspaceNapkinReferenceLines() {
  styleNapkinBottomUpReference(workspaceNapkinChart);
}

function commitWorkspaceNapkinEdit(definition) {
  workspaceNapkinHasPendingCommit = false;
  refreshDetailAfterTopDownEdit(definition);
  renderMetricList();
  renderCanvas();
  requestAnimationFrame(renderCanvasLines);
  renderCatalogSource();
}

function workspaceGapSummary(definition) {
  const topDown = activeTopDownSeries(definition);
  const bottomUp = activeWorkspaceBottomUpSeries(definition);
  return {
    topDown,
    bottomUp,
    summary: bottomUp ? seriesDeltaSummary(topDown, bottomUp) : null,
    reconciliation: metricReconciliation(definition.id),
  };
}

function updateWorkspaceHeaderStatus(definition) {
  const badge = document.querySelector("[data-workspace-reconciliation-badge]");
  if (!badge || !definition) return;
  const { reconciliation, summary } = workspaceGapSummary(definition);
  const workspaceMatched = summary ? Math.abs(summary.total) <= (definition.reconciliation?.tolerance || 1) : true;
  badge.textContent = reconciliation.enabled
    ? workspaceMatched
      ? "Matched"
      : `Gap ${formatMetricValue(definition, summary?.total || reconciliation.maxDelta, { precise: true })}`
    : "Manual";
  badge.className = `reconciliation-badge ${reconciliation.enabled ? workspaceMatched ? "matched" : "open" : ""}`;
}

function updateWorkspaceSummaryCards(definition) {
  const topElement = document.querySelector('[data-workspace-summary-value="top"]');
  const bottomElement = document.querySelector('[data-workspace-summary-value="bottom"]');
  const gapElement = document.querySelector('[data-workspace-summary-value="gap"]');
  const gapCard = document.querySelector('[data-workspace-summary-card="gap"]');
  if (!definition || !topElement || !bottomElement || !gapElement) return;
  const { topDown, bottomUp, summary } = workspaceGapSummary(definition);
  topElement.textContent = formatMetricValue(definition, totalSeriesValue(topDown), { precise: true });
  bottomElement.textContent = bottomUp ? formatMetricValue(definition, totalSeriesValue(bottomUp), { precise: true }) : "-";
  gapElement.textContent = summary ? formatMetricValue(definition, summary.total, { precise: true }) : "-";
  if (gapCard) {
    const isOpen = summary && Math.abs(summary.total) > (definition.reconciliation?.tolerance || 1);
    gapCard.classList.toggle("open", Boolean(isOpen));
    gapCard.classList.toggle("matched", !isOpen);
  }
}

function updateWorkspaceChildValues(definition) {
  childIds(definition.id).forEach(childId => {
    const child = metricDefinitions()[childId];
    const element = document.querySelector(`[data-workspace-child-value="${CSS.escape(childId)}"]`);
    if (element) {
      element.textContent = formatMetricValue(child, totalSeriesValue(metricTopDownSeries(childId)), { precise: true });
    }
  });
}

function updateWorkspaceDimensionGroupValues(definition) {
  const mixContext = dimensionMixContext(definition);
  if (!mixContext) return;
  dimensionGroupsForContext(mixContext).forEach(group => {
    const element = document.querySelector(`[data-workspace-dimension-group-value="${CSS.escape(group.id)}"]`);
    if (element) {
      element.textContent = formatMetricValue(definition, totalSeriesValue(dimensionGroupSeries(mixContext, group)), { precise: true });
    }
  });
}

function syncWorkspaceChildNapkinChart(chartKey, item) {
  const chart = item?.chart;
  const childId = item?.childId;
  const parentMetricId = item?.parentMetricId;
  const child = metricDefinitions()[childId];
  if (!chart || !child) return;
  const locked = workspaceIsLocked(workspaceChildLockKey(parentMetricId, childId));
  const points = manualControlPoints(childId);
  const bottomUp = metricDirectBottomUpSeries(childId);
  workspaceChildNapkinChartSyncing.add(chartKey);
  chart.lines = [
    {
      name: child.label || childId,
      color: child.presentation?.color || "#111827",
      data: points,
      editable: !locked,
    },
    ...(bottomUp ? [{
      name: "Bottom-Up",
      color: "#4f7fb8",
      data: YEARS.map((year, index) => [year, Number(bottomUp[index] || 0)]),
      editable: false,
    }] : []),
  ];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart._refreshChart();
  styleNapkinBottomUpReference(chart);
  workspaceChildNapkinChartSyncing.delete(chartKey);
}

function syncWorkspaceChildNapkinCharts(skipChartKey = null) {
  workspaceChildNapkinCharts.forEach((item, chartKey) => {
    if (chartKey === skipChartKey || item?.chart?._isDragging) return;
    syncWorkspaceChildNapkinChart(chartKey, item);
  });
}

function syncWorkspaceDimensionGroupNapkinChart(chartKey, item) {
  const chart = item?.chart;
  const definition = item?.definition;
  const mixContext = item?.mixContext;
  const group = item?.group;
  if (!chart || !definition || !mixContext || !group) return;
  const locked = workspaceIsLocked(workspaceDimensionGroupLockKey(mixContext, group.id));
  const points = workspaceDimensionGroupControlPoints(definition, mixContext, group);
  const bottomUp = dimensionGroupBottomUpSeries(mixContext, group);
  workspaceDimensionMemberChartSyncing.add(chartKey);
  chart.lines = [
    {
      name: group.label,
      color: group.color || dimensionGroupColor(group.id),
      data: points,
      editable: !locked,
    },
    ...(bottomUp ? [{
      name: "Bottom-Up",
      color: "#4f7fb8",
      data: YEARS.map((year, index) => [year, Number(bottomUp[index] || 0)]),
      editable: false,
    }] : []),
  ];
  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart._refreshChart();
  styleNapkinBottomUpReference(chart);
  workspaceDimensionMemberChartSyncing.delete(chartKey);
}

function syncWorkspaceDimensionGroupNapkinCharts(skipChartKey = null) {
  workspaceDimensionMemberCharts.forEach((item, chartKey) => {
    if (chartKey === skipChartKey || item?.chart?._isDragging) return;
    syncWorkspaceDimensionGroupNapkinChart(chartKey, item);
  });
}

function refreshWorkspaceDimensionGroupLines(definition) {
  if (!window.echarts) return;
  const element = document.getElementById("metric-workspace-dimension-group-lines-chart");
  if (!element) return;
  const existing = echarts.getInstanceByDom(element);
  if (existing) {
    existing.dispose();
    metricWorkspaceCharts = metricWorkspaceCharts.filter(item => item !== existing);
  }
  renderWorkspaceDimensionGroupLinesChart(definition, dimensionMixContext(definition));
}

function refreshMetricWorkspaceLive(options = {}) {
  const definition = selectedDefinition();
  if (!definition) return;
  updateWorkspaceHeaderStatus(definition);
  updateWorkspaceSummaryCards(definition);
  updateWorkspaceChildValues(definition);
  updateWorkspaceDimensionGroupValues(definition);
  if (!options.skipWorkspaceNapkin) syncWorkspaceNapkinChart();
  syncWorkspaceChildNapkinCharts(options.skipChildChartKey || null);
  syncWorkspaceDimensionMixChart();
  syncWorkspaceDimensionGroupNapkinCharts(options.skipDimensionGroupChartKey || null);
  refreshWorkspaceDimensionGroupLines(definition);
}

function scheduleMetricWorkspaceLiveRefresh(options = {}) {
  pendingWorkspaceLiveRefreshOptions = {
    ...(pendingWorkspaceLiveRefreshOptions || {}),
    ...options,
  };
  if (workspaceLiveRefreshFrame) return;
  workspaceLiveRefreshFrame = requestAnimationFrame(() => {
    const nextOptions = pendingWorkspaceLiveRefreshOptions || {};
    workspaceLiveRefreshFrame = null;
    pendingWorkspaceLiveRefreshOptions = null;
    refreshMetricWorkspaceLive(nextOptions);
  });
}

function renderWorkspaceNapkinEditor(definition) {
  const container = document.getElementById("metric-workspace-napkin-chart");
  if (!container || !definition || metricIsLagged(definition)) return;
  if (typeof NapkinChart !== "function" || !window.echarts) {
    container.innerHTML = `<div class="chart-empty">NapkinChart did not load.</div>`;
    return;
  }

  const currentSelectionKey = selectedNodeKey();
  if (workspaceNapkinMetricId === currentSelectionKey && workspaceNapkinChart) {
    syncWorkspaceNapkinChart();
    return;
  }

  disposeWorkspaceNapkinChart();
  container.innerHTML = "";
  const points = activeTopDownControlPoints(definition);
  const bottomUp = activeWorkspaceBottomUpSeries(definition);
  const rawYMax = Math.max(10, ...points.map(point => Number(point[1] || 0)), ...(bottomUp || [])) * 1.25;
  const yMax = snapSafeNapkinYMax(rawYMax);
  const context = selectedDimensionContextForMetric(definition.id);
  const lineName = context ? `${definition.label}: ${context.label}` : definition.label;
  workspaceNapkinMetricId = currentSelectionKey;
  workspaceNapkinChart = new NapkinChart(
    "metric-workspace-napkin-chart",
    [
    {
      name: lineName,
      color: "#111827",
      data: points,
      editable: true,
    },
    ...(bottomUp ? [{
      name: "Bottom-Up",
      color: "#4f7fb8",
      data: YEARS.map((year, index) => [year, Number(bottomUp[index] || 0)]),
      editable: false,
    }] : []),
    ],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1 },
      yAxis: { type: "value", min: 0, max: yMax, axisLabel: { formatter: value => formatMetricValue(definition, value) } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: { trigger: "axis", valueFormatter: value => formatMetricValue(definition, value, { precise: true }) },
    },
    "none",
    false
  );
  syncWorkspaceNapkinChart();
  scheduleWorkspaceNapkinReferenceLineStyle();

  workspaceNapkinChart.chart.getZr().on("mouseup", () => {
    if (!workspaceNapkinHasPendingCommit) return;
    requestAnimationFrame(() => {
      if (workspaceNapkinHasPendingCommit) commitWorkspaceNapkinEdit(definition);
    });
  });

  workspaceNapkinChart.onDataChanged = () => {
    if (workspaceNapkinIsSyncing || !workspaceNapkinChart?.lines?.[0]) return;
    scheduleWorkspaceNapkinReferenceLineStyle();
    const nextPoints = normalizeNapkinPoints(workspaceNapkinChart.lines[0].data);
    setActiveTopDownSeries(definition, nextPoints);
    if (workspaceNapkinChart._isDragging) {
      workspaceNapkinHasPendingCommit = true;
      scheduleMetricWorkspaceLiveRefresh({ skipWorkspaceNapkin: true });
      return;
    }
    refreshDetailAfterTopDownEdit(definition);
    commitWorkspaceNapkinEdit(definition);
  };
}

function syncWorkspaceDimensionMixChart() {
  if (!workspaceDimensionMixChart || !workspaceDimensionMixMetricId) return;
  const definition = metricDefinitions()[workspaceDimensionMixMetricId];
  const mixContext = dimensionMixContext(definition);
  const groups = mixContext ? dimensionGroupsForContext(mixContext) : [];
  if (!definition || !mixContext || groups.length < 2) return;

  const lines = dimensionGroupShareBoundaryLines(mixContext);
  const topGroup = groups[0];
  workspaceDimensionMixIsSyncing = true;
  workspaceDimensionMixChart.lines = lines;
  workspaceDimensionMixChart.topAreaLabel = topGroup.label;
  workspaceDimensionMixChart.topAreaColor = topGroup.color;
  workspaceDimensionMixChart.baseOption.topAreaLabel = workspaceDimensionMixChart.topAreaLabel;
  workspaceDimensionMixChart.baseOption.topAreaColor = workspaceDimensionMixChart.topAreaColor;
  workspaceDimensionMixChart.globalMaxX = YEARS[YEARS.length - 1];
  workspaceDimensionMixChart.windowStartX = YEARS[0];
  workspaceDimensionMixChart.windowEndX = YEARS[YEARS.length - 1];
  workspaceDimensionMixChart._refreshChart();
  workspaceDimensionMixChart.resize();
  workspaceDimensionMixIsSyncing = false;
}

function commitWorkspaceDimensionMixEdit(definition) {
  workspaceDimensionMixHasPendingCommit = false;
  refreshDetailAfterDimensionMixEdit(definition);
  renderMetricList();
  renderCanvas();
  renderMetricWorkspacePanel();
  requestAnimationFrame(renderCanvasLines);
  renderCatalogSource();
}

function renderWorkspaceDimensionMixEditor(definition) {
  const container = document.getElementById("metric-workspace-dimension-mix-chart");
  const mixContext = dimensionMixContext(definition);
  const groups = mixContext ? dimensionGroupsForContext(mixContext) : [];
  if (!container || !definition || !mixContext || groups.length < 2) {
    disposeWorkspaceDimensionMixChart();
    return;
  }

  if (typeof NapkinChartArea !== "function" || !window.echarts) {
    disposeWorkspaceDimensionMixChart();
    container.innerHTML = `<div class="chart-empty">NapkinChartArea did not load.</div>`;
    return;
  }

  if (workspaceDimensionMixMetricId === definition.id && workspaceDimensionMixChart) {
    syncWorkspaceDimensionMixChart();
    return;
  }

  disposeWorkspaceDimensionMixChart();
  container.innerHTML = "";
  workspaceDimensionMixMetricId = definition.id;
  workspaceDimensionMixChart = new NapkinChartArea(
    "metric-workspace-dimension-mix-chart",
    dimensionGroupShareBoundaryLines(mixContext),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1 },
      yAxis: {
        type: "value",
        min: 0,
        max: 100,
        axisLabel: { formatter: value => `${Number(value || 0).toFixed(0)}%` },
      },
      topAreaLabel: groups[0].label,
      topAreaColor: groups[0].color,
      areaTint: 0.42,
      grid: { left: 12, right: 18, top: 18, bottom: 34, containLabel: true },
      tooltip: { trigger: "axis" },
    },
    "none",
    false
  );
  workspaceDimensionMixChart.enableZoomBar = false;
  syncWorkspaceDimensionMixChart();

  workspaceDimensionMixChart.chart.getZr().on("mouseup", () => {
    if (!workspaceDimensionMixHasPendingCommit) return;
    requestAnimationFrame(() => {
      if (workspaceDimensionMixHasPendingCommit) commitWorkspaceDimensionMixEdit(definition);
    });
  });

  workspaceDimensionMixChart.onDataChanged = () => {
    if (workspaceDimensionMixIsSyncing || !workspaceDimensionMixChart?.lines) return;
    applyDimensionGroupShares(mixContext, workspaceDimensionMixChart.lines);
    refreshDetailAfterDimensionMixEdit(definition);
    if (workspaceDimensionMixChart._isDragging) {
      workspaceDimensionMixHasPendingCommit = true;
      return;
    }
    commitWorkspaceDimensionMixEdit(definition);
  };
}

function workspaceDimensionGroupControlPoints(_definition, mixContext, group) {
  return YEARS.map((year, index) => [year, Number(dimensionGroupSeries(mixContext, group)[index] || 0)]);
}

function clearAncestorControlPointsForCoordinate(metricId, coordinate) {
  const metric = ensureMetricScenario(metricId);
  const editedKey = coordinateKey(coordinate);
  if (metric.controlPoints && typeof metric.controlPoints === "object") {
    Object.keys(metric.controlPoints).forEach(key => {
      if (key === editedKey) return;
      const filters = parseCoordinateKey(key);
      if (coordinateMatchesFilters(editedKey, filters)) delete metric.controlPoints[key];
    });
    if (!Object.keys(metric.controlPoints).length) delete metric.controlPoints;
  }
  if (editedKey !== TOTAL_COORDINATE_KEY) delete metric.manualControlPoints;
}

function syncWorkspaceTotalsAfterDimensionMemberEdit(definition, mixContext) {
  const metric = ensureMetricScenario(definition.id);
  if (metric.dimensionShareControlPointsByContext) {
    delete metric.dimensionShareControlPointsByContext[dimensionMixControlKey(mixContext)];
    if (!Object.keys(metric.dimensionShareControlPointsByContext).length) {
      delete metric.dimensionShareControlPointsByContext;
    }
  }
  if (
    !mixContext.baseKey
    && mixContext.dimensionId === primaryDimensionId(mixContext.metricId)
    && metricDimensionIds(mixContext.metricId).length === 1
  ) {
    delete metric.dimensionShareControlPoints;
  }
  clearCalculationCache();
  syncWorkspaceDimensionMixChart();
  syncWorkspaceNapkinChart();
}

function renderWorkspaceDimensionGroupNapkin(definition, mixContext, group, index) {
  const container = document.getElementById(`metric-workspace-dimension-chart-${index}`);
  if (!container || !definition || !mixContext || typeof NapkinChart !== "function" || !window.echarts) return;
  const groupKey = `${dimensionMixControlKey(mixContext)}::${group.id}`;
  const lockKey = workspaceDimensionGroupLockKey(mixContext, group.id);
  const locked = workspaceIsLocked(lockKey);
  const points = workspaceDimensionGroupControlPoints(definition, mixContext, group);
  const bottomUp = dimensionGroupBottomUpSeries(mixContext, group);
  const rawYMax = Math.max(10, ...points.map(point => Number(point[1] || 0)), ...(bottomUp || [])) * 1.25;
  const yMax = snapSafeNapkinYMax(rawYMax);
  const lineColor = group.color || dimensionGroupColor(group.id);

  container.innerHTML = "";
  const chart = new NapkinChart(
    `metric-workspace-dimension-chart-${index}`,
    [
      {
        name: group.label,
        color: lineColor,
        data: points,
        editable: !locked,
      },
      ...(bottomUp ? [{
        name: "Bottom-Up",
        color: "#4f7fb8",
        data: YEARS.map((year, yearIndex) => [year, Number(bottomUp[yearIndex] || 0)]),
        editable: false,
      }] : []),
    ],
    true,
    {
      animation: false,
      xAxis: {
        type: "value",
        min: YEARS[0],
        max: YEARS[YEARS.length - 1],
        minInterval: 1,
      },
      yAxis: { type: "value", min: 0, max: yMax, axisLabel: { formatter: value => formatMetricValue(definition, value) } },
      grid: { left: 12, right: 14, top: 8, bottom: 28, containLabel: true },
      tooltip: { trigger: "axis", valueFormatter: value => formatMetricValue(definition, value, { precise: true }) },
    },
    "none",
    false
  );

  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart._refreshChart();
  styleNapkinBottomUpReference(chart);

  const applyGroupEdit = ({ refreshDetail = false } = {}) => {
    if (locked || !chart?.lines?.[0]) return;
    const nextPoints = normalizeNapkinPoints(chart.lines[0].data);
    const series = seriesFromControlPoints(nextPoints);
    setDimensionGroupSeries(mixContext, group, series);
    group.members.forEach(member => {
      clearAncestorControlPointsForCoordinate(definition.id, { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id });
    });
    syncWorkspaceTotalsAfterDimensionMemberEdit(definition, mixContext);
    if (refreshDetail) refreshDetailAfterDimensionMixEdit(definition);
  };

  workspaceDimensionMemberCharts.set(groupKey, { chart, definition, mixContext, group });
  chart.chart.getZr().on("mouseup", () => {
    if (!workspaceDimensionMemberChartPendingCommits.has(groupKey)) return;
    requestAnimationFrame(() => {
      if (!workspaceDimensionMemberChartPendingCommits.has(groupKey)) return;
      workspaceDimensionMemberChartPendingCommits.delete(groupKey);
      applyGroupEdit({ refreshDetail: true });
      renderMetricWorkspacePanel();
      renderMetricList();
      renderCanvas();
      requestAnimationFrame(renderCanvasLines);
      renderCatalogSource();
    });
  });

  chart.onDataChanged = () => {
    if (locked || workspaceDimensionMemberChartSyncing.has(groupKey) || !chart?.lines?.[0]) return;
    if (chart._isDragging) {
      applyGroupEdit();
      workspaceDimensionMemberChartPendingCommits.add(groupKey);
      scheduleMetricWorkspaceLiveRefresh({ skipDimensionGroupChartKey: groupKey });
      return;
    }
    applyGroupEdit({ refreshDetail: true });
    workspaceDimensionMemberChartPendingCommits.delete(groupKey);
    renderMetricWorkspacePanel();
    renderMetricList();
    renderCanvas();
    requestAnimationFrame(renderCanvasLines);
    renderCatalogSource();
  };
}

function renderWorkspaceDimensionMemberNapkins(definition, mixContext) {
  disposeWorkspaceDimensionMemberCharts();
  if (!mixContext) return;
  dimensionGroupsForContext(mixContext).forEach((group, index) => {
    renderWorkspaceDimensionGroupNapkin(definition, mixContext, group, index);
  });
}

function renderWorkspaceChildNapkin(parentMetricId, childId, index) {
  const child = metricDefinitions()[childId];
  const container = document.getElementById(`metric-workspace-child-chart-${index}`);
  if (!container || !child || typeof NapkinChart !== "function" || !window.echarts) return;

  const points = manualControlPoints(childId);
  const bottomUp = metricDirectBottomUpSeries(childId);
  const rawYMax = Math.max(10, ...points.map(point => Number(point[1] || 0)), ...(bottomUp || [])) * 1.25;
  const yMax = snapSafeNapkinYMax(rawYMax);
  const lineColor = child.presentation?.color || "#111827";
  const chartKey = `${childId}::${index}`;
  const lockKey = workspaceChildLockKey(parentMetricId, childId);
  const locked = workspaceIsLocked(lockKey);

  container.innerHTML = "";
  const chart = new NapkinChart(
    `metric-workspace-child-chart-${index}`,
    [
      {
        name: child.label || childId,
        color: lineColor,
        data: points,
        editable: !locked,
      },
      ...(bottomUp ? [{
        name: "Bottom-Up",
        color: "#4f7fb8",
        data: YEARS.map((year, yearIndex) => [year, Number(bottomUp[yearIndex] || 0)]),
        editable: false,
      }] : []),
    ],
    true,
    {
      animation: false,
      xAxis: {
        type: "value",
        min: YEARS[0],
        max: YEARS[YEARS.length - 1],
        minInterval: 1,
      },
      yAxis: { type: "value", min: 0, max: yMax, axisLabel: { formatter: value => formatMetricValue(child, value) } },
      grid: { left: 12, right: 14, top: 8, bottom: 28, containLabel: true },
      tooltip: { trigger: "axis", valueFormatter: value => formatMetricValue(child, value, { precise: true }) },
    },
    "none",
    false
  );

  chart.globalMaxX = YEARS[YEARS.length - 1];
  chart.windowStartX = YEARS[0];
  chart.windowEndX = YEARS[YEARS.length - 1];
  chart._refreshChart();

  styleNapkinBottomUpReference(chart);

  const applyChildEdit = ({ refreshDetail = false } = {}) => {
    if (locked || !chart?.lines?.[0]) return;
    const nextPoints = normalizeNapkinPoints(chart.lines[0].data);
    const series = seriesFromControlPoints(nextPoints);
    setMetricControlPointsFromSeries(childId, series);
    if (refreshDetail) refreshDetailAfterTopDownEdit(child);
  };

  workspaceChildNapkinCharts.set(chartKey, { chart, parentMetricId, childId });
  chart.chart.getZr().on("mouseup", () => {
    if (!workspaceChildNapkinChartPendingCommits.has(chartKey)) return;
    requestAnimationFrame(() => {
      if (!workspaceChildNapkinChartPendingCommits.has(chartKey)) return;
      workspaceChildNapkinChartPendingCommits.delete(chartKey);
      applyChildEdit({ refreshDetail: true });
      renderMetricWorkspacePanel();
      renderMetricList();
      renderCanvas();
      requestAnimationFrame(renderCanvasLines);
      renderCatalogSource();
    });
  });

  chart.onDataChanged = () => {
    if (locked || workspaceChildNapkinChartSyncing.has(chartKey) || !chart?.lines?.[0]) return;
    if (chart._isDragging) {
      applyChildEdit();
      workspaceChildNapkinChartPendingCommits.add(chartKey);
      scheduleMetricWorkspaceLiveRefresh({ skipChildChartKey: chartKey });
      return;
    }
    applyChildEdit({ refreshDetail: true });
    workspaceChildNapkinChartPendingCommits.delete(chartKey);
    renderMetricWorkspacePanel();
    renderMetricList();
    renderCanvas();
    requestAnimationFrame(renderCanvasLines);
    renderCatalogSource();
  };
}

function renderWorkspaceDimensionGroupLinesChart(definition, mixContext) {
  const element = document.getElementById("metric-workspace-dimension-group-lines-chart");
  if (!element || !window.echarts || !definition || !mixContext) return;
  const groups = dimensionGroupsForContext(mixContext);
  const chartType = state.dimensionGroupLinesChartType === "bar" ? "bar" : "line";
  const groupedSeries = state.dimensionGroupLinesShowGrouped ? groups.map(group => ({
    name: group.label,
    data: dimensionGroupSeries(mixContext, group),
    color: group.color,
    width: 3,
    layer: "grouped",
  })) : [];
  const rawSeries = state.dimensionGroupLinesShowUngrouped ? mixContext.members.map(member => ({
    name: member.label,
    data: metricTopDownFilteredSeries(
      definition.id,
      { ...mixContext.baseFilters, [mixContext.dimensionId]: member.id }
    ),
    color: "#98a2b3",
    width: 1.5,
    type: "dashed",
    layer: "ungrouped",
  })) : [];
  const seriesList = [...groupedSeries, ...rawSeries];
  if (!seriesList.length) {
    element.innerHTML = `<div class="metric-workspace-empty"><strong>No lines selected</strong><span>Turn on Grouped, Ungrouped, or both.</span></div>`;
    return;
  }
  const instance = echarts.init(element);
  const values = seriesList.flatMap(item => item.data || []);
  const maxValue = Math.max(1, ...values.map(value => Math.abs(Number(value || 0))));
  instance.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      valueFormatter: value => formatMetricValue(definition, value, { precise: true }),
    },
    legend: { top: 0, type: "scroll" },
    grid: { left: 10, right: 10, top: 34, bottom: 24, containLabel: true },
    xAxis: {
      type: "category",
      data: YEARS.map(String),
      axisTick: { alignWithLabel: true },
    },
    yAxis: {
      type: "value",
      min: value => Math.min(0, value.min),
      max: value => Math.max(maxValue, value.max),
      axisLabel: { formatter: value => formatMetricValue(definition, value) },
    },
    series: seriesList.map(item => ({
      name: item.name,
      type: chartType,
      data: item.data,
      symbolSize: item.layer === "grouped" ? 6 : 3,
      barGap: item.layer === "grouped" ? "20%" : "0%",
      barMaxWidth: item.layer === "grouped" ? 28 : 14,
      lineStyle: {
        width: item.width || 2,
        type: chartType === "line" ? item.type || "solid" : "solid",
        opacity: item.layer === "ungrouped" ? 0.7 : 1,
      },
      itemStyle: {
        color: item.color,
        opacity: item.layer === "ungrouped" ? 0.55 : 0.95,
      },
    })),
  }, true);
  metricWorkspaceCharts.push(instance);
}

function renderWorkspaceWeightedAverageCharts(definition) {
  if (!definition || definition.bottomUp?.type !== "weightedAverage") return;
  const context = selectedDimensionContextForMetric(definition.id);
  const filters = context?.coordinate || {};
  const { valueMetricId, weightMetricId } = weightedAverageConfig(definition);
  const valueDefinition = metricDefinitions()[valueMetricId];
  const weightDefinition = metricDefinitions()[weightMetricId];
  const weightedSeries = activeWorkspaceBottomUpSeries(definition);
  const topDownSeries = activeTopDownSeries(definition);
  const valueSeries = formulaInputSeriesForFilters(valueMetricId, filters);
  const weightSeries = formulaInputSeriesForFilters(weightMetricId, filters);

  renderWorkspaceChart("metric-workspace-weighted-rate-chart", definition, [
    { name: "Top-Down Rate", data: topDownSeries, color: "#111827" },
    ...(weightedSeries ? [{ name: "Weighted Average", data: weightedSeries, color: "#b45309", width: 3 }] : []),
  ]);
  if (valueDefinition && valueSeries) {
    renderWorkspaceChart("metric-workspace-weighted-value-chart", valueDefinition, [
      { name: valueDefinition.label, data: valueSeries, color: valueDefinition.presentation?.color || "#7b6fb8", width: 3 },
    ], { compact: true });
  }
  if (weightDefinition && weightSeries) {
    renderWorkspaceChart("metric-workspace-weighted-weight-chart", weightDefinition, [
      { name: weightDefinition.label, data: weightSeries, color: weightDefinition.presentation?.color || "#2f6f73", width: 3 },
    ], { compact: true });
  }
}

function renderMetricWorkspacePanel() {
  const panel = document.getElementById("metric-workspace-panel");
  if (!panel) return;
  metricWorkspaceRenderToken += 1;
  const renderToken = metricWorkspaceRenderToken;
  cancelWorkspaceLiveRefresh();
  disposeMetricWorkspaceCharts();
  disposeWorkspaceNapkinChart();
  disposeWorkspaceDimensionMixChart();
  disposeWorkspaceDimensionMemberCharts();
  disposeWorkspaceChildNapkinCharts();
  const definition = selectedDefinition();
  if (!definition) {
    panel.innerHTML = `
      <div class="metric-workspace-empty">
        <strong>Select a metric</strong>
        <span>The workspace will show the selected metric and its immediate children.</span>
      </div>
    `;
    return;
  }

  const metricId = definition.id;
  const children = childIds(metricId);
  const bottomUp = activeWorkspaceBottomUpSeries(definition);
  const topDown = activeTopDownSeries(definition);
  const reconciliation = metricReconciliation(metricId);
  const formulaType = definition.bottomUp?.type || "manual";
  const canFitChildren = formulaType === "sum" && children.length > 0;
  const isWeightedAverage = formulaType === "weightedAverage";
  const metricTypeLabel = isWeightedAverage
    ? "Weighted average metric"
    : children.length
      ? `${children.length} immediate child${children.length === 1 ? "" : "ren"}`
      : "Leaf metric";
  const gap = bottomUp ? seriesDeltaSummary(topDown, bottomUp).total : null;
  const workspaceMatched = bottomUp ? Math.abs(gap) <= (definition.reconciliation?.tolerance || 1) : true;
  const mode = activeWorkspaceMode(definition);
  const modeContent = mode === "children"
    ? workspaceChildrenModeHtml(definition, children)
    : mode === "weightedAverage"
      ? workspaceWeightedAverageModeHtml(definition)
      : mode === "dimensions"
        ? workspaceDimensionsModeHtml(definition)
        : mode === "cohort"
          ? workspaceCohortModeHtml(definition)
          : "";

  panel.innerHTML = `
    <div class="metric-workspace-heading">
      <div>
        <p class="eyebrow">Metric Workspace</p>
        <h2>${escapeHtml(definition.label)}</h2>
        <span>${metricTypeLabel} | ${escapeHtml(formulaText(metricId))}</span>
      </div>
      <div class="metric-workspace-actions">
        <span data-workspace-reconciliation-badge class="reconciliation-badge ${reconciliation.enabled ? workspaceMatched ? "matched" : "open" : ""}">
          ${reconciliation.enabled ? workspaceMatched ? "Matched" : `Gap ${formatMetricValue(definition, gap || reconciliation.maxDelta, { precise: true })}` : "Manual"}
        </span>
        <button type="button" data-workspace-set-parent="${escapeHtml(metricId)}" ${bottomUp ? "" : "disabled"}>${isWeightedAverage ? "Set to Weighted Avg" : "Set Parent to Children"}</button>
        ${isWeightedAverage ? "" : `<button type="button" data-workspace-fit-children="${escapeHtml(metricId)}" ${canFitChildren ? "" : "disabled"}>Fit Children to Parent</button>`}
      </div>
    </div>
    ${workspaceTotalModeHtml(definition, bottomUp)}
    ${workspaceModeTabs(definition)}
    ${modeContent || `
      <section class="metric-workspace-children">
        <div class="metric-workspace-empty">
          <strong>No explanation lens yet</strong>
          <span>This metric is currently just the selected total above. Add dimensions, children, or cohort logic to create a deeper editing view.</span>
        </div>
      </section>
    `}
  `;

  requestAnimationFrame(() => {
    if (renderToken !== metricWorkspaceRenderToken) return;
    if (metricIsLagged(definition)) {
      const selectedSeries = [{ name: "Top-Down", data: topDown, color: "#111827" }];
      if (bottomUp) selectedSeries.push({ name: "Bottom-Up", data: bottomUp, color: "#4f7fb8", type: "dashed" });
      renderWorkspaceChart("metric-workspace-selected-chart", definition, selectedSeries);
    } else {
      renderWorkspaceNapkinEditor(definition);
    }

    if (mode === "children") {
      children.forEach((childId, index) => renderWorkspaceChildNapkin(definition.id, childId, index));
    } else if (mode === "weightedAverage") {
      renderWorkspaceWeightedAverageCharts(definition);
      renderWorkspaceDimensionMemberNapkins(definition, dimensionMixContext(definition));
    } else if (mode === "dimensions") {
      const mixContext = dimensionMixContext(definition);
      const groups = mixContext ? dimensionGroupsForContext(mixContext) : [];
      if (groups.length >= 2 && state.dimensionMixView === "pctTotal") {
        renderWorkspaceDimensionMixEditor(definition);
      } else {
        disposeWorkspaceDimensionMixChart();
        renderWorkspaceDimensionGroupLinesChart(definition, mixContext);
      }
      renderWorkspaceDimensionMemberNapkins(definition, mixContext);
    }
  });
}

function addAssumption() {
  const definition = selectedDefinition();
  if (!definition) return;
  const claimInput = document.getElementById("assumption-claim-input");
  const driverInput = document.getElementById("assumption-driver-input");
  const ownerInput = document.getElementById("assumption-owner-input");
  const impactInput = document.getElementById("assumption-impact-input");
  const claim = claimInput.value.trim();
  if (!claim) {
    setSourceStatus("Add a claim before saving the assumption.", "error");
    claimInput.focus();
    return;
  }
  const timestamp = new Date().toISOString();
  assumptions().push({
    id: uniqueAssumptionId(),
    metricId: definition.id,
    scope: assumptionScopeForSelection(definition.id),
    baselineRole: "bau",
    comparisonRole: "target",
    years: [...YEARS],
    claim,
    driverType: driverInput.value || "initiative",
    owner: ownerInput.value.trim() || null,
    impact: {
      mode: "pctOfDelta",
      pct: Number(impactInput.value || 0),
    },
    evidence: [],
    status: "draft",
    createdAt: timestamp,
    updatedAt: timestamp,
  });
  claimInput.value = "";
  ownerInput.value = "";
  impactInput.value = "";
  catalogSourceIsDirty = false;
  setSourceStatus(`Added assumption for ${definition.label}.`, "success");
  render();
}

function deleteAssumption(assumptionId) {
  const index = assumptions().findIndex(item => item.id === assumptionId);
  if (index < 0) return;
  const [removed] = assumptions().splice(index, 1);
  catalogSourceIsDirty = false;
  setSourceStatus(`Removed assumption ${removed.id}.`, "success");
  render();
}

function referenceSnapshotMatchesCurrent(role) {
  const reference = referenceScenarios()[role];
  if (!reference?.snapshot) return false;
  return JSON.stringify(reference.snapshot.metrics || {}) === JSON.stringify(scenarioSnapshot().metrics || {});
}

function setReferenceScenario(role) {
  if (!["bau", "target"].includes(role)) return;
  const summary = activeScenarioReconciliationSummary();
  const label = role.toUpperCase();
  if (!summary.matched) {
    setSourceStatus(
      `${label} can only be set when top-down and bottom-up match. Resolve ${summary.openMetricCount} open metric${summary.openMetricCount === 1 ? "" : "s"} first.`,
      "error"
    );
    return;
  }
  if (referenceSnapshotMatchesCurrent(role)) {
    setSourceStatus(`Current scenario already matches ${label}.`, "success");
    return;
  }
  const source = scenarioSnapshot();
  const timestamp = new Date().toISOString();
  const reference = {
    id: `${role}-${timestamp.replace(/[^0-9]/g, "")}`,
    role,
    name: label,
    assignedAt: timestamp,
    sourceScenarioId: source.id,
    sourceScenarioName: source.name || source.id,
    sourceScenarioUpdatedAt: source.updatedAt || null,
    snapshot: source,
  };
  referenceScenarios()[role] = reference;
  scenarioRoles()[role] = reference.id;
  catalogSourceIsDirty = false;
  setSourceStatus(`${label} set from ${source.name || source.id}.`, "success");
  render();
}

function setSourceStatus(message, tone = "") {
  const status = document.getElementById("catalog-source-status");
  status.textContent = message;
  status.className = `source-status ${tone}`;
}

function renderCatalogSource() {
  const source = document.getElementById("catalog-source");
  if (document.activeElement === source && catalogSourceIsDirty) {
    renderJsonExplorerFromText(source.value);
    return;
  }
  source.value = sourceSnapshot();
  renderJsonExplorerFromText(source.value);
  catalogSourceIsDirty = false;
}

function normalizeImportedModel(rawModel) {
  if (!rawModel || typeof rawModel !== "object") {
    throw new Error("Catalog source must be a JSON object.");
  }
  if (!rawModel.metricDefinitions || typeof rawModel.metricDefinitions !== "object") {
    throw new Error("Catalog source must include a metricDefinitions object.");
  }
  const rawScenarios = rawModel.scenarios && typeof rawModel.scenarios === "object"
    ? rawModel.scenarios
    : rawModel.scenario && typeof rawModel.scenario === "object"
      ? { [rawModel.scenario.id || "scenario-1"]: rawModel.scenario }
      : null;
  if (!rawScenarios) {
    throw new Error("Catalog source must include a scenarios object.");
  }

  const importedYears = normalizeYears(rawModel.years);
  YEARS = [...importedYears];
  setEngineYears(YEARS);

  const definitions = {};
  Object.entries(rawModel.metricDefinitions).forEach(([id, definition]) => {
    if (!definition || typeof definition !== "object") {
      throw new Error(`Metric ${id} must be an object.`);
    }
    const normalizedDefinition = {
      ...definition,
      id,
      label: definition.label || id,
      unit: definition.unit || "currency",
      dimensions: Array.isArray(definition.dimensions) ? definition.dimensions : [],
      topDown: definition.topDown || { type: "manualSeries", editable: true },
      bottomUp: definition.bottomUp || null,
      reconciliation: definition.reconciliation || { enabled: Boolean(definition.bottomUp), tolerance: 1 },
      time: definition.time || createManualMetric({ id, label: definition.label || id }).time,
      presentation: definition.presentation || { color: "#4f7fb8" },
    };
    if (metricIsLagged(normalizedDefinition)) {
      normalizedDefinition.topDown = laggedTopDownDefinition();
    }
    definitions[id] = normalizedDefinition;
  });

  const scenarios = {};
  Object.entries(rawScenarios).forEach(([rawScenarioId, rawScenario]) => {
    const scenarioId = rawScenario?.id || rawScenarioId || "scenario-1";
    const metrics = {};
    Object.keys(definitions).forEach(id => {
      metrics[id] = normalizeScenarioMetric(definitions[id], rawScenario?.metrics?.[id] || {});
    });
    scenarios[scenarioId] = {
      ...rawScenario,
      id: scenarioId,
      name: rawScenario?.name || scenarioId,
      metrics,
    };
  });
  const firstScenarioId = Object.keys(scenarios)[0] || "scenario-1";
  if (!scenarios[firstScenarioId]) {
    scenarios[firstScenarioId] = { id: firstScenarioId, name: "Scenario 1", metrics: {} };
  }
  const activeId = rawModel.activeScenarioId && scenarios[rawModel.activeScenarioId]
    ? rawModel.activeScenarioId
    : firstScenarioId;
  const importedReferences = rawModel.referenceScenarios && typeof rawModel.referenceScenarios === "object"
    ? rawModel.referenceScenarios
    : {};
  const referenceScenariosValue = {
    bau: importedReferences.bau || null,
    target: importedReferences.target || null,
  };
  const importedRoles = rawModel.scenarioRoles && typeof rawModel.scenarioRoles === "object"
    ? rawModel.scenarioRoles
    : {};

  return {
    name: rawModel.name || "Custom Model",
    years: importedYears,
    scenarioRoles: {
      working: activeId,
      bau: referenceScenariosValue.bau?.id || importedRoles.bau || null,
      target: referenceScenariosValue.target?.id || importedRoles.target || null,
    },
    referenceScenarios: referenceScenariosValue,
    dimensions: rawModel.dimensions || {},
    metricDefinitions: definitions,
    assumptions: Array.isArray(rawModel.assumptions) ? rawModel.assumptions.map(normalizeAssumption) : [],
    activeScenarioId: activeId,
    scenarios,
  };
}

function renderMetricList() {
  const list = document.getElementById("metric-list");
  const definitions = Object.values(metricDefinitions());
  document.getElementById("model-name").textContent = state.model.name;
  if (!definitions.length) {
    list.innerHTML = `
      <div class="empty-state">
        <strong>No metrics yet</strong>
        <span>Add a metric to start building the canvas, or load the starter example.</span>
      </div>
    `;
    return;
  }
  list.innerHTML = definitions.map(definition => {
    const reconciliation = metricReconciliation(definition.id);
    return `
      <button class="metric-list-item ${definition.id === state.selectedMetricId ? "active" : ""}" data-metric-id="${definition.id}" type="button">
        <span>${definition.label}</span>
        ${reconciliation.enabled ? `<small class="${reconciliation.matched ? "matched" : "open"}">${reconciliation.matched ? "Matched" : "Mismatch"}</small>` : "<small>Manual</small>"}
      </button>
    `;
  }).join("");
}

function renderDimensionCatalog() {
  const list = document.getElementById("dimension-list");
  const nameInput = document.getElementById("dimension-name-input");
  const membersInput = document.getElementById("dimension-members-input");
  const dimensions = Object.values(dimensionDefinitions());

  if (!dimensions.length) {
    list.innerHTML = `
      <div class="empty-state">
        <strong>No dimensions yet</strong>
        <span>Create one to segment metric values.</span>
      </div>
    `;
  } else {
    list.innerHTML = dimensions.map(dimension => {
      const members = dimensionMembers(dimension.id);
      const assignedCount = Object.values(metricDefinitions()).filter(definition => metricDimensionIds(definition.id).includes(dimension.id)).length;
      return `
        <button class="dimension-list-item ${dimension.id === state.selectedDimensionId ? "active" : ""}" data-catalog-dimension-id="${dimension.id}" type="button">
          <span>${dimension.label}</span>
          <small>${members.length} members${assignedCount ? ` | ${assignedCount} metrics` : ""}</small>
        </button>
      `;
    }).join("");
  }

  const selectedDimension = state.selectedDimensionId ? dimensionDefinitions()[state.selectedDimensionId] : null;
  nameInput.value = selectedDimension?.label || "";
  membersInput.value = selectedDimension ? dimensionMembersText(selectedDimension.id) : "";
}

function renderMetricDimensionAssignment(definition) {
  const section = document.getElementById("metric-dimension-section");
  const input = document.getElementById("metric-dimension-input");
  const note = document.getElementById("metric-dimension-note");
  if (!definition || metricIsLagged(definition) || selectedDimensionContextForMetric(definition.id)) {
    section.classList.add("is-hidden");
    input.innerHTML = "";
    note.textContent = "";
    return;
  }

  const dimensions = Object.values(dimensionDefinitions());
  section.classList.remove("is-hidden");
  const assignedDimensions = metricDimensionIds(definition.id);
  const maxLevels = Math.min(dimensions.length, Math.max(1, assignedDimensions.length + 1));
  input.innerHTML = dimensions.length
    ? Array.from({ length: maxLevels }, (_item, index) => {
      const selectedId = assignedDimensions[index] || "";
      return `
        <label class="metric-dimension-level">
          <span>Level ${index + 1}</span>
          <select data-metric-dimension-level="${index}">
            <option value="">None</option>
            ${dimensions.map(dimension => `
              <option value="${escapeHtml(dimension.id)}" ${dimension.id === selectedId ? "selected" : ""}>
                ${escapeHtml(dimension.label)}
              </option>
            `).join("")}
          </select>
        </label>
      `;
    }).join("")
    : `<div class="empty-state">Create a dimension in the catalog before assigning one to this metric.</div>`;

  if (assignedDimensions.length) {
    const coordinateCount = metricCoordinateKeys(definition.id).length;
    note.textContent = `Current hierarchy: ${assignedDimensions.map((dimensionId, index) => `Level ${index + 1} ${dimensionDefinitions()[dimensionId]?.label || dimensionId}`).join(" -> ")} (${coordinateCount} coordinate${coordinateCount === 1 ? "" : "s"}). Editing total reallocates values across the current coordinate mix.`;
  } else if (dimensions.length) {
    note.textContent = "Level 1 is the first way you want to break the metric down; Level 2 is the next drilldown inside Level 1. Applying preserves the current total and splits it evenly across new coordinates.";
  } else {
    note.textContent = "Create a dimension in the catalog before assigning one to this metric.";
  }
}

function renderCanvas() {
  const canvas = document.getElementById("metric-canvas");
  renderCanvasViewToggle();
  canvas.classList.toggle("chart-mode", state.canvasViewMode === "chart");
  const levels = canvasLevels();
  if (!levels.length) {
    canvas.innerHTML = `
      <div class="canvas-empty">
        <strong>Blank canvas</strong>
        <span>Create the first metric, then define how it connects to other metrics.</span>
      </div>
    `;
    return;
  }

  canvas.innerHTML = `
    <svg id="canvas-lines" class="canvas-lines" aria-hidden="true"></svg>
    ${levels.map((level, levelIndex) => `
      <div class="canvas-level dynamic-level" style="--level-index: ${levelIndex}">
        ${level.map(nodeKey => renderNode(nodeKey)).join("")}
      </div>
    `).join("")}
  `;
}

function renderCanvasViewToggle() {
  const toggle = document.getElementById("canvas-view-toggle");
  if (!toggle) return;
  toggle.innerHTML = ["number", "chart"].map(mode => `
    <button class="view-toggle-button ${state.canvasViewMode === mode ? "active" : ""}" type="button" data-canvas-view-mode="${mode}">
      ${mode === "number" ? "Numbers" : "Charts"}
    </button>
  `).join("");
}

function inlineChartPath(values, minValue, maxValue, width, height, margin) {
  const range = maxValue - minValue || 1;
  return values.map((value, index) => {
    const x = margin.left + (values.length === 1 ? 0 : (index / (values.length - 1)) * (width - margin.left - margin.right));
    const y = margin.top + ((maxValue - Number(value || 0)) / range) * (height - margin.top - margin.bottom);
    return `${index === 0 ? "M" : "L"} ${trimFixed(x, 2)} ${trimFixed(y, 2)}`;
  }).join(" ");
}

function renderInlineMetricChart(definition, topDown, bottomUp = null) {
  const width = 220;
  const height = 84;
  const margin = { top: 10, right: 8, bottom: 16, left: 8 };
  const allValues = [...topDown, ...(bottomUp || [])].map(value => Number(value || 0));
  const rawMin = Math.min(0, ...allValues);
  const rawMax = Math.max(0, ...allValues);
  const padding = rawMax === rawMin ? 1 : (rawMax - rawMin) * 0.08;
  const minValue = rawMin - padding;
  const maxValue = rawMax + padding;
  const zeroY = margin.top + ((maxValue - 0) / (maxValue - minValue || 1)) * (height - margin.top - margin.bottom);
  const topDownPath = inlineChartPath(topDown, minValue, maxValue, width, height, margin);
  const bottomUpPath = bottomUp ? inlineChartPath(bottomUp, minValue, maxValue, width, height, margin) : "";
  return `
    <div class="metric-node-chart" aria-hidden="true">
      <svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
        <line x1="${margin.left}" y1="${trimFixed(zeroY, 2)}" x2="${width - margin.right}" y2="${trimFixed(zeroY, 2)}" class="metric-node-chart-axis"></line>
        <path d="${topDownPath}" class="metric-node-chart-line top-down"></path>
        ${bottomUp ? `<path d="${bottomUpPath}" class="metric-node-chart-line bottom-up"></path>` : ""}
      </svg>
      <div class="metric-node-chart-legend">
        <span><i class="legend-line top-down"></i>Top</span>
        ${bottomUp ? `<span><i class="legend-line bottom-up"></i>Bottom</span>` : ""}
      </div>
    </div>
  `;
}

function renderNode(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  const definition = metricDefinitions()[parsed.metricId];
  const isMember = Boolean(parsed.memberId);
  const member = isMember
    ? dimensionMembers(parsed.dimensionId).find(item => item.id === parsed.memberId)
    : null;
  const { topDown: series, bottomUp } = canvasNodeComparisonSeries(nodeKey);
  const reconciliation = isMember
    ? metricMemberReconciliation(parsed.metricId, parsed.dimensionId, parsed.memberId)
    : metricReconciliation(parsed.metricId);
  const isSelected = nodeKey === selectedNodeKey();
  const children = nodeChildKeys(nodeKey);
  const dimensionIds = isMember ? [] : metricDimensionIds(parsed.metricId);
  const coordinateCount = isMember || !dimensionIds.length ? 0 : metricCoordinateKeys(parsed.metricId).length;
  const nodeCaption = isMember
    ? definition.bottomUp ? "Member top-down" : "Member series"
    : [
      children.length
        ? `${children.length} inputs`
        : definition.bottomUp?.type === "laggedMetric"
          ? "Lagged metric"
          : definition.bottomUp
            ? "Top-down + bottom-up"
            : "Manual series",
      dimensionIds.length ? `${dimensionIds.map((_dimensionId, index) => `L${index + 1}`).join("/")} | ${coordinateCount} coordinates` : "",
    ].filter(Boolean).join(" | ");
  return `
    <button class="metric-node ${isMember ? "member-node" : ""} ${isSelected ? "active" : ""} ${reconciliation.enabled && !reconciliation.matched ? "mismatch" : ""}" data-node-key="${nodeKey}" type="button">
      <span>${isMember ? member?.label || parsed.memberId : definition.label}</span>
      <strong>${formatMetricValue(definition, series[series.length - 1])}</strong>
      <small>${nodeCaption}</small>
      ${state.canvasViewMode === "chart" ? renderInlineMetricChart(definition, series, bottomUp) : ""}
    </button>
  `;
}

function renderCanvasLines() {
  const canvas = document.getElementById("metric-canvas");
  const svg = document.getElementById("canvas-lines");
  if (!svg) return;

  const nodeById = new Map();
  canvas.querySelectorAll(".metric-node").forEach(node => {
    nodeById.set(node.dataset.nodeKey, node);
  });

  const canvasRect = canvas.getBoundingClientRect();
  const pointForNode = (node, edge) => {
    const rect = node.getBoundingClientRect();
    return {
      x: rect.left - canvasRect.left + canvas.scrollLeft + rect.width / 2,
      y: rect.top - canvasRect.top + canvas.scrollTop + (edge === "bottom" ? rect.height : 0),
    };
  };

  const paths = [];
  nodeById.forEach((parentNode, nodeKey) => {
    if (!parentNode) return;
    nodeChildKeys(nodeKey).forEach(childKey => {
      const childNode = nodeById.get(childKey);
      if (!childNode) return;
      const start = pointForNode(parentNode, "bottom");
      const end = pointForNode(childNode, "top");
      const midY = start.y + Math.max(28, (end.y - start.y) / 2);
      paths.push(`<path d="M ${start.x} ${start.y} C ${start.x} ${midY}, ${end.x} ${midY}, ${end.x} ${end.y}" />`);
    });
  });

  const width = Math.max(canvas.scrollWidth, canvas.clientWidth);
  const height = Math.max(canvas.scrollHeight, canvas.clientHeight);
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.style.width = `${width}px`;
  svg.style.height = `${height}px`;
  svg.innerHTML = paths.join("");
}

function renderEmptyDetail() {
  document.getElementById("detail-title").textContent = "Metric Detail";
  document.getElementById("detail-subtitle").textContent = "Select or add a metric.";
  document.getElementById("top-down-definition").textContent = "-";
  document.getElementById("bottom-up-definition").textContent = "-";
  document.getElementById("time-definition").textContent = "-";
  document.getElementById("dimension-definition").textContent = "-";
  renderDimensionMemberControls(null);
  renderMetricDimensionAssignment(null);
  const badge = document.getElementById("reconciliation-badge");
  badge.textContent = "-";
  badge.className = "reconciliation-badge";
  renderBottomUpForm(null);
  renderCohortMatrixBuilder(null);
  renderStockFlowBuilder(null);
  renderTopDownNapkinEditor(null);
  renderLaggedOpeningEditor(null);
  renderAssumptionPanel(null);
  renderDimensionBreakdown(null);
  renderDimensionMixEditor(null);
  chartElement.classList.add("is-hidden");
  chart.clear();
  if (!window.echarts) {
    chartElement.innerHTML = `<div class="chart-empty">ECharts did not load. The catalog tools still work.</div>`;
  }
}

function renderDetail() {
  const definition = selectedDefinition();
  if (!definition) {
    renderEmptyDetail();
    return;
  }

  const metricId = definition.id;
  const selectedMember = selectedDimensionContextForMetric(metricId);
  if (selectedMember) {
    renderMemberDetail(definition, selectedMember);
    return;
  }

  const parents = parentIds(metricId).map(id => metricDefinitions()[id].label);
  const children = childIds(metricId).map(id => metricDefinitions()[id].label);

  document.getElementById("detail-title").textContent = definition.label;
  document.getElementById("detail-subtitle").textContent = [
    parents.length ? `Parent: ${parents.join(", ")}` : "Root metric",
    children.length ? `Inputs: ${children.join(", ")}` : "Leaf metric",
  ].join(" | ");
  document.getElementById("top-down-definition").textContent = definition.topDown
    ? `${definition.topDown.type}${definition.topDown.editable ? " (editable)" : ""}`
    : "None";
  document.getElementById("bottom-up-definition").textContent = formulaText(metricId);
  document.getElementById("time-definition").textContent = `${definition.time.flowType}, ${definition.time.aggregateMethod}`;
  document.getElementById("dimension-definition").textContent = metricDimensionIds(metricId)
    .map((dimensionId, index) => `L${index + 1} ${dimensionDefinitions()[dimensionId]?.label || dimensionId}`)
    .join(" -> ") || "None";
  renderMetricDimensionAssignment(definition);
  renderDimensionMemberControls(definition);

  renderReconciliationBadge(metricId);
  renderBottomUpForm(definition);
  renderCohortMatrixBuilder(definition);
  renderStockFlowBuilder(definition);
  renderTopDownNapkinEditor(null);
  renderLaggedOpeningEditor(definition);
  renderAssumptionPanel(definition);
  renderDimensionBreakdown(metricIsLagged(definition) ? null : definition);
  renderDimensionMixEditor(null);
  renderDefinitionComparisonChart(definition);
}

function renderMemberDetail(definition, context) {
  document.getElementById("detail-title").textContent = `${definition.label}: ${context.label}`;
  document.getElementById("detail-subtitle").textContent = `Filtered coordinate view`;
  document.getElementById("top-down-definition").textContent = definition.topDown
    ? `${definition.topDown.type}${definition.topDown.editable ? " (editable)" : ""}`
    : "None";
  document.getElementById("bottom-up-definition").textContent = definition.bottomUp
    ? `${formulaText(definition.id)} in ${context.label} context`
    : "No member-specific bottom-up definition";
  document.getElementById("time-definition").textContent = `${definition.time.flowType}, ${definition.time.aggregateMethod}`;
  document.getElementById("dimension-definition").textContent = context.label;
  renderMetricDimensionAssignment(null);
  renderDimensionMemberControls(definition);

  renderReconciliationBadge(definition.id);
  renderBottomUpForm(null);
  document.getElementById("formula-input-list").innerHTML = `
    <div class="empty-state">
      <strong>Member slice</strong>
      <span>This tile can edit its own top-down series. Member-specific bottom-up formulas come next.</span>
    </div>
  `;
  renderCohortMatrixBuilder(null);
  renderStockFlowBuilder(null);
  renderTopDownNapkinEditor(null);
  renderLaggedOpeningEditor(definition);
  renderAssumptionPanel(definition);
  renderDimensionBreakdown(null);
  renderDimensionMixEditor(null);
  renderDefinitionComparisonChart(definition);
}

function renderAssumptionPanel(definition) {
  const section = document.getElementById("assumption-panel");
  const deltaGrid = document.getElementById("assumption-delta-grid");
  const list = document.getElementById("assumption-list");
  const summary = document.getElementById("assumption-summary");
  if (!definition) {
    section.classList.add("is-hidden");
    deltaGrid.innerHTML = "";
    list.innerHTML = "";
    return;
  }

  const context = selectedDimensionContextForMetric(definition.id);
  const filters = context?.coordinate || {};
  const activeSeries = activeTopDownSeries(definition);
  const bauSeries = referenceMetricSeries("bau", definition.id, filters);
  const targetSeries = referenceMetricSeries("target", definition.id, filters);
  const bauDelta = bauSeries ? seriesDeltaSummary(activeSeries, bauSeries) : null;
  const targetDelta = targetSeries ? seriesDeltaSummary(activeSeries, targetSeries) : null;
  const attached = assumptionsForSelection(definition.id);
  const coverage = assumptionCoverage(attached);
  const coverageReferenceDelta = bauDelta || targetDelta;
  const explainedValue = coverageReferenceDelta ? explainedDeltaValue(coverageReferenceDelta.total, coverage.pct) : null;
  const unexplainedValue = coverageReferenceDelta ? coverageReferenceDelta.total - explainedValue : null;
  const coverageLabel = coverage.status === "over"
    ? "Over-explained"
    : coverage.status === "full"
      ? "Fully explained"
      : coverage.status === "partial"
        ? "Partially explained"
        : "No explanation";

  section.classList.remove("is-hidden");
  summary.textContent = context
    ? `Assumptions attached to ${definition.label}: ${context.label}.`
    : `Assumptions attached to ${definition.label}.`;
  deltaGrid.innerHTML = `
    <div>
      <span>Delta to BAU</span>
      <strong>${bauDelta ? formatMetricValue(definition, bauDelta.total, { precise: true }) : "BAU unset"}</strong>
    </div>
    <div>
      <span>Delta to Target</span>
      <strong>${targetDelta ? formatMetricValue(definition, targetDelta.total, { precise: true }) : "Target unset"}</strong>
    </div>
    <div>
      <span>Attached</span>
      <strong>${attached.length}</strong>
    </div>
    <div>
      <span>Coverage</span>
      <strong>${trimFixed(coverage.pct, 1)}% | ${coverageLabel}</strong>
    </div>
    <div>
      <span>Explained</span>
      <strong>${explainedValue !== null ? formatMetricValue(definition, explainedValue, { precise: true }) : "No reference"}</strong>
    </div>
    <div>
      <span>Unexplained</span>
      <strong>${unexplainedValue !== null ? formatMetricValue(definition, unexplainedValue, { precise: true }) : "No reference"}</strong>
    </div>
  `;

  list.innerHTML = attached.length ? attached.map(item => `
    <article class="assumption-card">
      <div>
        <strong>${escapeHtml(item.claim)}</strong>
        <span>${escapeHtml(item.driverType || "other")}${item.owner ? ` | ${escapeHtml(item.owner)}` : ""} | ${trimFixed(Number(item.impact?.pct || 0), 1)}% of delta</span>
      </div>
      <button type="button" data-delete-assumption="${escapeHtml(item.id)}" aria-label="Delete ${escapeHtml(item.id)}">x</button>
    </article>
  `).join("") : `
    <div class="empty-state compact-empty">
      <strong>No assumptions yet</strong>
      <span>Add a claim that explains why this metric should differ from BAU or Target.</span>
    </div>
  `;
}

function renderDimensionMemberControls(definition) {
  const section = document.getElementById("dimension-member-section");
  const list = document.getElementById("dimension-member-list");
  if (!definition) {
    section.classList.add("is-hidden");
    list.innerHTML = "";
    return;
  }

  const dimensionIds = metricDimensionIds(definition.id);
  if (!dimensionIds.length) {
    section.classList.add("is-hidden");
    list.innerHTML = "";
    return;
  }

  const selectedContext = selectedDimensionContextForMetric(definition.id);
  section.classList.remove("is-hidden");
  document.getElementById("dimension-member-title").textContent = "Coordinate Filters";
  list.innerHTML = [
    `<button class="dimension-member-chip ${selectedContext ? "" : "active"}" data-member-total="${definition.id}" type="button">Total</button>`,
    ...dimensionIds.map(dimensionId => {
      const selectedMemberId = selectedContext?.coordinate?.[dimensionId] || "";
      return `
        <label class="dimension-filter-field">
          <span>${escapeHtml(metricDimensionLevelLabel(definition.id, dimensionId))}</span>
          <select data-coordinate-filter="${escapeHtml(definition.id)}" data-dimension-id="${escapeHtml(dimensionId)}">
            <option value="">All</option>
            ${dimensionMembers(dimensionId).map(member => `
              <option value="${escapeHtml(member.id)}" ${member.id === selectedMemberId ? "selected" : ""}>${escapeHtml(member.label)}</option>
            `).join("")}
          </select>
        </label>
      `;
    }),
  ].join("");
}

function renderDimensionBreakdown(definition) {
  const section = document.getElementById("dimension-breakdown-section");
  if (!definition) {
    section.classList.add("is-hidden");
    dimensionBreakdownChart.clear();
    return;
  }

  const dimensionId = primaryDimensionId(definition.id);
  const dimensionIds = metricDimensionIds(definition.id);
  const coordinateKeys = metricCoordinateKeys(definition.id);
  if (!dimensionIds.length || !coordinateKeys.length) {
    section.classList.add("is-hidden");
    dimensionBreakdownChart.clear();
    return;
  }

  section.classList.remove("is-hidden");
  document.getElementById("dimension-breakdown-title").textContent = `${dimensionIds.map((id, index) => `Level ${index + 1}: ${dimensionDefinitions()[id]?.label || id}`).join(" -> ")} Top-Down Breakdown`;
  if (!window.echarts) return;

  dimensionBreakdownChart.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      valueFormatter: value => compactCurrencyTooltip(value),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 34, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: compactCurrency } },
    series: coordinateKeys.map(key => ({
      name: coordinateLabelForMetric(definition.id, key),
      type: "line",
      data: metricTopDownCoordinateSeries(definition.id, key),
      symbolSize: 6,
      lineStyle: { width: 2.5 },
    })),
  }, true);
}

function disposeDimensionMixChart() {
  if (dimensionMixChart?.chart) {
    dimensionMixChart.chart.dispose();
  }
  dimensionMixChart = null;
  dimensionMixMetricId = null;
  dimensionMixHasPendingCommit = false;
}

function syncDimensionMixChart() {
  if (!dimensionMixChart || !dimensionMixMetricId) return;
  const definition = metricDefinitions()[dimensionMixMetricId];
  const dimensionId = primaryDimensionId(dimensionMixMetricId);
  const members = dimensionMembers(dimensionId);
  if (!definition || members.length < 2) return;

  const lines = dimensionShareBoundaryLines(definition.id, members);
  const topMember = members[0];
  dimensionMixIsSyncing = true;
  dimensionMixChart.lines = lines;
  dimensionMixChart.topAreaLabel = topMember.label;
  dimensionMixChart.topAreaColor = dimensionMemberColor(members, topMember.id);
  dimensionMixChart.baseOption.topAreaLabel = dimensionMixChart.topAreaLabel;
  dimensionMixChart.baseOption.topAreaColor = dimensionMixChart.topAreaColor;
  dimensionMixChart.globalMaxX = YEARS[YEARS.length - 1];
  dimensionMixChart.windowStartX = YEARS[0];
  dimensionMixChart.windowEndX = YEARS[YEARS.length - 1];
  dimensionMixChart._refreshChart();
  dimensionMixChart.resize();
  dimensionMixIsSyncing = false;
}

function renderDimensionMixEditor(definition) {
  const section = document.getElementById("dimension-mix-section");
  const container = document.getElementById("dimension-mix-chart");
  const dimensionId = definition ? primaryDimensionId(definition.id) : null;
  const members = dimensionMembers(dimensionId);
  const shouldShow = Boolean(definition && metricDimensionIds(definition.id).length === 1 && !metricIsLagged(definition) && !metricIsRate(definition) && dimensionId && members.length >= 2);
  section.classList.toggle("is-hidden", !shouldShow);

  if (!shouldShow) {
    disposeDimensionMixChart();
    container.innerHTML = "";
    return;
  }

  document.getElementById("dimension-mix-title").textContent = `${dimensionDefinitions()[dimensionId]?.label || dimensionId} % Total`;

  if (typeof NapkinChartArea !== "function" || !window.echarts) {
    disposeDimensionMixChart();
    container.innerHTML = `<div class="chart-empty">NapkinChartArea did not load.</div>`;
    return;
  }

  if (dimensionMixMetricId === definition.id && dimensionMixChart) {
    syncDimensionMixChart();
    return;
  }

  disposeDimensionMixChart();
  container.innerHTML = "";
  dimensionMixMetricId = definition.id;
  dimensionMixChart = new NapkinChartArea(
    "dimension-mix-chart",
    dimensionShareBoundaryLines(definition.id, members),
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1 },
      yAxis: {
        type: "value",
        min: 0,
        max: 100,
        axisLabel: { formatter: value => `${Number(value || 0).toFixed(0)}%` },
      },
      topAreaLabel: members[0].label,
      topAreaColor: dimensionMemberColor(members, members[0].id),
      areaTint: 0.42,
      grid: { left: 12, right: 18, top: 18, bottom: 34, containLabel: true },
      tooltip: { trigger: "axis" },
    },
    "none",
    false
  );
  dimensionMixChart.enableZoomBar = false;
  syncDimensionMixChart();

  dimensionMixChart.chart.getZr().on("mouseup", () => {
    if (!dimensionMixHasPendingCommit) return;
    requestAnimationFrame(() => {
      if (dimensionMixHasPendingCommit) commitDimensionMixEdit(definition);
    });
  });

  dimensionMixChart.onDataChanged = () => {
    if (dimensionMixIsSyncing || !dimensionMixChart?.lines) return;
    applyDimensionShares(definition.id, members, dimensionMixChart.lines);
    refreshDetailAfterDimensionMixEdit(definition);
    if (dimensionMixChart._isDragging) {
      dimensionMixHasPendingCommit = true;
      return;
    }
    commitDimensionMixEdit(definition);
  };
}

function refreshDetailAfterDimensionMixEdit(definition) {
  renderDimensionBreakdown(definition);
  renderReconciliationBadge(definition.id);
  renderDefinitionComparisonChart(definition);
}

function commitDimensionMixEdit(definition) {
  dimensionMixHasPendingCommit = false;
  refreshDetailAfterDimensionMixEdit(definition);
  renderMetricList();
  renderCanvas();
  requestAnimationFrame(renderCanvasLines);
  renderCatalogSource();
}

function renderReconciliationBadge(metricId) {
  const context = selectedDimensionContextForMetric(metricId);
  const reconciliation = context
    ? metricFilteredReconciliation(metricId, context)
    : metricReconciliation(metricId);
  const badge = document.getElementById("reconciliation-badge");
  badge.textContent = reconciliation.enabled
    ? reconciliation.matched ? "Matched" : `Mismatch ${formatCurrency(reconciliation.maxDelta)}`
    : "Manual";
  badge.className = `reconciliation-badge ${reconciliation.enabled ? reconciliation.matched ? "matched" : "open" : ""}`;
}

function renderDefinitionComparisonChart(definition) {
  if (!definition) {
    chartElement.classList.add("is-hidden");
    chart.clear();
    return;
  }

  const metricId = definition.id;
  const context = selectedDimensionContextForMetric(metricId);
  const laggedLines = metricIsLagged(definition)
    ? laggedComparisonLines(definition, context)
    : null;
  const topDown = laggedLines?.topDown || activeTopDownSeries(definition);
  const bottomUp = laggedLines ? laggedLines.bottomUp : activeBottomUpSeries(definition);
  const topDownName = laggedLines ? `${laggedLines.sourceLabel} Top-Down (Lagged)` : "Top-Down";
  const bottomUpName = laggedLines ? `${laggedLines.sourceLabel} Bottom-Up (Lagged)` : "Bottom-Up";
  const series = [{
    name: topDownName,
    type: "line",
    data: topDown,
    symbolSize: 7,
    lineStyle: { color: "#2f6f73", width: 3 },
    itemStyle: { color: "#2f6f73" },
  }];
  if (bottomUp) {
    series.push({
      name: bottomUpName,
      type: "line",
      data: bottomUp,
      symbolSize: 6,
      lineStyle: { color: "#4f7fb8", width: 3 },
      itemStyle: { color: "#4f7fb8" },
    });
  }
  const cohortTotal = context ? null : cohortTotalSeries(metricId);
  if (cohortTotal) {
    series.push({
      name: "Cohort Total",
      type: "line",
      data: cohortTotal,
      symbolSize: 5,
      lineStyle: { color: "#7b6fb8", width: 3 },
      itemStyle: { color: "#7b6fb8" },
    });
    Object.entries(metricCohorts(metricId)).forEach(([cohortYear, row]) => {
      series.push({
        name: `Cohort ${cohortYear}`,
        type: "line",
        data: YEARS.map(year => cohortValueForYear(row, year)),
        symbolSize: 4,
        lineStyle: { width: 1.5 },
      });
    });
  }
  if (definition.bottomUp?.type === "cohortMatrixFromStartsAndAgeYoy") {
    const matrix = generatedCohortMatrix(metricId);
    Object.entries(matrix).forEach(([cohortYear, row]) => {
      series.push({
        name: `Cohort ${cohortYear}`,
        type: "line",
        data: YEARS.map(year => cohortValueForYear(row, year)),
        symbolSize: 4,
        lineStyle: { width: 1.5 },
      });
    });
  }
  if (definition.bottomUp?.type === "cohortMatrix") {
    const matrix = generatedGenericCohortMatrix(metricId, context?.coordinate || {});
    Object.entries(matrix).forEach(([cohortYear, row]) => {
      series.push({
        name: cohortYear.startsWith("opening_") ? cohortYear.replace("opening_", "Opening ") : `Cohort ${cohortYear}`,
        type: "line",
        data: YEARS.map(year => cohortValueForYear(row, year)),
        symbolSize: 4,
        lineStyle: { width: 1.5 },
      });
    });
  }

  if (!window.echarts) {
    chartElement.innerHTML = `<div class="chart-empty">ECharts did not load. The catalog tools still work.</div>`;
    return;
  }

  if (!definition.bottomUp) {
    chartElement.classList.add("is-hidden");
    chart.clear();
    return;
  }
  chartElement.classList.remove("is-hidden");

  chart.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      valueFormatter: value => compactCurrencyTooltip(value),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 34, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: compactCurrency } },
    series,
  }, true);
}

function refreshDetailAfterTopDownEdit(definition) {
  renderReconciliationBadge(definition.id);
  if (selectedDimensionContextForMetric(definition.id) || metricIsLagged(definition)) {
    renderDimensionBreakdown(null);
    renderDimensionMixEditor(null);
  } else {
    renderDimensionBreakdown(definition);
    renderDimensionMixEditor(null);
  }
  renderLaggedOpeningEditor(definition);
  renderDefinitionComparisonChart(definition);
}

function commitTopDownEdit(definition) {
  topDownNapkinHasPendingCommit = false;
  refreshDetailAfterTopDownEdit(definition);
  renderMetricList();
  renderCanvas();
  requestAnimationFrame(renderCanvasLines);
  renderCatalogSource();
}

function renderBottomUpForm(definition) {
  const form = document.getElementById("bottom-up-form");
  const typeInput = document.getElementById("formula-type-input");

  if (!definition) {
    const inputList = document.getElementById("formula-input-list");
    form.classList.add("disabled");
    typeInput.value = "";
    inputList.innerHTML = `<div class="empty-state">Select or create a metric first.</div>`;
    return;
  }

  form.classList.remove("disabled");
  typeInput.value = definition.bottomUp?.type || "";
  renderFormulaInputList(definition);
}

function renderFormulaInputList(definition) {
  const inputList = document.getElementById("formula-input-list");
  if (!definition) {
    inputList.innerHTML = `<div class="empty-state">Select or create a metric first.</div>`;
    return;
  }
  const type = selectedFormulaType();

  const selectedInputs = new Set(metricInputIds(definition));
  const blockedInputs = descendantIds(definition.id);
  blockedInputs.add(definition.id);
  const options = Object.values(metricDefinitions()).filter(candidate => selectedInputs.has(candidate.id) || !blockedInputs.has(candidate.id));

  if (type === "weightedAverage") {
    const { weightMetricId = "" } = weightedAverageConfig(definition);
    const optionMarkup = options.map(candidate => `
      <option value="${escapeHtml(candidate.id)}">${escapeHtml(candidate.label)}</option>
    `).join("");
    inputList.innerHTML = `
      <div class="formula-ratio-grid">
        <label>
          <span>Weight</span>
          <select id="formula-weight-metric">
            <option value="">Select weight metric</option>
            ${optionMarkup}
          </select>
        </label>
      </div>
    `;
    document.getElementById("formula-weight-metric").value = weightMetricId;
    return;
  }

  if (type === "ratio") {
    const [numeratorId = "", denominatorId = ""] = metricInputIds(definition);
    const optionMarkup = options.map(candidate => `
      <option value="${escapeHtml(candidate.id)}">${escapeHtml(candidate.label)}</option>
    `).join("");
    inputList.innerHTML = `
      <div class="formula-ratio-grid">
        <label>
          <span>Numerator</span>
          <select id="formula-ratio-numerator">
            <option value="">Select numerator</option>
            ${optionMarkup}
          </select>
        </label>
        <label>
          <span>Denominator</span>
          <select id="formula-ratio-denominator">
            <option value="">Select denominator</option>
            ${optionMarkup}
          </select>
        </label>
      </div>
    `;
    document.getElementById("formula-ratio-numerator").value = numeratorId;
    document.getElementById("formula-ratio-denominator").value = denominatorId;
    return;
  }

  if (type === "cohortMatrixFromStartsAndAgeYoy" || type === "cohortMatrix" || type === "cumulativeNetFlow") {
    inputList.innerHTML = `<div class="empty-state">This methodology uses the builder below.</div>`;
    return;
  }

  if (!options.length) {
    inputList.innerHTML = `<div class="empty-state">Create another metric to use as an input.</div>`;
    return;
  }

  inputList.innerHTML = options.map(candidate => `
    <label class="formula-input-chip">
      <input type="checkbox" value="${candidate.id}" ${selectedInputs.has(candidate.id) ? "checked" : ""}>
      <span>${candidate.label}</span>
    </label>
  `).join("");
}

function disposeTopDownNapkinChart() {
  if (topDownNapkinChart?.chart) {
    topDownNapkinChart.chart.dispose();
  }
  topDownNapkinChart = null;
  topDownNapkinMetricId = null;
  topDownNapkinHasPendingCommit = false;
}

function renderLaggedOpeningEditor(definition) {
  const form = document.getElementById("lagged-opening-form");
  const inputList = document.getElementById("lagged-opening-inputs");
  const context = definition ? selectedDimensionContextForMetric(definition.id) : null;
  const shouldShow = Boolean(definition && metricIsLagged(definition));
  form.classList.toggle("is-hidden", !shouldShow);

  if (!shouldShow) {
    form.classList.add("disabled");
    inputList.innerHTML = "";
    return;
  }

  form.classList.remove("disabled");
  if (context?.entries?.length === 1) {
    inputList.innerHTML = `
      <label>
        ${context.member.label}
        <input type="number" step="1" data-lagged-opening-member-id="${context.memberId}" value="${metricOpeningValue(definition.id, context.memberId)}">
      </label>
    `;
    return;
  } else if (context) {
    inputList.innerHTML = `
      <div class="empty-state">
        <strong>Filtered opening value</strong>
        <span>Opening values are currently editable by individual dimension member. Clear this filter or select a single-member slice.</span>
      </div>
    `;
    return;
  }

  const dimensionId = primaryDimensionId(definition.id);
  const members = dimensionMembers(dimensionId);
  if (dimensionId && members.length) {
    inputList.innerHTML = members.map(member => `
      <label>
        ${member.label}
        <input type="number" step="1" data-lagged-opening-member-id="${member.id}" value="${metricOpeningValue(definition.id, member.id)}">
      </label>
    `).join("");
    return;
  }

  inputList.innerHTML = `
    <label>
      Opening Value
      <input type="number" step="1" data-lagged-opening-value value="${metricOpeningValue(definition.id)}">
    </label>
  `;
}

function syncTopDownNapkinChart() {
  if (!topDownNapkinChart || !topDownNapkinMetricId) return;
  const parsed = parseNodeKey(topDownNapkinMetricId);
  const definition = metricDefinitions()[parsed.metricId];
  if (!definition) return;
  const points = activeTopDownControlPoints(definition);
  const context = selectedDimensionContextForMetric(definition.id);
  const label = context
    ? `${definition.label}: ${context.label}`
    : parsed.memberId
    ? `${definition.label}: ${dimensionMembers(parsed.dimensionId).find(member => member.id === parsed.memberId)?.label || parsed.memberId}`
    : definition.label;
  topDownNapkinIsSyncing = true;
  topDownNapkinChart.lines = [{
    name: label,
    color: "#111827",
    data: points,
    editable: true,
  }];
  topDownNapkinChart.globalMaxX = YEARS[YEARS.length - 1];
  topDownNapkinChart.windowStartX = YEARS[0];
  topDownNapkinChart.windowEndX = YEARS[YEARS.length - 1];
  topDownNapkinChart._refreshChart();
  topDownNapkinIsSyncing = false;
}

function setTopDownEditorView(view) {
  state.topDownEditorView = view === "sheet" ? "sheet" : "chart";
}

function renderTopDownViewToggle() {
  document.querySelectorAll("[data-top-down-view]").forEach(button => {
    button.classList.toggle("active", button.dataset.topDownView === state.topDownEditorView);
  });
}

function sheetRowKey(metricId, filters) {
  return `${metricId}:${coordinateKey(filters) || "total"}`;
}

function isSheetRowExpanded(metricId, filters) {
  return state.topDownSheetExpandedKeys.has(sheetRowKey(metricId, filters));
}

function toggleSheetRow(metricId, filters) {
  const key = sheetRowKey(metricId, filters);
  if (state.topDownSheetExpandedKeys.has(key)) {
    state.topDownSheetExpandedKeys.delete(key);
  } else {
    state.topDownSheetExpandedKeys.add(key);
  }
}

function seriesForSheetFilters(metricId, filters = {}) {
  return Object.keys(filters).length
    ? metricTopDownFilteredSeries(metricId, filters)
    : metricTopDownSeries(metricId);
}

function buildSheetRollupRows(definition, filters = {}, rows = []) {
  const metricId = definition.id;
  const dimensionIds = metricDimensionIds(metricId);
  const depth = Object.keys(filters).length;
  const nextDimensionId = dimensionIds.find(dimensionId => !filters[dimensionId]);
  const hasChildren = Boolean(nextDimensionId);
  const rowKey = sheetRowKey(metricId, filters);

  rows.push({
    key: rowKey,
    filters: { ...filters },
    depth,
    hasChildren,
    isExpanded: hasChildren && isSheetRowExpanded(metricId, filters),
    series: seriesForSheetFilters(metricId, filters),
  });

  if (!hasChildren || !isSheetRowExpanded(metricId, filters)) return rows;

  dimensionMembers(nextDimensionId).forEach(member => {
    buildSheetRollupRows(definition, { ...filters, [nextDimensionId]: member.id }, rows);
  });
  return rows;
}

function sheetCellLabel(metricId, filters, dimensionId, depth) {
  const memberId = filters[dimensionId];
  if (memberId) {
    return dimensionMembers(dimensionId).find(member => member.id === memberId)?.label || memberId;
  }
  const dimensionIndex = metricDimensionIds(metricId).indexOf(dimensionId);
  if (depth === 0 && dimensionIndex === 0) return "Total";
  return "";
}

function renderTopDownSpreadsheetEditor(definition) {
  const container = document.getElementById("top-down-spreadsheet");
  const context = selectedDimensionContextForMetric(definition.id);
  const baseFilters = context?.coordinate || {};
  const rows = metricDimensionIds(definition.id).length
    ? buildSheetRollupRows(definition, baseFilters)
    : [{
      key: sheetRowKey(definition.id, {}),
      filters: {},
      depth: 0,
      hasChildren: false,
      isExpanded: false,
      series: metricTopDownSeries(definition.id),
    }];
  const dimensionIds = metricDimensionIds(definition.id);
  container.classList.remove("is-hidden");
  container.innerHTML = `
    <div class="top-down-spreadsheet-scroll">
      <table class="top-down-table">
        <thead>
          <tr>
            ${dimensionIds.length
              ? dimensionIds.map((dimensionId, index) => `<th>Level ${index + 1}<br><span>${escapeHtml(dimensionDefinitions()[dimensionId]?.label || dimensionId)}</span></th>`).join("")
              : "<th>Series</th>"}
            ${YEARS.map(year => `<th>${year}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${rows.map(row => `
            <tr class="${row.hasChildren ? "rollup-row" : "leaf-row"}" style="--depth: ${row.depth}">
              ${dimensionIds.length
                ? dimensionIds.map((dimensionId, index) => `
                  <th class="${index === Math.max(0, row.depth - 1) || (row.depth === 0 && index === 0) ? "active-level-cell" : ""}">
                    ${index === Math.max(0, row.depth - 1) || (row.depth === 0 && index === 0)
                      ? row.hasChildren
                        ? `<button class="sheet-expand-button" data-sheet-rollup="${definition.id}" data-rollup-key="${escapeHtml(coordinateKey(row.filters))}" type="button" aria-label="${row.isExpanded ? "Collapse" : "Expand"} rollup">${row.isExpanded ? "-" : "+"}</button>`
                        : `<span class="sheet-leaf-spacer"></span>`
                      : ""}
                    <span>${escapeHtml(sheetCellLabel(definition.id, row.filters, dimensionId, row.depth))}</span>
                  </th>
                `).join("")
                : `<th><span>Top-Down</span></th>`}
              ${YEARS.map((_year, index) => `<td>${compactCurrencyTooltip(row.series[index])}</td>`).join("")}
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderTopDownNapkinEditor(definition) {
  const section = document.getElementById("top-down-napkin-section");
  const container = document.getElementById("top-down-napkin-chart");
  const spreadsheet = document.getElementById("top-down-spreadsheet");
  const shouldShow = Boolean(definition && !metricIsLagged(definition));
  section.classList.toggle("is-hidden", !shouldShow);
  renderTopDownViewToggle();

  if (!shouldShow) {
    disposeTopDownNapkinChart();
    container.innerHTML = "";
    spreadsheet.innerHTML = "";
    spreadsheet.classList.add("is-hidden");
    return;
  }

  if (state.topDownEditorView === "sheet") {
    disposeTopDownNapkinChart();
    container.classList.add("is-hidden");
    container.innerHTML = "";
    renderTopDownSpreadsheetEditor(definition);
    return;
  }

  spreadsheet.innerHTML = "";
  spreadsheet.classList.add("is-hidden");
  container.classList.remove("is-hidden");

  if (typeof NapkinChart !== "function" || !window.echarts) {
    disposeTopDownNapkinChart();
    container.innerHTML = `<div class="chart-empty">NapkinChart did not load.</div>`;
    return;
  }

  const currentSelectionKey = selectedNodeKey();
  if (topDownNapkinMetricId === currentSelectionKey && topDownNapkinChart) {
    syncTopDownNapkinChart();
    return;
  }
  disposeTopDownNapkinChart();
  container.innerHTML = "";

  const points = activeTopDownControlPoints(definition);
  const rawYMax = Math.max(10, ...points.map(point => Number(point[1] || 0))) * 1.25;
  const yMax = snapSafeNapkinYMax(rawYMax);
  const context = selectedDimensionContextForMetric(definition.id);
  const lineName = context ? `${definition.label}: ${context.label}` : definition.label;
  topDownNapkinMetricId = currentSelectionKey;
  topDownNapkinChart = new NapkinChart(
    "top-down-napkin-chart",
    [{
      name: lineName,
      color: "#111827",
      data: points,
      editable: true,
    }],
    true,
    {
      animation: false,
      xAxis: { type: "value", min: YEARS[0], max: YEARS[YEARS.length - 1], minInterval: 1 },
      yAxis: { type: "value", min: 0, max: yMax, axisLabel: { formatter: compactCurrency } },
      grid: { left: 12, right: 18, top: 14, bottom: 34, containLabel: true },
      tooltip: { trigger: "axis", valueFormatter: value => compactCurrencyTooltip(value) },
    },
    "none",
    false
  );
  syncTopDownNapkinChart();

  topDownNapkinChart.chart.getZr().on("mouseup", () => {
    if (!topDownNapkinHasPendingCommit) return;
    requestAnimationFrame(() => {
      if (topDownNapkinHasPendingCommit) commitTopDownEdit(definition);
    });
  });

  topDownNapkinChart.onDataChanged = () => {
    if (topDownNapkinIsSyncing || !topDownNapkinChart?.lines?.[0]) return;
    const nextPoints = normalizeNapkinPoints(topDownNapkinChart.lines[0].data);
    setActiveTopDownSeries(definition, nextPoints);
    refreshDetailAfterTopDownEdit(definition);
    if (topDownNapkinChart._isDragging) {
      topDownNapkinHasPendingCommit = true;
      return;
    }
    commitTopDownEdit(definition);
  };
}

function selectableInputMetricOptions(definition) {
  if (!definition) return [];
  const blockedInputs = descendantIds(definition.id);
  blockedInputs.add(definition.id);
  const currentInputs = new Set(metricInputIds(definition));
  return Object.values(metricDefinitions()).filter(candidate => !blockedInputs.has(candidate.id) || currentInputs.has(candidate.id));
}

function renderSelectOptions(select, options, selectedId, emptyLabel) {
  select.innerHTML = [
    `<option value="">${emptyLabel}</option>`,
    ...options.map(option => `<option value="${option.id}" ${option.id === selectedId ? "selected" : ""}>${option.label}</option>`),
  ].join("");
}

function selectedFormulaType() {
  return document.getElementById("formula-type-input").value;
}

function renderStockFlowBuilder(definition) {
  const form = document.getElementById("stock-flow-builder-form");
  const openingInput = document.getElementById("stock-opening-value-input");
  const inflowInput = document.getElementById("stock-inflow-input");
  const outflowInput = document.getElementById("stock-outflow-input");
  const shouldShow = Boolean(definition && (definition.bottomUp?.type === "cumulativeNetFlow" || selectedFormulaType() === "cumulativeNetFlow"));
  form.classList.toggle("is-hidden", !shouldShow);

  if (!shouldShow) {
    form.classList.add("disabled");
    openingInput.value = "0";
    renderSelectOptions(inflowInput, [], "", "Select additions");
    renderSelectOptions(outflowInput, [], "", "Select losses");
    return;
  }

  form.classList.remove("disabled");
  const options = selectableInputMetricOptions(definition);
  const [inflowId = "", outflowId = ""] = definition.bottomUp?.type === "cumulativeNetFlow"
    ? definition.bottomUp.inputs
    : [];
  openingInput.value = metricOpeningValue(definition.id);
  renderSelectOptions(inflowInput, options, inflowId, "Select additions");
  renderSelectOptions(outflowInput, options, outflowId, "Select losses");
}

function renderCohortMatrixBuilder(definition) {
  const form = document.getElementById("cohort-matrix-builder-form");
  const startSourceInput = document.getElementById("cohort-start-source-input");
  const curveTypeInput = document.getElementById("cohort-curve-type-input");
  const startTimingInput = document.getElementById("cohort-start-timing-input");
  const startsInputs = document.getElementById("cohort-starts-inputs");
  const yoyInputs = document.getElementById("cohort-yoy-inputs");
  const openingInputs = document.getElementById("cohort-opening-inputs");
  const openingCard = document.getElementById("cohort-opening-card");
  const previewCard = document.getElementById("cohort-matrix-preview-card");
  const preview = document.getElementById("cohort-matrix-preview");
  const curveTitle = document.getElementById("cohort-curve-title");
  const selectedType = selectedFormulaType();
  const activeType = selectedType || definition?.bottomUp?.type;
  const isGenericCohortMatrix = activeType === "cohortMatrix";
  const isLegacyCohortMatrix = activeType === "cohortMatrixFromStartsAndAgeYoy";
  const shouldShow = Boolean(definition && (isGenericCohortMatrix || isLegacyCohortMatrix));
  form.classList.toggle("is-hidden", !shouldShow);

  if (!shouldShow) {
    form.classList.add("disabled");
    renderSelectOptions(startSourceInput, [], "", "Manual starting values");
    curveTypeInput.value = "survivalRate";
    startTimingInput.value = "currentPeriod";
    startsInputs.innerHTML = `<div class="empty-state">Select or create a metric first.</div>`;
    yoyInputs.innerHTML = `<div class="empty-state">Select or create a metric first.</div>`;
    openingInputs.innerHTML = "";
    openingCard.classList.add("is-hidden");
    preview.innerHTML = "";
    previewCard.classList.add("is-hidden");
    cohortStartsChart.clear();
    cohortYoyChart.clear();
    return;
  }

  form.classList.remove("disabled");
  state.cohortBuilderIsDirty = false;
  openingCard.classList.toggle("is-hidden", !isGenericCohortMatrix);
  previewCard.classList.remove("is-hidden");
  const starts = metricCohortStarts(definition.id);
  const legacyStartSourceId = metricCohortStartSourceId(definition.id);
  const sourceOptions = selectableInputMetricOptions(definition);
  const sourceOptionIds = new Set(sourceOptions.map(option => option.id));
  const rawSelectedSourceId = isGenericCohortMatrix ? metricInputIds(definition)[0] || "" : legacyStartSourceId;
  const selectedSourceId = sourceOptionIds.has(rawSelectedSourceId) ? rawSelectedSourceId : "";
  const ageCurve = isGenericCohortMatrix ? cohortAgeCurve(definition.id) : metricCohortAgeYoy(definition.id);
  const maxAge = Math.max(1, YEARS.length - 1);
  const activeCurveType = isGenericCohortMatrix ? cohortCurveType(definition) : "yoyGrowth";
  const ages = Array.from({ length: maxAge }, (_item, index) => activeCurveType === "survivalRate" ? index : index + 1);
  const context = selectedDimensionContextForMetric(definition.id);
  const contextFilters = context?.coordinate || {};
  const sourceSeries = isGenericCohortMatrix ? sourceSeriesForCohortMetric(definition.id, contextFilters) : null;
  const displayOpeningCohorts = isGenericCohortMatrix ? rollupOpeningCohorts(definition.id, contextFilters) : {};
  renderSelectOptions(startSourceInput, sourceOptions, selectedSourceId, "Manual starting values");
  curveTypeInput.closest("label").classList.toggle("is-hidden", !isGenericCohortMatrix);
  startTimingInput.closest("label").classList.toggle("is-hidden", !isGenericCohortMatrix);
  curveTypeInput.value = activeCurveType;
  startTimingInput.value = isGenericCohortMatrix ? cohortStartTiming(definition) : "currentPeriod";
  curveTitle.textContent = activeCurveType === "survivalRate"
    ? "Retention / Survival By Cohort Age"
    : "YoY % Change By Cohort Age";

  startsInputs.innerHTML = YEARS.map(year => `
    <label>
      <span>${year}</span>
      <input type="number" step="1" data-cohort-start-year="${year}" value="${isGenericCohortMatrix ? Number(sourceSeries[YEARS.indexOf(year)] || 0) : cohortStartValueForYear(starts, year)}" ${selectedSourceId ? "disabled" : ""}>
    </label>
  `).join("");

  yoyInputs.innerHTML = ages.map(age => `
    <label>
      <span>Age ${age}</span>
      <input type="number" step="0.1" data-cohort-yoy-age="${age}" value="${(cohortCurveValueForAge(ageCurve, age, 0) * 100).toFixed(1)}">
    </label>
  `).join("");
  openingInputs.innerHTML = isGenericCohortMatrix
    ? Array.from({ length: maxAge }, (_item, index) => index + 1).map(age => `
      <label>
        <span>Age ${age}</span>
        <input type="number" step="1" data-cohort-opening-age="${age}" value="${Number(displayOpeningCohorts[`age_${age}`] || 0)}">
      </label>
    `).join("")
    : "";
  renderCohortMatrixPreview(definition);

  if (!window.echarts) return;

  cohortStartsChart.setOption({
    animation: false,
    tooltip: { trigger: "axis", valueFormatter: value => formatMetricValue(definition, value, { precise: true }) },
    grid: { left: 10, right: 16, top: 16, bottom: 28, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: value => formatMetricValue(definition, value) } },
    series: [{
      name: "Starting Value",
      type: "line",
      data: YEARS.map((year, index) => isGenericCohortMatrix ? Number(sourceSeries[index] || 0) : cohortStartValueForYear(starts, year)),
      symbolSize: 7,
      lineStyle: { color: "#2f6f73", width: 3 },
      itemStyle: { color: "#2f6f73" },
    }],
  }, true);

  cohortYoyChart.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      valueFormatter: value => `${Number(value || 0).toFixed(1)}%`,
    },
    grid: { left: 10, right: 16, top: 16, bottom: 28, containLabel: true },
    xAxis: { type: "category", data: ages.map(age => `Age ${age}`) },
    yAxis: {
      type: "value",
      min: activeCurveType === "survivalRate" ? 0 : null,
      max: activeCurveType === "survivalRate" ? Math.max(100, ...ages.map(age => cohortCurveValueForAge(ageCurve, age, 0) * 100)) : undefined,
      axisLabel: { formatter: value => `${Number(value || 0).toFixed(0)}%` },
    },
    series: [{
      name: activeCurveType === "survivalRate" ? "Retention" : "YoY Change",
      type: "line",
      data: ages.map(age => cohortCurveValueForAge(ageCurve, age, 0) * 100),
      symbolSize: 7,
      lineStyle: { color: "#4f7fb8", width: 3 },
      itemStyle: { color: "#4f7fb8" },
    }],
  }, true);
}

function renderCohortMatrixPreview(definition) {
  const preview = document.getElementById("cohort-matrix-preview");
  preview.innerHTML = cohortMatrixPreviewMarkup(definition);
}

function syncCohortBuilderChartsFromForm(definition) {
  if (!definition || !window.echarts) return;
  const sourceId = document.getElementById("cohort-start-source-input")?.value || "";
  const context = selectedDimensionContextForMetric(definition.id);
  const sourceFilters = sourceId && metricDefinitions()[sourceId]
    ? relevantCoordinateFilters(sourceId, context)
    : {};
  const startSeries = sourceId && metricDefinitions()[sourceId]
    ? Object.keys(sourceFilters).length
      ? metricTopDownFilteredSeries(sourceId, sourceFilters)
      : metricTopDownSeries(sourceId)
    : YEARS.map(year => Number(document.querySelector(`[data-cohort-start-year="${year}"]`)?.value || 0));
  const curveInputs = [...document.querySelectorAll("[data-cohort-yoy-age]")];
  const ages = curveInputs.map(input => Number(input.dataset.cohortYoyAge));
  const curveData = curveInputs.map(input => Number(input.value || 0));
  const curveType = document.getElementById("cohort-curve-type-input")?.value || "yoyGrowth";

  cohortStartsChart.setOption({
    tooltip: { valueFormatter: value => formatMetricValue(definition, value, { precise: true }) },
    yAxis: { axisLabel: { formatter: value => formatMetricValue(definition, value) } },
    series: [{ data: startSeries }],
  });
  cohortYoyChart.setOption({
    xAxis: { data: ages.map(age => `Age ${age}`) },
    yAxis: {
      min: curveType === "survivalRate" ? 0 : undefined,
      max: curveType === "survivalRate" ? Math.max(100, ...curveData) : undefined,
      axisLabel: { formatter: value => `${Number(value || 0).toFixed(0)}%` },
    },
    series: [{
      name: curveType === "survivalRate" ? "Retention" : "YoY Change",
      data: curveData,
    }],
  });
}

function setTopToBottom(metricId, { direct = false } = {}) {
  if (!metricId) return;
  const context = selectedDimensionContextForMetric(metricId);
  const bottomUp = context
    ? direct
      ? metricDirectBottomUpSeries(metricId, context)
      : metricBottomUpFilteredSeries(metricId, context.coordinate)
    : direct
      ? metricDirectBottomUpSeries(metricId)
      : metricBottomUpSeries(metricId);
  if (!bottomUp) return;
  const metric = ensureMetricScenario(metricId);
  const points = YEARS.map((year, index) => [year, Number(bottomUp[index] || 0)]);
  if (context) {
    setMetricTopDownFilteredSeries(metricId, context.coordinate, bottomUp);
  } else {
    setMetricTopDownSeries(metricId, bottomUp);
    if (!metric.controlPoints) metric.controlPoints = {};
    metric.controlPoints[TOTAL_COORDINATE_KEY] = points;
  }
  render();
}

function newBlankModel() {
  state.model = createBlankModel();
  state.selectedMetricId = null;
  state.selectedDimensionContext = null;
  state.selectedDimensionId = null;
  state.topDownSheetExpandedKeys.clear();
  catalogSourceIsDirty = false;
  render();
}

function loadStarterModel() {
  state.model = createStarterModel();
  state.selectedMetricId = "profit";
  state.selectedDimensionContext = null;
  state.selectedDimensionId = Object.keys(dimensionDefinitions())[0] || null;
  state.topDownSheetExpandedKeys.clear();
  catalogSourceIsDirty = false;
  render();
}

function addMetric() {
  const labelInput = document.getElementById("metric-name-input");
  const unitInput = document.getElementById("metric-unit-input");
  const label = labelInput.value;
  if (!label?.trim()) return;
  const id = uniqueMetricId(label);
  metricDefinitions()[id] = createManualMetric({
    id,
    label: label.trim(),
    unit: unitInput.value,
    color: "#4f7fb8",
  });
  Object.values(scenarioCollection()).forEach(sourceScenario => {
    ensureMetricScenarioFor(sourceScenario, id);
  });
  state.selectedMetricId = id;
  state.selectedDimensionContext = null;
  labelInput.value = "";
  catalogSourceIsDirty = false;
  setSourceStatus(`Created ${label.trim()}.`, "success");
  render();
}

function applyBottomUpFormula() {
  const definition = selectedDefinition();
  if (!definition) return;

  const type = document.getElementById("formula-type-input").value;
  const inputs = type === "weightedAverage"
    ? [document.getElementById("formula-weight-metric")?.value || ""].filter(Boolean)
    : type === "ratio"
    ? [
      document.getElementById("formula-ratio-numerator")?.value || "",
      document.getElementById("formula-ratio-denominator")?.value || "",
    ].filter(Boolean)
    : [...document.querySelectorAll("#formula-input-list input:checked")].map(input => input.value);

  if (!type) {
    definition.bottomUp = null;
    definition.reconciliation = { enabled: false, tolerance: 1 };
    setSourceStatus(`${definition.label} is now manually defined.`, "success");
    catalogSourceIsDirty = false;
    render();
    return;
  }

  if (type === "difference" && inputs.length !== 2) {
    setSourceStatus("A difference formula needs exactly two inputs.", "error");
    return;
  }

  if (type === "sum" && inputs.length < 1) {
    setSourceStatus("A sum formula needs at least one input.", "error");
    return;
  }

  if (type === "product" && inputs.length < 2) {
    setSourceStatus("A product formula needs at least two inputs.", "error");
    return;
  }

  if (type === "ratio" && inputs.length !== 2) {
    setSourceStatus("A ratio formula needs a numerator and denominator.", "error");
    return;
  }

  if (type === "ratio" && inputs[0] === inputs[1]) {
    setSourceStatus("A ratio numerator and denominator should be different metrics.", "error");
    return;
  }

  if (type === "weightedAverage" && inputs.length !== 1) {
    setSourceStatus("A weighted average needs one weight metric.", "error");
    return;
  }

  if (type === "cohortAgeProduct" && inputs.length !== 2) {
    setSourceStatus("A cohort age product needs exactly two inputs: cohort stock first, age curve second.", "error");
    return;
  }

  if (type === "laggedRetentionFlow" && inputs.length !== 2) {
    setSourceStatus("A lagged retention flow needs exactly two inputs: stock metric first, retention ratio second.", "error");
    return;
  }

  if (type === "laggedMetric" && inputs.length !== 1) {
    setSourceStatus("A lagged metric needs exactly one source metric.", "error");
    return;
  }

  if (type === "cohortMatrixFromStartsAndAgeYoy") {
    inputs.length = 0;
  }

  if (type === "cohortMatrix") {
    applyCohortMatrixInputs();
    return;
  }

  if (type === "cumulativeNetFlow") {
    applyStockFlowInputs();
    return;
  }

  definition.bottomUp = type === "laggedMetric"
    ? { type, inputs, lag: definition.bottomUp?.lag || 1 }
    : type === "ratio"
      ? { type, inputs, numerator: inputs[0], denominator: inputs[1] }
      : type === "weightedAverage"
        ? { type, inputs, weightMetricId: inputs[0] }
      : { type, inputs };
  definition.topDown = type === "laggedMetric"
    ? laggedTopDownDefinition()
    : { type: "manualSeries", editable: true };
  definition.reconciliation = { enabled: true, tolerance: 1 };
  if (type === "laggedMetric") {
    sanitizeLaggedMetricScenario(definition.id);
  } else {
    const metric = ensureMetricScenario(definition.id);
    if (!metric.controlPoints) metric.controlPoints = {};
    metric.controlPoints[TOTAL_COORDINATE_KEY] = manualControlPoints(definition.id);
  }
  setSourceStatus(`${definition.label} formula updated: ${formulaText(definition.id)}.`, "success");
  catalogSourceIsDirty = false;
  render();
}

function applyStockFlowInputs() {
  const definition = selectedDefinition();
  if (!definition) return;

  const openingValue = Number(document.getElementById("stock-opening-value-input").value || 0);
  const inflowId = document.getElementById("stock-inflow-input").value;
  const outflowId = document.getElementById("stock-outflow-input").value;

  if (!inflowId || !outflowId) {
    setSourceStatus("Choose both additions and losses metrics for a cumulative net flow.", "error");
    return;
  }

  if (inflowId === outflowId) {
    setSourceStatus("Additions and losses should be different metrics.", "error");
    return;
  }

  definition.bottomUp = { type: "cumulativeNetFlow", inputs: [inflowId, outflowId] };
  definition.reconciliation = { enabled: true, tolerance: 1 };
  scenario().metrics[definition.id] = {
    ...metricScenario(definition.id),
    openingValue,
  };

  catalogSourceIsDirty = false;
  setSourceStatus(`${definition.label} stock flow updated: ${formulaText(definition.id)}.`, "success");
  render();
}

function applyLaggedOpeningValue() {
  const definition = selectedDefinition();
  if (!definition || !metricIsLagged(definition)) return;
  const metric = ensureMetricScenario(definition.id);
  const memberInputs = [...document.querySelectorAll("[data-lagged-opening-member-id]")];
  if (memberInputs.length) {
    const nextOpeningValue = {
      ...(metric.openingValue && typeof metric.openingValue === "object" && !Array.isArray(metric.openingValue)
        ? metric.openingValue
        : {}),
    };
    memberInputs.forEach(input => {
      nextOpeningValue[input.dataset.laggedOpeningMemberId] = Number(input.value || 0);
    });
    metric.openingValue = nextOpeningValue;
  } else {
    const input = document.querySelector("[data-lagged-opening-value]");
    metric.openingValue = Number(input?.value || 0);
  }
  clearCalculationCache();
  setSourceStatus(`${definition.label} opening value updated.`, "success");
  render();
}

function applyCohortMatrixInputs() {
  const definition = selectedDefinition();
  if (!definition) return;

  const formulaType = selectedFormulaType();
  const isGenericCohortMatrix = formulaType === "cohortMatrix";
  const cohortStartSourceId = document.getElementById("cohort-start-source-input").value;
  if (cohortStartSourceId === definition.id) {
    setSourceStatus("A metric cannot source starting cohort values from itself.", "error");
    return;
  }
  const cohortStarts = {};
  document.querySelectorAll("[data-cohort-start-year]").forEach(input => {
    cohortStarts[input.dataset.cohortStartYear] = Number(input.value || 0);
  });

  const cohortAgeCurve = {};
  document.querySelectorAll("[data-cohort-yoy-age]").forEach(input => {
    cohortAgeCurve[input.dataset.cohortYoyAge] = Number(input.value || 0) / 100;
  });

  const openingCohorts = {};
  document.querySelectorAll("[data-cohort-opening-age]").forEach(input => {
    openingCohorts[`age_${input.dataset.cohortOpeningAge}`] = Number(input.value || 0);
  });

  if (isGenericCohortMatrix) {
    const context = selectedDimensionContextForMetric(definition.id);
    scenario().metrics[definition.id] = {
      ...metricScenario(definition.id),
      cohortStarts,
      openingCohorts: serializeOpeningCohortsForSave(definition.id, openingCohorts, context?.coordinate || {}),
      cohortAgeCurve,
    };
    definition.bottomUp = {
      type: "cohortMatrix",
      inputs: cohortStartSourceId ? [cohortStartSourceId] : [],
      curveType: document.getElementById("cohort-curve-type-input").value || "survivalRate",
      startTiming: document.getElementById("cohort-start-timing-input").value || "currentPeriod",
      rollup: "calendarYear",
    };
    definition.reconciliation = { enabled: true, tolerance: 1 };
    clearCalculationCache();
    catalogSourceIsDirty = false;
    state.cohortBuilderIsDirty = false;
    setSourceStatus(`${definition.label} cohort matrix inputs updated.`, "success");
    render();
    return;
  }

  scenario().metrics[definition.id] = {
    ...metricScenario(definition.id),
    cohortStarts,
    cohortStartSourceId: cohortStartSourceId || null,
    cohortAgeYoy: cohortAgeCurve,
  };

  if (definition.bottomUp?.type !== "cohortMatrixFromStartsAndAgeYoy") {
    definition.bottomUp = { type: "cohortMatrixFromStartsAndAgeYoy", inputs: [] };
    definition.reconciliation = { enabled: true, tolerance: 1 };
  }

  catalogSourceIsDirty = false;
  setSourceStatus(`${definition.label} cohort matrix inputs updated.`, "success");
  render();
}

function applyCatalogSource() {
  const source = document.getElementById("catalog-source");
  const previousYears = [...YEARS];
  try {
    const parsed = JSON.parse(source.value);
    const imported = normalizeImportedModel(parsed);
    state.model = imported;
    const importedSelection = parsed.selectedMetricId;
    state.selectedMetricId = metricDefinitions()[importedSelection]
      ? importedSelection
      : metricDefinitions()[state.selectedMetricId]
        ? state.selectedMetricId
        : Object.keys(metricDefinitions())[0] || null;
    state.selectedDimensionContext = parsed.selectedDimensionContext || null;
    state.selectedDimensionId = Object.keys(dimensionDefinitions())[0] || null;
    state.topDownSheetExpandedKeys.clear();
    catalogSourceIsDirty = false;
    setSourceStatus("Applied catalog JSON.", "success");
    render();
  } catch (error) {
    YEARS = previousYears;
    setEngineYears(YEARS);
    setSourceStatus(error.message, "error");
  }
}

async function copyCatalogSource() {
  const source = document.getElementById("catalog-source");
  try {
    await navigator.clipboard.writeText(source.value);
    setSourceStatus("Copied catalog JSON.", "success");
  } catch (_error) {
    source.select();
    setSourceStatus("Catalog JSON selected. Copy it manually.", "error");
  }
}

function renderScenarioControls() {
  const select = document.getElementById("scenario-select");
  if (!select) return;
  const activeId = activeScenarioId();
  select.innerHTML = Object.values(scenarioCollection()).map(item => `
    <option value="${escapeHtml(item.id)}" ${item.id === activeId ? "selected" : ""}>
      ${escapeHtml(item.name || item.id)}
    </option>
  `).join("");

  const controls = document.getElementById("scenario-role-controls");
  if (!controls) return;
  const references = referenceScenarios();
  const roles = scenarioRoles();
  const summary = activeScenarioReconciliationSummary();
  const publishDisabled = !summary.matched;
  const roleLabel = role => {
    const reference = references[role];
    if (!reference) return "Unset";
    const sourceName = reference.sourceScenarioName || reference.sourceScenarioId || "Snapshot";
    return `${sourceName}`;
  };
  controls.innerHTML = `
    <span class="scenario-role-pill">Working <strong>${escapeHtml(scenario().name || activeId)}</strong></span>
    <span class="scenario-role-pill">BAU <strong>${escapeHtml(roleLabel("bau"))}</strong></span>
    <span class="scenario-role-pill">Target <strong>${escapeHtml(roleLabel("target"))}</strong></span>
    <span class="scenario-role-status ${summary.matched ? "matched" : "open"}">
      ${summary.matched ? "Reconciled" : `${summary.openMetricCount} mismatch${summary.openMetricCount === 1 ? "" : "es"}`}
    </span>
    <button id="set-reference-bau" type="button" ${publishDisabled || referenceSnapshotMatchesCurrent("bau") ? "disabled" : ""}>Set as BAU</button>
    <button id="set-reference-target" type="button" ${publishDisabled || referenceSnapshotMatchesCurrent("target") ? "disabled" : ""}>Set as Target</button>
  `;
  roles.working = activeId;
}

function renderTimeRangeControls() {
  const startInput = document.getElementById("start-year-input");
  const endInput = document.getElementById("end-year-input");
  if (!startInput || !endInput) return;
  startInput.value = YEARS[0] || DEFAULT_YEARS[0];
  endInput.value = YEARS[YEARS.length - 1] || DEFAULT_YEARS[DEFAULT_YEARS.length - 1];
}

function applyTimeRange() {
  const startInput = document.getElementById("start-year-input");
  const endInput = document.getElementById("end-year-input");
  try {
    const nextYears = yearsFromRange(startInput.value, endInput.value);
    const changed = migrateModelYears(nextYears);
    catalogSourceIsDirty = false;
    setSourceStatus(changed
      ? `Updated model years to ${YEARS[0]}-${YEARS[YEARS.length - 1]}.`
      : "Model years were already up to date.",
    changed ? "success" : "");
    render();
  } catch (error) {
    setSourceStatus(error.message, "error");
  }
}

function switchScenario(scenarioId) {
  if (!scenarioCollection()[scenarioId]) return;
  state.model.activeScenarioId = scenarioId;
  state.selectedDimensionContext = null;
  state.topDownSheetExpandedKeys.clear();
  catalogSourceIsDirty = false;
  setSourceStatus(`Switched to ${scenario().name || scenario().id}.`, "success");
  render();
}

function saveScenario() {
  const active = scenario();
  scenarioCollection()[active.id] = {
    ...scenarioSnapshot(active),
    updatedAt: new Date().toISOString(),
  };
  catalogSourceIsDirty = false;
  setSourceStatus(`Saved ${scenario().name || scenario().id}.`, "success");
  render();
}

function saveScenarioAsNew() {
  const source = scenarioSnapshot();
  const id = uniqueScenarioId();
  const timestamp = new Date().toISOString();
  scenarioCollection()[id] = {
    ...clone(source),
    id,
    name: uniqueScenarioName(),
    sourceScenarioId: source.id,
    sourceScenarioName: source.name || source.id,
    createdAt: timestamp,
    updatedAt: timestamp,
  };
  state.model.activeScenarioId = id;
  catalogSourceIsDirty = false;
  setSourceStatus(`Saved as ${scenario().name}.`, "success");
  render();
}

function render() {
  syncActiveYears();
  clearCalculationCache();
  renderScenarioControls();
  renderTimeRangeControls();
  renderDeltaReviewPanel();
  renderMetricWorkspacePanel();
  renderMetricList();
  renderDimensionCatalog();
  renderCanvas();
  renderDetail();
  renderCatalogSource();
  const definition = selectedDefinition();
  document.getElementById("set-top-to-bottom").disabled = !definition || !activeBottomUpSeries(definition);
  requestAnimationFrame(renderCanvasLines);
}

document.body.addEventListener("click", event => {
  const jsonToggle = event.target.closest("[data-json-explorer-path]");
  if (jsonToggle) {
    const path = jsonToggle.dataset.jsonExplorerPath || "";
    if (state.jsonExplorerExpandedPaths.has(path)) {
      state.jsonExplorerExpandedPaths.delete(path);
    } else {
      state.jsonExplorerExpandedPaths.add(path);
    }
    renderJsonExplorerFromText(document.getElementById("catalog-source").value);
    return;
  }

  const dimensionButton = event.target.closest("[data-catalog-dimension-id]");
  if (dimensionButton) {
    state.selectedDimensionId = dimensionButton.dataset.catalogDimensionId;
    renderDimensionCatalog();
    return;
  }

  const memberTotalButton = event.target.closest("[data-member-total]");
  if (memberTotalButton) {
    state.selectedMetricId = memberTotalButton.dataset.memberTotal;
    state.selectedDimensionContext = null;
    render();
    return;
  }

  const memberContextButton = event.target.closest("[data-member-context]");
  if (memberContextButton) {
    state.selectedMetricId = memberContextButton.dataset.memberContext;
    state.selectedDimensionContext = {
      metricId: memberContextButton.dataset.memberContext,
      coordinate: { [memberContextButton.dataset.dimensionId]: memberContextButton.dataset.memberId },
    };
    render();
    return;
  }

  const nodeButton = event.target.closest("[data-node-key]");
  if (nodeButton) {
    selectNodeKey(nodeButton.dataset.nodeKey);
    render();
    return;
  }

  const workspaceLockButton = event.target.closest("[data-workspace-lock-key]");
  if (workspaceLockButton) {
    toggleWorkspaceLock(workspaceLockButton.dataset.workspaceLockKey);
    renderMetricWorkspacePanel();
    return;
  }

  const metricButton = event.target.closest("[data-metric-id]");
  if (metricButton) {
    selectMetric(metricButton.dataset.metricId);
    render();
    return;
  }
  const deltaMetricButton = event.target.closest("[data-delta-metric-id]");
  if (deltaMetricButton) {
    selectMetric(deltaMetricButton.dataset.deltaMetricId);
    render();
    return;
  }
  const canvasViewButton = event.target.closest("[data-canvas-view-mode]");
  if (canvasViewButton) {
    state.canvasViewMode = canvasViewButton.dataset.canvasViewMode === "chart" ? "chart" : "number";
    renderCanvas();
    requestAnimationFrame(renderCanvasLines);
    return;
  }
  const workspaceModeButton = event.target.closest("[data-workspace-mode]");
  if (workspaceModeButton) {
    state.metricWorkspaceMode = workspaceModeButton.dataset.workspaceMode || "total";
    renderMetricWorkspacePanel();
    return;
  }
  const dimensionGroupToolButton = event.target.closest("[data-dimension-group-tool]");
  if (dimensionGroupToolButton) {
    state.dimensionGroupTool = dimensionGroupToolButton.dataset.dimensionGroupTool === "blue" ? "blue" : "red";
    renderMetricWorkspacePanel();
    return;
  }
  const dimensionMixViewButton = event.target.closest("[data-dimension-mix-view]");
  if (dimensionMixViewButton) {
    state.dimensionMixView = dimensionMixViewButton.dataset.dimensionMixView === "total" ? "total" : "pctTotal";
    renderMetricWorkspacePanel();
    return;
  }
  const dimensionGroupLineLayerButton = event.target.closest("[data-dimension-group-line-layer]");
  if (dimensionGroupLineLayerButton) {
    const layer = dimensionGroupLineLayerButton.dataset.dimensionGroupLineLayer;
    if (layer === "grouped") state.dimensionGroupLinesShowGrouped = !state.dimensionGroupLinesShowGrouped;
    if (layer === "ungrouped") state.dimensionGroupLinesShowUngrouped = !state.dimensionGroupLinesShowUngrouped;
    if (!state.dimensionGroupLinesShowGrouped && !state.dimensionGroupLinesShowUngrouped) {
      if (layer === "grouped") state.dimensionGroupLinesShowGrouped = true;
      if (layer === "ungrouped") state.dimensionGroupLinesShowUngrouped = true;
    }
    renderMetricWorkspacePanel();
    return;
  }
  const dimensionGroupLineChartTypeButton = event.target.closest("[data-dimension-group-line-chart-type]");
  if (dimensionGroupLineChartTypeButton) {
    state.dimensionGroupLinesChartType = dimensionGroupLineChartTypeButton.dataset.dimensionGroupLineChartType === "bar" ? "bar" : "line";
    renderMetricWorkspacePanel();
    return;
  }
  const dimensionGroupMemberButton = event.target.closest("[data-dimension-group-member]");
  if (dimensionGroupMemberButton) {
    const definition = selectedDefinition();
    const mixContext = dimensionMixContext(definition);
    if (mixContext) {
      const memberId = dimensionGroupMemberButton.dataset.dimensionGroupMember;
      const grouping = dimensionGroupingForContext(mixContext);
      const activeGroup = state.dimensionGroupTool === "blue" ? "blue" : "red";
      const currentGroup = grouping.red.includes(memberId) ? "red" : grouping.blue.includes(memberId) ? "blue" : "other";
      setDimensionMemberGroup(mixContext, memberId, currentGroup === activeGroup ? "other" : activeGroup);
      catalogSourceIsDirty = false;
      renderMetricWorkspacePanel();
      renderCatalogSource();
    }
    return;
  }
  if (event.target.closest("#set-top-to-bottom")) {
    setTopToBottom(state.selectedMetricId);
    return;
  }
  const workspaceSetParentButton = event.target.closest("[data-workspace-set-parent]");
  if (workspaceSetParentButton) {
    setTopToBottom(workspaceSetParentButton.dataset.workspaceSetParent, { direct: true });
    return;
  }
  const workspaceFitChildrenButton = event.target.closest("[data-workspace-fit-children]");
  if (workspaceFitChildrenButton) {
    fitChildrenToParent(workspaceFitChildrenButton.dataset.workspaceFitChildren);
    return;
  }
  const workspaceCoordinateButton = event.target.closest("[data-workspace-coordinate-metric]");
  if (workspaceCoordinateButton) {
    state.selectedMetricId = workspaceCoordinateButton.dataset.workspaceCoordinateMetric;
    const coordinate = workspaceCoordinateButton.dataset.workspaceCoordinateJson
      ? JSON.parse(workspaceCoordinateButton.dataset.workspaceCoordinateJson)
      : parseCoordinateKey(workspaceCoordinateButton.dataset.workspaceCoordinateKey || "");
    state.selectedDimensionContext = {
      metricId: state.selectedMetricId,
      coordinate,
    };
    render();
    return;
  }
  if (event.target.closest("[data-focus-cohort-builder]")) {
    document.getElementById("cohort-matrix-builder-form")?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  if (event.target.closest("#new-blank-model")) {
    newBlankModel();
    return;
  }
  if (event.target.closest("#load-starter-model")) {
    loadStarterModel();
    return;
  }
  if (event.target.closest("#save-scenario")) {
    saveScenario();
    return;
  }
  if (event.target.closest("#save-scenario-as-new")) {
    saveScenarioAsNew();
    return;
  }
  if (event.target.closest("#set-reference-bau")) {
    setReferenceScenario("bau");
    return;
  }
  if (event.target.closest("#set-reference-target")) {
    setReferenceScenario("target");
    return;
  }
  if (event.target.closest("#add-metric")) {
    document.getElementById("metric-name-input").focus();
    return;
  }
  if (event.target.closest("#new-dimension")) {
    newDimensionDraft();
    return;
  }
  const sheetRollupButton = event.target.closest("[data-sheet-rollup]");
  if (sheetRollupButton) {
    toggleSheetRow(sheetRollupButton.dataset.sheetRollup, parseCoordinateKey(sheetRollupButton.dataset.rollupKey || ""));
    render();
    return;
  }
  const topDownViewButton = event.target.closest("[data-top-down-view]");
  if (topDownViewButton) {
    setTopDownEditorView(topDownViewButton.dataset.topDownView);
    render();
    return;
  }
  if (event.target.closest("#assign-metric-dimension")) {
    const definition = selectedDefinition();
    if (!definition) return;
    const dimensionIds = [...document.querySelectorAll("[data-metric-dimension-level]")]
      .sort((left, right) => Number(left.dataset.metricDimensionLevel || 0) - Number(right.dataset.metricDimensionLevel || 0))
      .map(input => input.value)
      .filter(Boolean);
    assignMetricDimensions(definition.id, dimensionIds);
    catalogSourceIsDirty = false;
    setSourceStatus(`${definition.label} dimension assignment updated.`, "success");
    render();
    return;
  }
  if (event.target.closest("#apply-catalog-source")) {
    applyCatalogSource();
    return;
  }
  if (event.target.closest("#copy-catalog-source")) {
    copyCatalogSource();
    return;
  }
  const deleteAssumptionButton = event.target.closest("[data-delete-assumption]");
  if (deleteAssumptionButton) {
    deleteAssumption(deleteAssumptionButton.dataset.deleteAssumption);
  }
});

document.getElementById("metric-create-form").addEventListener("submit", event => {
  event.preventDefault();
  addMetric();
});

document.getElementById("dimension-form").addEventListener("submit", event => {
  event.preventDefault();
  saveDimensionDefinition();
});

document.getElementById("assumption-form").addEventListener("submit", event => {
  event.preventDefault();
  addAssumption();
});

document.getElementById("time-range-form").addEventListener("submit", event => {
  event.preventDefault();
  applyTimeRange();
});

document.getElementById("bottom-up-form").addEventListener("submit", event => {
  event.preventDefault();
  applyBottomUpFormula();
});

document.getElementById("cohort-matrix-builder-form").addEventListener("submit", event => {
  event.preventDefault();
  applyCohortMatrixInputs();
});

document.getElementById("cohort-matrix-builder-form").addEventListener("input", event => {
  if (!event.target.closest("[data-cohort-start-year], [data-cohort-yoy-age], [data-cohort-opening-age]")) return;
  state.cohortBuilderIsDirty = true;
  const definition = selectedDefinition();
  syncCohortBuilderChartsFromForm(definition);
  renderCohortMatrixPreview(definition);
});

document.getElementById("stock-flow-builder-form").addEventListener("submit", event => {
  event.preventDefault();
  applyStockFlowInputs();
});

document.getElementById("lagged-opening-form").addEventListener("submit", event => {
  event.preventDefault();
  applyLaggedOpeningValue();
});

document.getElementById("catalog-source").addEventListener("input", () => {
  catalogSourceIsDirty = true;
  renderJsonExplorerFromText(document.getElementById("catalog-source").value);
  setSourceStatus("Catalog JSON has unapplied edits.");
});

document.body.addEventListener("change", event => {
  const scenarioSelect = event.target.closest("#scenario-select");
  if (scenarioSelect) {
    switchScenario(scenarioSelect.value);
    return;
  }

  if (event.target.closest("#cohort-start-source-input")) {
    state.cohortBuilderIsDirty = true;
    const definition = selectedDefinition();
    syncCohortBuilderChartsFromForm(definition);
    renderCohortMatrixPreview(definition);
    return;
  }

  if (event.target.closest("#cohort-curve-type-input")) {
    const definition = selectedDefinition();
    if (definition?.bottomUp?.type === "cohortMatrix") {
      definition.bottomUp.curveType = event.target.value || "survivalRate";
    }
    renderCohortMatrixBuilder(definition);
    return;
  }

  if (event.target.closest("#cohort-start-timing-input")) {
    const definition = selectedDefinition();
    if (definition?.bottomUp?.type === "cohortMatrix") {
      definition.bottomUp.startTiming = event.target.value || "currentPeriod";
    }
    renderCohortMatrixBuilder(definition);
    return;
  }

  const coordinateFilter = event.target.closest("[data-coordinate-filter]");
  if (coordinateFilter) {
    const metricId = coordinateFilter.dataset.coordinateFilter;
    const coordinate = {};
    document.querySelectorAll("[data-coordinate-filter]").forEach(select => {
      if (select.dataset.coordinateFilter !== metricId) return;
      if (select.value) coordinate[select.dataset.dimensionId] = select.value;
    });
    state.selectedMetricId = metricId;
    state.selectedDimensionContext = Object.keys(coordinate).length
      ? { metricId, coordinate }
      : null;
    render();
  }
});

document.getElementById("formula-type-input").addEventListener("change", () => {
  const definition = selectedDefinition();
  renderFormulaInputList(definition);
  renderCohortMatrixBuilder(definition);
  renderStockFlowBuilder(definition);
  renderTopDownNapkinEditor(null);
  renderMetricWorkspacePanel();
  renderLaggedOpeningEditor(definition);
});

window.addEventListener("resize", () => {
  chart.resize();
  topDownNapkinChart?.chart?.resize();
  workspaceNapkinChart?.chart?.resize();
  dimensionMixChart?.chart?.resize();
  workspaceDimensionMixChart?.chart?.resize();
  metricWorkspaceCharts.forEach(item => item.resize());
  dimensionBreakdownChart.resize();
  cohortStartsChart.resize();
  cohortYoyChart.resize();
  renderCanvasLines();
});

render();
