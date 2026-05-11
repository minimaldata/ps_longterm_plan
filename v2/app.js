import {
  YEARS,
  createBlankModel,
  createManualMetric,
  createStarterModel,
  defaultSeries,
} from "./metric-catalog.js";

const state = {
  model: createBlankModel(),
  selectedMetricId: null,
  selectedDimensionContext: null,
  selectedDimensionId: null,
  topDownEditorView: "chart",
  topDownSheetExpandedKeys: new Set(),
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
const TOTAL_COORDINATE_KEY = "";
const UNASSIGNED_MEMBER_ID = "unassigned";

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
let dimensionMixChart = null;
let dimensionMixMetricId = null;
let dimensionMixIsSyncing = false;
let dimensionMixHasPendingCommit = false;
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
    const memberId = rawCoordinate[dimensionId];
    if (!memberId) return;
    const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
    if (member) coordinate[dimensionId] = memberId;
  });
  if (!Object.keys(coordinate).length) return null;
  const key = coordinateKey(coordinate);
  const entries = Object.entries(coordinate).map(([dimensionId, memberId]) => {
    const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
    return { dimensionId, memberId, member };
  });
  return {
    metricId,
    coordinate,
    coordinateKey: key,
    entries,
    label: coordinateLabelForMetric(metricId, key),
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

function coordinateKey(coordinate = {}) {
  return Object.entries(coordinate)
    .filter(([_dimensionId, memberId]) => memberId !== undefined && memberId !== null && memberId !== "")
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([dimensionId, memberId]) => `${encodeURIComponent(dimensionId)}=${encodeURIComponent(memberId)}`)
    .join("|");
}

function parseCoordinateKey(key) {
  if (!key) return {};
  return String(key).split("|").reduce((coordinate, part) => {
    const [dimensionId, memberId] = part.split("=");
    if (!dimensionId || memberId === undefined) return coordinate;
    coordinate[decodeURIComponent(dimensionId)] = decodeURIComponent(memberId);
    return coordinate;
  }, {});
}

function memberCoordinateKey(dimensionId, memberId) {
  return coordinateKey({ [dimensionId]: memberId });
}

function coordinateMatches(key, filters = {}) {
  const coordinate = parseCoordinateKey(key);
  return Object.entries(filters).every(([dimensionId, memberId]) => coordinate[dimensionId] === memberId);
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

function sumSeries(left = defaultSeries(), right = defaultSeries()) {
  return YEARS.map((_year, index) => Number(left[index] || 0) + Number(right[index] || 0));
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
    .filter(([key]) => coordinateMatches(key, filters));
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

function metricTopDownFilteredSeries(metricId, filters = {}) {
  return cachedCalculation(
    `topDownFiltered:${metricId}:${coordinateKey(filters)}`,
    () => aggregateSeriesMap(metricId, filters)
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
  const filterKey = coordinateKey(filters);
  const coordinateKeys = metricCoordinateKeys(metricId)
    .filter(key => coordinateMatches(key, filters));

  if (!filterKey || !Object.keys(filters).length) {
    if (!metric.controlPoints) metric.controlPoints = {};
    metric.controlPoints[TOTAL_COORDINATE_KEY] = YEARS.map((year, index) => [year, Number(series[index] || 0)]);
    setMetricTopDownSeries(metricId, series);
    return;
  }

  if (coordinateKeys.length === 1 && coordinateKeys[0] === filterKey) {
    setMetricTopDownCoordinateControlPoints(metricId, filterKey, series);
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
  metric.controlPoints[filterKey] = YEARS.map((year, index) => [year, Number(series[index] || 0)]);
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
  return cachedCalculation(`topDown:${metricId}`, () => aggregateSeriesMap(metricId));
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
  const orderedDimensionIds = [
    ...metricDimensionIds(metricId),
    ...Object.keys(coordinate).filter(dimensionId => !metricDimensionIds(metricId).includes(dimensionId)),
  ];
  return orderedDimensionIds
    .filter(dimensionId => coordinate[dimensionId])
    .map(dimensionId => {
      const memberId = coordinate[dimensionId];
      const member = dimensionMembers(dimensionId).find(item => item.id === memberId);
      const levelIndex = metricDimensionIds(metricId).indexOf(dimensionId);
      const levelPrefix = levelIndex >= 0 ? `L${levelIndex + 1} ` : "";
      return `${levelPrefix}${dimensionDefinitions()[dimensionId]?.label || dimensionId}: ${member?.label || memberId}`;
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

function cohortValueForYear(row, year) {
  if (!row || typeof row !== "object") return 0;
  return Number(row[String(year)] || 0);
}

function cohortStartValueForYear(starts, year) {
  return Number(starts[String(year)] || 0);
}

function cohortYoyValueForAge(ageYoy, age) {
  if (age <= 0) return 0;
  if (Array.isArray(ageYoy)) return Number(ageYoy[age - 1] || 0);
  return Number(ageYoy[String(age)] || 0);
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

function metricInputIds(definition) {
  return (definition?.bottomUp?.inputs || []).map(input => {
    if (typeof input === "string") return input;
    return input?.metricId;
  }).filter(Boolean);
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
  const context = parsed.memberId
    ? {
      metricId: parsed.metricId,
      coordinate: { [parsed.dimensionId]: parsed.memberId },
      dimensionId: parsed.dimensionId,
      memberId: parsed.memberId,
      member: dimensionMembers(parsed.dimensionId).find(item => item.id === parsed.memberId),
    }
    : null;
  if (metricIsLagged(definition)) {
    const laggedLines = laggedComparisonLines(definition, context);
    if (laggedLines?.topDown) return laggedLines.topDown;
  }
  return context
    ? metricTopDownFilteredSeries(parsed.metricId, context.coordinate)
    : metricTopDownSeries(parsed.metricId);
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

  if (definition.bottomUp.type === "sum") {
    return inputs.reduce((sum, inputId) => sum + metricValueAtIndex(inputId, index, nextStack, memberContext), 0);
  }

  if (definition.bottomUp.type === "product") {
    return inputs.reduce((product, inputId) => product * metricValueAtIndex(inputId, index, nextStack, memberContext), 1);
  }

  if (definition.bottomUp.type === "difference") {
    const [leftId, rightId] = inputs;
    return metricValueAtIndex(leftId, index, nextStack, memberContext) - metricValueAtIndex(rightId, index, nextStack, memberContext);
  }

  if (definition.bottomUp.type === "cumulativeNetFlow") {
    const [inflowId, outflowId] = inputs;
    let runningValue = metricOpeningValue(metricId, memberId);
    for (let currentIndex = 0; currentIndex <= index; currentIndex += 1) {
      runningValue += metricValueAtIndex(inflowId, currentIndex, nextStack, memberContext) - metricValueAtIndex(outflowId, currentIndex, nextStack, memberContext);
    }
    return runningValue;
  }

  if (definition.bottomUp.type === "laggedMetric") {
    const [sourceMetricId] = inputs;
    const sourceIndex = index - lagForDefinition(definition);
    if (sourceIndex < 0) return laggedOpeningValue(definition, sourceMetricId, memberId);
    return metricValueAtIndex(sourceMetricId, sourceIndex, nextStack, memberContext);
  }

  if (definition.bottomUp.type === "laggedRetentionFlow") {
    const [stockMetricId, retentionMetricId] = inputs;
    const priorStock = index <= 0
      ? metricOpeningValue(metricId, memberId) || metricOpeningValue(stockMetricId, memberId)
      : metricValueAtIndex(stockMetricId, index - 1, nextStack, memberContext);
    return priorStock * metricValueAtIndex(retentionMetricId, index, nextStack, memberContext);
  }

  const bottomUpSeries = metricBottomUpSeries(metricId);
  return Number(bottomUpSeries?.[index] || 0);
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

    const coordinateKeys = metricCoordinateKeys(metricId);
    if (metricDimensionIds(metricId).length && coordinateKeys.length) {
      const seriesByCoordinate = Object.fromEntries(coordinateKeys.map(key => [
        key,
        YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { coordinate: parseCoordinateKey(key) })),
      ]));
      return aggregateMemberSeries(metricId, seriesByCoordinate);
    }

    return YEARS.map((_year, index) => metricValueAtIndex(metricId, index));
  });
}

function metricBottomUpMemberSeries(metricId, dimensionId, memberId) {
  return cachedCalculation(`bottomUpMember:${metricId}:${dimensionId}:${memberId}`, () => {
    const definition = metricDefinitions()[metricId];
    if (!definition?.bottomUp || !metricUsesDimension(metricId, dimensionId)) return null;
    return YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { coordinate: { [dimensionId]: memberId } }));
  });
}

function metricBottomUpFilteredSeries(metricId, filters = {}) {
  return cachedCalculation(`bottomUpFiltered:${metricId}:${coordinateKey(filters)}`, () => {
    const definition = metricDefinitions()[metricId];
    if (!definition?.bottomUp) return null;
    return YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { coordinate: filters }));
  });
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

function childIds(metricId) {
  const definition = metricDefinitions()[metricId];
  if (definition?.bottomUp?.type === "laggedMetric") return [];
  if (definition?.bottomUp?.type === "laggedRetentionFlow") return Array.from(new Set(metricInputIds(definition).slice(1)));
  const ids = metricInputIds(definition);
  if (definition?.bottomUp?.type === "cohortMatrixFromStartsAndAgeYoy") {
    const startSourceId = metricCohortStartSourceId(metricId);
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
    dimensions: dimensionDefinitions(),
    metricDefinitions: metricDefinitions(),
    scenarios: scenariosSnapshot(),
  }, null, 2);
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

function setSourceStatus(message, tone = "") {
  const status = document.getElementById("catalog-source-status");
  status.textContent = message;
  status.className = `source-status ${tone}`;
}

function renderCatalogSource() {
  const source = document.getElementById("catalog-source");
  if (document.activeElement === source && catalogSourceIsDirty) return;
  source.value = sourceSnapshot();
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

  return {
    name: rawModel.name || "Custom Model",
    dimensions: rawModel.dimensions || {},
    metricDefinitions: definitions,
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

function renderNode(nodeKey) {
  const parsed = parseNodeKey(nodeKey);
  const definition = metricDefinitions()[parsed.metricId];
  const isMember = Boolean(parsed.memberId);
  const member = isMember
    ? dimensionMembers(parsed.dimensionId).find(item => item.id === parsed.memberId)
    : null;
  const series = canvasNodeSeries(nodeKey);
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
      <strong>${compactCurrency(series[series.length - 1])}</strong>
      <small>${nodeCaption}</small>
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
  renderTopDownNapkinEditor(definition);
  renderLaggedOpeningEditor(definition);
  renderDimensionBreakdown(metricIsLagged(definition) ? null : definition);
  renderDimensionMixEditor(metricIsLagged(definition) ? null : definition);
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
  renderTopDownNapkinEditor(definition);
  renderLaggedOpeningEditor(definition);
  renderDimensionBreakdown(null);
  renderDimensionMixEditor(null);
  renderDefinitionComparisonChart(definition);
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
    renderDimensionMixEditor(definition);
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

  if (type === "cohortMatrixFromStartsAndAgeYoy" || type === "cumulativeNetFlow") {
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
  const yMax = Math.max(10, ...points.map(point => Number(point[1] || 0))) * 1.25;
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
  return Object.values(metricDefinitions()).filter(candidate => !blockedInputs.has(candidate.id));
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
  const startsInputs = document.getElementById("cohort-starts-inputs");
  const yoyInputs = document.getElementById("cohort-yoy-inputs");
  const shouldShow = Boolean(definition && (definition.bottomUp?.type === "cohortMatrixFromStartsAndAgeYoy" || selectedFormulaType() === "cohortMatrixFromStartsAndAgeYoy"));
  form.classList.toggle("is-hidden", !shouldShow);

  if (!shouldShow) {
    form.classList.add("disabled");
    renderSelectOptions(startSourceInput, [], "", "Manual starting values");
    startsInputs.innerHTML = `<div class="empty-state">Select or create a metric first.</div>`;
    yoyInputs.innerHTML = `<div class="empty-state">Select or create a metric first.</div>`;
    cohortStartsChart.clear();
    cohortYoyChart.clear();
    return;
  }

  form.classList.remove("disabled");
  const starts = metricCohortStarts(definition.id);
  const startSourceId = metricCohortStartSourceId(definition.id);
  const ageYoy = metricCohortAgeYoy(definition.id);
  const sourceOptions = selectableInputMetricOptions(definition);
  const maxAge = Math.max(1, YEARS.length - 1);
  const ages = Array.from({ length: maxAge }, (_item, index) => index + 1);
  renderSelectOptions(startSourceInput, sourceOptions, startSourceId, "Manual starting values");

  startsInputs.innerHTML = YEARS.map(year => `
    <label>
      <span>${year}</span>
      <input type="number" step="1" data-cohort-start-year="${year}" value="${cohortStartValueForYear(starts, year)}" ${startSourceId ? "disabled" : ""}>
    </label>
  `).join("");

  yoyInputs.innerHTML = ages.map(age => `
    <label>
      <span>Age ${age}</span>
      <input type="number" step="0.1" data-cohort-yoy-age="${age}" value="${(cohortYoyValueForAge(ageYoy, age) * 100).toFixed(1)}">
    </label>
  `).join("");

  if (!window.echarts) return;

  cohortStartsChart.setOption({
    animation: false,
    tooltip: { trigger: "axis", valueFormatter: value => compactCurrencyTooltip(value) },
    grid: { left: 10, right: 16, top: 16, bottom: 28, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: compactCurrency } },
    series: [{
      name: "Starting Value",
      type: "line",
      data: YEARS.map(year => cohortStartValueForYear(starts, year)),
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
    yAxis: { type: "value", axisLabel: { formatter: value => `${Number(value || 0).toFixed(0)}%` } },
    series: [{
      name: "YoY Change",
      type: "line",
      data: ages.map(age => cohortYoyValueForAge(ageYoy, age) * 100),
      symbolSize: 7,
      lineStyle: { color: "#4f7fb8", width: 3 },
      itemStyle: { color: "#4f7fb8" },
    }],
  }, true);
}

function setTopToBottom(metricId) {
  if (!metricId) return;
  const context = selectedDimensionContextForMetric(metricId);
  const bottomUp = context
    ? metricBottomUpFilteredSeries(metricId, context.coordinate)
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
  const inputs = [...document.querySelectorAll("#formula-input-list input:checked")].map(input => input.value);

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

  if (type === "cumulativeNetFlow") {
    applyStockFlowInputs();
    return;
  }

  definition.bottomUp = type === "laggedMetric"
    ? { type, inputs, lag: definition.bottomUp?.lag || 1 }
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

  const cohortStartSourceId = document.getElementById("cohort-start-source-input").value;
  if (cohortStartSourceId === definition.id) {
    setSourceStatus("A metric cannot source starting cohort values from itself.", "error");
    return;
  }

  const cohortStarts = {};
  document.querySelectorAll("[data-cohort-start-year]").forEach(input => {
    cohortStarts[input.dataset.cohortStartYear] = Number(input.value || 0);
  });

  const cohortAgeYoy = {};
  document.querySelectorAll("[data-cohort-yoy-age]").forEach(input => {
    cohortAgeYoy[input.dataset.cohortYoyAge] = Number(input.value || 0) / 100;
  });

  scenario().metrics[definition.id] = {
    ...metricScenario(definition.id),
    cohortStarts,
    cohortStartSourceId: cohortStartSourceId || null,
    cohortAgeYoy,
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
  clearCalculationCache();
  renderScenarioControls();
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

  const metricButton = event.target.closest("[data-metric-id]");
  if (metricButton) {
    selectMetric(metricButton.dataset.metricId);
    render();
    return;
  }
  if (event.target.closest("#set-top-to-bottom")) {
    setTopToBottom(state.selectedMetricId);
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

document.getElementById("bottom-up-form").addEventListener("submit", event => {
  event.preventDefault();
  applyBottomUpFormula();
});

document.getElementById("cohort-matrix-builder-form").addEventListener("submit", event => {
  event.preventDefault();
  applyCohortMatrixInputs();
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
  setSourceStatus("Catalog JSON has unapplied edits.");
});

document.body.addEventListener("change", event => {
  const scenarioSelect = event.target.closest("#scenario-select");
  if (scenarioSelect) {
    switchScenario(scenarioSelect.value);
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
  renderTopDownNapkinEditor(definition);
  renderLaggedOpeningEditor(definition);
});

window.addEventListener("resize", () => {
  chart.resize();
  topDownNapkinChart?.chart?.resize();
  dimensionMixChart?.chart?.resize();
  dimensionBreakdownChart.resize();
  cohortStartsChart.resize();
  cohortYoyChart.resize();
  renderCanvasLines();
});

render();
