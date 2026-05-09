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

function scenario() {
  return state.model.scenario;
}

function selectedDefinition() {
  return state.selectedMetricId ? metricDefinitions()[state.selectedMetricId] : null;
}

function selectedDimensionContextForMetric(metricId = state.selectedMetricId) {
  const context = state.selectedDimensionContext;
  if (!context || context.metricId !== metricId) return null;
  if (!metricUsesDimension(metricId, context.dimensionId)) return null;
  const member = dimensionMembers(context.dimensionId).find(item => item.id === context.memberId);
  return member ? { ...context, member } : null;
}

function selectedNodeKey() {
  const context = selectedDimensionContextForMetric();
  return context
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
    ? { metricId: parsed.metricId, dimensionId: parsed.dimensionId, memberId: parsed.memberId }
    : null;
}

function metricScenario(metricId) {
  return scenario().metrics[metricId] || {};
}

function metricDimensionIds(metricId) {
  const dimensions = metricDefinitions()[metricId]?.dimensions;
  return Array.isArray(dimensions) ? dimensions : [];
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
    const topDown = metricScenario(metricId).topDown;
    if (Array.isArray(topDown)) return metricDimensionIds(metricId).length ? defaultSeries() : topDown;
    if (topDown && typeof topDown === "object" && Array.isArray(topDown[memberId])) {
      return topDown[memberId];
    }
    return defaultSeries();
  });
}

function metricMemberControlPoints(metricId, memberId) {
  const stored = metricScenario(metricId).memberManualControlPoints?.[memberId];
  if (Array.isArray(stored) && stored.length) return normalizeNapkinPoints(stored);
  return YEARS.map((year, index) => [year, Number(metricTopDownMemberSeries(metricId, memberId)[index] || 0)]);
}

function setMetricTopDownMemberSeries(metricId, memberId, series) {
  const metric = ensureMetricScenario(metricId);
  const dimensionId = primaryDimensionId(metricId);
  const members = dimensionMembers(dimensionId);
  const existingTopDown = metric.topDown;

  if (!existingTopDown || Array.isArray(existingTopDown)) {
    const totalSeries = Array.isArray(existingTopDown) ? existingTopDown : metricTopDownSeries(metricId);
    const evenShare = members.length ? 1 / members.length : 1;
    metric.topDown = Object.fromEntries(members.map(member => [
      member.id,
      YEARS.map((_year, index) => Number(totalSeries[index] || 0) * evenShare),
    ]));
  }

  if (!metric.topDown || Array.isArray(metric.topDown)) metric.topDown = {};
  metric.topDown[memberId] = [...series];
  delete metric.topDownTotal;
  delete metric.manualControlPoints;
  delete metric.dimensionShareControlPoints;
  clearCalculationCache();
}

function metricTopDownSeries(metricId) {
  return cachedCalculation(`topDown:${metricId}`, () => {
    const override = metricScenario(metricId).topDownTotal;
    if (Array.isArray(override)) return override;
    const topDown = metricScenario(metricId).topDown;
    if (Array.isArray(topDown)) return topDown;
    if (topDown && typeof topDown === "object") return aggregateMemberSeries(metricId, topDown);
    return defaultSeries();
  });
}

function setMetricTopDownSeries(metricId, series) {
  if (allocateDimensionTopDownSeries(metricId, series)) return;
  const metric = ensureMetricScenario(metricId);
  if (metric.topDown && typeof metric.topDown === "object" && !Array.isArray(metric.topDown)) {
    metric.topDownTotal = [...series];
    clearCalculationCache();
    return;
  }
  metric.topDown = [...series];
  clearCalculationCache();
}

function allocateDimensionTopDownSeries(metricId, series) {
  const definition = metricDefinitions()[metricId];
  if (metricIsRate(definition)) return false;

  const dimensionId = primaryDimensionId(metricId);
  const members = dimensionMembers(dimensionId);
  if (!dimensionId || !members.length) return false;

  const existingSeriesByMember = Object.fromEntries(members.map(member => [
    member.id,
    metricTopDownMemberSeries(metricId, member.id),
  ]));
  const memberSharesByYear = YEARS.map((_year, index) => {
    const directShare = sharesForIndex(members, existingSeriesByMember, index);
    if (directShare) return directShare;

    const nearestIndex = nearestIndexWithMemberMix(members, existingSeriesByMember, index);
    if (nearestIndex !== null) return sharesForIndex(members, existingSeriesByMember, nearestIndex);

    const evenShare = 1 / members.length;
    return Object.fromEntries(members.map(member => [member.id, evenShare]));
  });

  const allocatedTopDown = Object.fromEntries(members.map(member => [
    member.id,
    YEARS.map((_year, index) => Number(series[index] || 0) * Number(memberSharesByYear[index]?.[member.id] || 0)),
  ]));

  const metric = ensureMetricScenario(metricId);
  metric.topDown = allocatedTopDown;
  delete metric.topDownTotal;
  delete metric.memberManualControlPoints;
  clearCalculationCache();
  return true;
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
  metric.topDown = Object.fromEntries(members.map(member => [
    member.id,
    YEARS.map((_year, index) => Number(totals[index] || 0) * Number(sharesByYear[index]?.[member.id] || 0)),
  ]));
  delete metric.topDownTotal;
  delete metric.memberManualControlPoints;
  clearCalculationCache();
}

function ensureMetricScenario(metricId) {
  if (!scenario().metrics[metricId]) {
    scenario().metrics[metricId] = { topDown: defaultSeries() };
  }
  return scenario().metrics[metricId];
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
  const stored = metricScenario(metricId).manualControlPoints;
  if (Array.isArray(stored) && stored.length) return normalizeNapkinPoints(stored);
  return YEARS.map((year, index) => [year, Number(metricTopDownSeries(metricId)[index] || 0)]);
}

function activeTopDownControlPoints(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  return context
    ? metricMemberControlPoints(definition.id, context.memberId)
    : manualControlPoints(definition.id);
}

function activeTopDownSeries(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  return context
    ? metricTopDownMemberSeries(definition.id, context.memberId)
    : metricTopDownSeries(definition.id);
}

function activeBottomUpSeries(definition) {
  const context = selectedDimensionContextForMetric(definition?.id);
  return context
    ? metricBottomUpMemberSeries(definition.id, context.dimensionId, context.memberId)
    : metricBottomUpSeries(definition.id);
}

function setActiveTopDownSeries(definition, points) {
  const series = seriesFromControlPoints(points);
  const metric = ensureMetricScenario(definition.id);
  const context = selectedDimensionContextForMetric(definition.id);

  if (context) {
    if (!metric.memberManualControlPoints) metric.memberManualControlPoints = {};
    metric.memberManualControlPoints[context.memberId] = points;
    setMetricTopDownMemberSeries(definition.id, context.memberId, series);
    return;
  }

  metric.manualControlPoints = points;
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
  return context && metricUsesDimension(metricId, context.dimensionId)
    ? metricTopDownMemberSeries(metricId, context.memberId)
    : metricTopDownSeries(metricId);
}

function bottomUpSeriesForContext(metricId, context = null) {
  return context && metricUsesDimension(metricId, context.dimensionId)
    ? metricBottomUpMemberSeries(metricId, context.dimensionId, context.memberId)
    : metricBottomUpSeries(metricId);
}

function laggedComparisonLines(definition, context = null) {
  const [sourceMetricId] = metricInputIds(definition);
  if (!sourceMetricId || !metricDefinitions()[sourceMetricId]) return null;
  const sourceDefinition = metricDefinitions()[sourceMetricId];
  const sourceContext = context && metricUsesDimension(sourceMetricId, context.dimensionId)
    ? context
    : null;
  const memberId = sourceContext?.memberId || null;
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
    ? metricTopDownMemberSeries(parsed.metricId, parsed.memberId)
    : metricTopDownSeries(parsed.metricId);
}

function metricValueAtIndex(metricId, index, stack = new Set(), memberContext = null) {
  const definition = metricDefinitions()[metricId];
  const memberId = metricUsesDimension(metricId, memberContext?.dimensionId) ? memberContext.memberId : null;
  if (!definition?.bottomUp) {
    const series = memberId ? metricTopDownMemberSeries(metricId, memberId) : metricTopDownSeries(metricId);
    return Number(series[index] || 0);
  }

  const stackKey = `${metricId}:${memberId || "total"}:${index}`;
  if (stack.has(stackKey)) {
    const series = memberId ? metricTopDownMemberSeries(metricId, memberId) : metricTopDownSeries(metricId);
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

    const dimensionId = primaryDimensionId(metricId);
    const members = dimensionMembers(dimensionId);
    if (members.length) {
      const seriesByMember = Object.fromEntries(members.map(member => [
        member.id,
        YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { dimensionId, memberId: member.id })),
      ]));
      return aggregateMemberSeries(metricId, seriesByMember);
    }

    return YEARS.map((_year, index) => metricValueAtIndex(metricId, index));
  });
}

function metricBottomUpMemberSeries(metricId, dimensionId, memberId) {
  return cachedCalculation(`bottomUpMember:${metricId}:${dimensionId}:${memberId}`, () => {
    const definition = metricDefinitions()[metricId];
    if (!definition?.bottomUp || !metricUsesDimension(metricId, dimensionId)) return null;
    return YEARS.map((_year, index) => metricValueAtIndex(metricId, index, new Set(), { dimensionId, memberId }));
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
    const level = current.filter(nodeKey => {
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

function sourceSnapshot() {
  return JSON.stringify({
    name: state.model.name,
    years: YEARS,
    selectedMetricId: state.selectedMetricId,
    selectedDimensionContext: state.selectedDimensionContext,
    dimensions: dimensionDefinitions(),
    metricDefinitions: metricDefinitions(),
    scenario: scenarioSnapshot(),
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
  return {
    ...rawMetric,
    topDown: rawMetric.topDown || defaultSeries(),
  };
}

function scenarioSnapshot() {
  const metrics = {};
  Object.entries(scenario().metrics || {}).forEach(([id, metric]) => {
    const definition = metricDefinitions()[id];
    metrics[id] = metricIsLagged(definition)
      ? laggedMetricScenarioSnapshot(metric)
      : metric;
  });

  return {
    ...scenario(),
    metrics,
  };
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
  if (!rawModel.scenario || typeof rawModel.scenario !== "object") {
    throw new Error("Catalog source must include a scenario object.");
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

  const metrics = {};
  Object.keys(definitions).forEach(id => {
    metrics[id] = normalizeScenarioMetric(definitions[id], rawModel.scenario.metrics?.[id] || {});
  });

  return {
    name: rawModel.name || "Custom Model",
    dimensions: rawModel.dimensions || {},
    metricDefinitions: definitions,
    scenario: {
      id: rawModel.scenario.id || "scenario-1",
      name: rawModel.scenario.name || "Scenario 1",
      metrics,
    },
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
  const dimensionId = isMember ? null : primaryDimensionId(parsed.metricId);
  const memberCount = isMember ? 0 : dimensionMembers(dimensionId).length;
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
      memberCount ? `${memberCount} members` : "",
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
    .map(dimensionId => dimensionDefinitions()[dimensionId]?.label || dimensionId)
    .join(", ") || "None";
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
  const dimensionLabel = dimensionDefinitions()[context.dimensionId]?.label || context.dimensionId;
  document.getElementById("detail-title").textContent = `${definition.label}: ${context.member.label}`;
  document.getElementById("detail-subtitle").textContent = `Member of ${dimensionLabel}`;
  document.getElementById("top-down-definition").textContent = definition.topDown
    ? `${definition.topDown.type}${definition.topDown.editable ? " (editable)" : ""}`
    : "None";
  document.getElementById("bottom-up-definition").textContent = definition.bottomUp
    ? `${formulaText(definition.id)} in ${context.member.label} context`
    : "No member-specific bottom-up definition";
  document.getElementById("time-definition").textContent = `${definition.time.flowType}, ${definition.time.aggregateMethod}`;
  document.getElementById("dimension-definition").textContent = `${dimensionLabel}: ${context.member.label}`;
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

  const dimensionId = primaryDimensionId(definition.id);
  const members = dimensionMembers(dimensionId);
  if (!dimensionId || !members.length) {
    section.classList.add("is-hidden");
    list.innerHTML = "";
    return;
  }

  const dimensionLabel = dimensionDefinitions()[dimensionId]?.label || dimensionId;
  const selectedMember = selectedDimensionContextForMetric(definition.id);
  section.classList.remove("is-hidden");
  document.getElementById("dimension-member-title").textContent = `${dimensionLabel} Members`;
  list.innerHTML = [
    `<button class="dimension-member-chip ${selectedMember ? "" : "active"}" data-member-total="${definition.id}" type="button">Total</button>`,
    ...members.map(member => `
      <button class="dimension-member-chip ${selectedMember?.memberId === member.id ? "active" : ""}" data-member-context="${definition.id}" data-dimension-id="${dimensionId}" data-member-id="${member.id}" type="button">
        ${member.label}
      </button>
    `),
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
  const members = dimensionMembers(dimensionId);
  if (!dimensionId || !members.length) {
    section.classList.add("is-hidden");
    dimensionBreakdownChart.clear();
    return;
  }

  section.classList.remove("is-hidden");
  document.getElementById("dimension-breakdown-title").textContent = `${dimensionDefinitions()[dimensionId]?.label || dimensionId} Top-Down Breakdown`;
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
    series: members.map(member => ({
      name: member.label,
      type: "line",
      data: metricTopDownMemberSeries(definition.id, member.id),
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
  const shouldShow = Boolean(definition && !metricIsLagged(definition) && !metricIsRate(definition) && dimensionId && members.length >= 2);
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
    ? metricMemberReconciliation(metricId, context.dimensionId, context.memberId)
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
  if (context) {
    inputList.innerHTML = `
      <label>
        ${context.member.label}
        <input type="number" step="1" data-lagged-opening-member-id="${context.memberId}" value="${metricOpeningValue(definition.id, context.memberId)}">
      </label>
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
  const label = parsed.memberId
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

function renderTopDownNapkinEditor(definition) {
  const section = document.getElementById("top-down-napkin-section");
  const container = document.getElementById("top-down-napkin-chart");
  const shouldShow = Boolean(definition && !metricIsLagged(definition));
  section.classList.toggle("is-hidden", !shouldShow);

  if (!shouldShow) {
    disposeTopDownNapkinChart();
    container.innerHTML = "";
    return;
  }

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
  const lineName = context ? `${definition.label}: ${context.member.label}` : definition.label;
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
    ? metricBottomUpMemberSeries(metricId, context.dimensionId, context.memberId)
    : metricBottomUpSeries(metricId);
  if (!bottomUp) return;
  const metric = ensureMetricScenario(metricId);
  const points = YEARS.map((year, index) => [year, Number(bottomUp[index] || 0)]);
  if (context) {
    if (!metric.memberManualControlPoints) metric.memberManualControlPoints = {};
    metric.memberManualControlPoints[context.memberId] = points;
    setMetricTopDownMemberSeries(metricId, context.memberId, bottomUp);
  } else {
    setMetricTopDownSeries(metricId, bottomUp);
    metric.manualControlPoints = points;
  }
  render();
}

function newBlankModel() {
  state.model = createBlankModel();
  state.selectedMetricId = null;
  state.selectedDimensionContext = null;
  catalogSourceIsDirty = false;
  render();
}

function loadStarterModel() {
  state.model = createStarterModel();
  state.selectedMetricId = "profit";
  state.selectedDimensionContext = null;
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
  scenario().metrics[id] = { topDown: defaultSeries() };
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
    ensureMetricScenario(definition.id).manualControlPoints = manualControlPoints(definition.id);
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

function render() {
  clearCalculationCache();
  renderMetricList();
  renderCanvas();
  renderDetail();
  renderCatalogSource();
  const definition = selectedDefinition();
  document.getElementById("set-top-to-bottom").disabled = !definition || !activeBottomUpSeries(definition);
  requestAnimationFrame(renderCanvasLines);
}

document.body.addEventListener("click", event => {
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
      dimensionId: memberContextButton.dataset.dimensionId,
      memberId: memberContextButton.dataset.memberId,
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
  if (event.target.closest("#add-metric")) {
    document.getElementById("metric-name-input").focus();
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
