import { YEARS as DEFAULT_YEARS } from "./metric-catalog.js";

let YEARS = [...DEFAULT_YEARS];

function defaultSeries() {
  return YEARS.map(() => 0);
}

export function setEngineYears(years) {
  YEARS = Array.isArray(years) && years.length ? [...years] : [...DEFAULT_YEARS];
}

export function coordinateKey(coordinate = {}) {
  return Object.entries(coordinate)
    .filter(([_dimensionId, memberId]) => memberId !== undefined && memberId !== null && memberId !== "")
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([dimensionId, memberId]) => `${encodeURIComponent(dimensionId)}=${encodeURIComponent(memberId)}`)
    .join("|");
}

export function parseCoordinateKey(key) {
  if (!key) return {};
  return String(key).split("|").reduce((coordinate, part) => {
    const [dimensionId, memberId] = part.split("=");
    if (!dimensionId || memberId === undefined) return coordinate;
    coordinate[decodeURIComponent(dimensionId)] = decodeURIComponent(memberId);
    return coordinate;
  }, {});
}

export function memberCoordinateKey(dimensionId, memberId) {
  return coordinateKey({ [dimensionId]: memberId });
}

export function coordinateMatches(key, filters = {}) {
  const coordinate = parseCoordinateKey(key);
  return Object.entries(filters).every(([dimensionId, memberId]) => coordinate[dimensionId] === memberId);
}

export function sumSeries(left = defaultSeries(), right = defaultSeries()) {
  return YEARS.map((_year, index) => Number(left[index] || 0) + Number(right[index] || 0));
}

export function splitSeries(series = defaultSeries(), share = 1) {
  return YEARS.map((_year, index) => Number(series[index] || 0) * share);
}

export function cohortValueForYear(row, year) {
  if (!row || typeof row !== "object") return 0;
  return Number(row[String(year)] || 0);
}

export function cohortStartValueForYear(starts, year) {
  return Number(starts?.[String(year)] || 0);
}

export function cohortYoyValueForAge(ageYoy, age) {
  if (age <= 0) return 0;
  if (Array.isArray(ageYoy)) return Number(ageYoy[age - 1] || 0);
  return Number(ageYoy?.[String(age)] || 0);
}

export function openingCohortAge(key) {
  const match = String(key).match(/\d+/);
  return match ? Number(match[0]) : 0;
}

export function cohortCurveValueForAge(curve = {}, age, fallbackValue = 0) {
  const stringAge = String(age);
  if (Object.prototype.hasOwnProperty.call(curve, stringAge)) return Number(curve[stringAge] || 0);
  const numericAges = Object.keys(curve)
    .map(value => Number(value))
    .filter(value => Number.isFinite(value))
    .sort((left, right) => left - right);
  const closestPriorAge = numericAges.filter(value => value <= age).at(-1);
  if (closestPriorAge !== undefined) return Number(curve[String(closestPriorAge)] || 0);
  return fallbackValue;
}

export function matrixTotalSeries(matrix = {}) {
  return YEARS.map(year => Object.values(matrix).reduce((sum, row) => sum + cohortValueForYear(row, year), 0));
}

export function addCohortMatrices(left = {}, right = {}) {
  const matrix = JSON.parse(JSON.stringify(left || {}));
  Object.entries(right || {}).forEach(([rowKey, row]) => {
    if (!matrix[rowKey]) matrix[rowKey] = {};
    YEARS.forEach(year => {
      const yearKey = String(year);
      const value = Number(row?.[yearKey] || 0);
      if (!value) return;
      matrix[rowKey][yearKey] = Number(matrix[rowKey]?.[yearKey] || 0) + value;
    });
  });
  return matrix;
}

export function rollupCohortMatrices(matrixByCoordinate = {}, filters = {}) {
  return Object.entries(matrixByCoordinate)
    .filter(([key]) => coordinateMatches(key, filters))
    .reduce((matrix, [_key, coordinateMatrix]) => addCohortMatrices(matrix, coordinateMatrix), {});
}

export function flatOpeningCohorts(rawOpeningCohorts = {}) {
  return Object.fromEntries(
    Object.entries(rawOpeningCohorts || {})
      .filter(([_key, value]) => typeof value !== "object" || value === null)
      .map(([key, value]) => [key, Number(value || 0)])
  );
}

export function coordinateOpeningCohorts(rawOpeningCohorts = {}, coordinateKeyValue) {
  const direct = rawOpeningCohorts?.[coordinateKeyValue];
  return direct && typeof direct === "object" && !Array.isArray(direct) ? direct : null;
}

export function splitOpeningCohortsByCoordinate(rawOpeningCohorts = {}, coordinateKeys, startSeriesByCoordinate) {
  const hasCoordinateSpecificValues = Object.values(rawOpeningCohorts || {})
    .some(value => value && typeof value === "object" && !Array.isArray(value));
  if (hasCoordinateSpecificValues) {
    return Object.fromEntries(coordinateKeys.map(key => [key, coordinateOpeningCohorts(rawOpeningCohorts, key) || {}]));
  }

  const flat = flatOpeningCohorts(rawOpeningCohorts);
  if (!Object.keys(flat).length) return Object.fromEntries(coordinateKeys.map(key => [key, {}]));
  if (coordinateKeys.length <= 1) return { [coordinateKeys[0] || ""]: flat };

  const firstYearValues = coordinateKeys.map(key => Number(startSeriesByCoordinate[key]?.[0] || 0));
  const firstYearTotal = firstYearValues.reduce((sum, value) => sum + value, 0);
  const evenShare = coordinateKeys.length ? 1 / coordinateKeys.length : 1;
  return Object.fromEntries(coordinateKeys.map((key, index) => {
    const share = firstYearTotal ? firstYearValues[index] / firstYearTotal : evenShare;
    return [key, Object.fromEntries(Object.entries(flat).map(([ageKey, value]) => [ageKey, Number(value || 0) * share]))];
  }));
}

export function buildGenericCohortMatrix({ startSeries, openingCohorts, curve, curveType, startTiming }) {
  const matrix = {};

  Object.entries(openingCohorts || {}).forEach(([ageKey, value]) => {
    const startAge = openingCohortAge(ageKey);
    const rowKey = `opening_${ageKey}`;
    let runningValue = Number(value || 0);
    if (!runningValue) return;
    matrix[rowKey] = {};
    YEARS.forEach((_year, yearIndex) => {
      const age = startAge + yearIndex;
      if (yearIndex > 0) {
        const priorAge = age - 1;
        const curveValue = cohortCurveValueForAge(curve, priorAge, curveType === "survivalRate" ? 1 : 0);
        runningValue = curveType === "survivalRate"
          ? runningValue * curveValue
          : runningValue * (1 + curveValue);
      }
      matrix[rowKey][String(YEARS[yearIndex])] = runningValue;
    });
  });

  YEARS.forEach((cohortYear, cohortIndex) => {
    const startValue = Number(startSeries[cohortIndex] || 0);
    if (!startValue) return;
    const row = {};
    let runningValue = startValue;
    YEARS.forEach((year, yearIndex) => {
      if (yearIndex < cohortIndex) return;
      const age = yearIndex - cohortIndex;
      if (startTiming === "nextPeriod" && age === 0) {
        row[String(year)] = 0;
        return;
      }
      if (curveType === "survivalRate") {
        if (startTiming === "nextPeriod") {
          const retentionAge = age - 1;
          if (retentionAge === 0) {
            runningValue = startValue * cohortCurveValueForAge(curve, 0, 1);
          } else {
            runningValue *= cohortCurveValueForAge(curve, retentionAge, 1);
          }
        } else if (age === 0) {
          runningValue = startValue;
        } else {
          runningValue *= cohortCurveValueForAge(curve, age - 1, 1);
        }
      } else if (curveType === "yoyGrowth") {
        if (age === 0) {
          runningValue = startValue;
        } else {
          runningValue *= 1 + cohortCurveValueForAge(curve, age, 0);
        }
      }
      row[String(year)] = runningValue;
    });
    matrix[String(cohortYear)] = row;
  });

  return matrix;
}

export function sortedCohortMatrixRowKeys(matrix = {}) {
  return Object.keys(matrix).sort((left, right) => {
    const leftOpening = String(left).startsWith("opening_");
    const rightOpening = String(right).startsWith("opening_");
    if (leftOpening !== rightOpening) return leftOpening ? -1 : 1;
    return Number(left.replace(/\D/g, "")) - Number(right.replace(/\D/g, ""));
  });
}
