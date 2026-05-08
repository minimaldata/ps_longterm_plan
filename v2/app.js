import { YEARS, metricDefinitions, scenario } from "./metric-catalog.js";

const state = {
  selectedMetricId: "profit",
};

const currencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 0,
});

const chart = echarts.init(document.getElementById("metric-detail-chart"));

function metricTopDownSeries(metricId) {
  return scenario.metrics[metricId]?.topDown || YEARS.map(() => 0);
}

function metricBottomUpSeries(metricId) {
  const definition = metricDefinitions[metricId];
  if (!definition?.bottomUp) return null;

  if (definition.bottomUp.type === "sum") {
    return YEARS.map((_year, index) => {
      return definition.bottomUp.inputs.reduce((sum, inputId) => {
        const inputSeries = metricEffectiveSeries(inputId);
        return sum + Number(inputSeries[index] || 0);
      }, 0);
    });
  }

  if (definition.bottomUp.type === "difference") {
    const [leftId, rightId] = definition.bottomUp.inputs;
    const left = metricEffectiveSeries(leftId);
    const right = metricEffectiveSeries(rightId);
    return YEARS.map((_year, index) => Number(left[index] || 0) - Number(right[index] || 0));
  }

  return null;
}

function metricEffectiveSeries(metricId) {
  return metricBottomUpSeries(metricId) || metricTopDownSeries(metricId);
}

function metricReconciliation(metricId) {
  const definition = metricDefinitions[metricId];
  if (!definition?.reconciliation?.enabled) return { enabled: false, matched: true, maxDelta: 0 };
  const topDown = metricTopDownSeries(metricId);
  const bottomUp = metricBottomUpSeries(metricId);
  if (!bottomUp) return { enabled: false, matched: true, maxDelta: 0 };
  const maxDelta = Math.max(...YEARS.map((_year, index) => Math.abs(Number(topDown[index] || 0) - Number(bottomUp[index] || 0))));
  return {
    enabled: true,
    matched: maxDelta <= Number(definition.reconciliation.tolerance || 0),
    maxDelta,
  };
}

function childIds(metricId) {
  return metricDefinitions[metricId]?.bottomUp?.inputs || [];
}

function parentIds(metricId) {
  return Object.values(metricDefinitions)
    .filter(definition => childIds(definition.id).includes(metricId))
    .map(definition => definition.id);
}

function formatCurrency(value) {
  return currencyFormatter.format(Number(value || 0));
}

function compactCurrency(value) {
  const numericValue = Number(value || 0);
  const abs = Math.abs(numericValue);
  const sign = numericValue < 0 ? "-" : "";
  if (abs >= 1000000) return `${sign}$${(abs / 1000000).toFixed(abs >= 100000000 ? 0 : 1)}M`;
  if (abs >= 1000) return `${sign}$${(abs / 1000).toFixed(0)}k`;
  return formatCurrency(numericValue);
}

function formulaText(metricId) {
  const definition = metricDefinitions[metricId];
  if (!definition?.bottomUp) return "No bottom-up definition";
  const labels = definition.bottomUp.inputs.map(inputId => metricDefinitions[inputId]?.label || inputId);
  if (definition.bottomUp.type === "sum") return labels.join(" + ");
  if (definition.bottomUp.type === "difference") return labels.join(" - ");
  return definition.bottomUp.type;
}

function renderMetricList() {
  const list = document.getElementById("metric-list");
  list.innerHTML = Object.values(metricDefinitions).map(definition => {
    const reconciliation = metricReconciliation(definition.id);
    return `
      <button class="metric-list-item ${definition.id === state.selectedMetricId ? "active" : ""}" data-metric-id="${definition.id}" type="button">
        <span>${definition.label}</span>
        ${reconciliation.enabled ? `<small class="${reconciliation.matched ? "matched" : "open"}">${reconciliation.matched ? "Matched" : "Mismatch"}</small>` : "<small>Leaf</small>"}
      </button>
    `;
  }).join("");
}

function renderCanvas() {
  const canvas = document.getElementById("metric-canvas");
  const levels = [
    ["profit"],
    ["revenue", "cost"],
    ["productARevenue", "productBRevenue", "servicesRevenue", "laborCost", "nonLaborCost"],
  ];

  const nodeMarkup = levels.map((level, levelIndex) => `
    <div class="canvas-level level-${levelIndex}">
      ${level.map(metricId => renderNode(metricId)).join("")}
    </div>
  `).join("");

  canvas.innerHTML = `
    <svg class="canvas-lines" viewBox="0 0 1000 520" preserveAspectRatio="none" aria-hidden="true">
      <path d="M 500 86 C 500 130, 290 130, 290 176" />
      <path d="M 500 86 C 500 130, 710 130, 710 176" />
      <path d="M 290 250 C 290 300, 140 300, 140 356" />
      <path d="M 290 250 C 290 300, 290 300, 290 356" />
      <path d="M 290 250 C 290 300, 440 300, 440 356" />
      <path d="M 710 250 C 710 300, 625 300, 625 356" />
      <path d="M 710 250 C 710 300, 795 300, 795 356" />
    </svg>
    ${nodeMarkup}
  `;
}

function renderNode(metricId) {
  const definition = metricDefinitions[metricId];
  const series = metricEffectiveSeries(metricId);
  const reconciliation = metricReconciliation(metricId);
  const isSelected = metricId === state.selectedMetricId;
  const hasChildren = childIds(metricId).length > 0;
  return `
    <button class="metric-node ${isSelected ? "active" : ""} ${reconciliation.enabled && !reconciliation.matched ? "mismatch" : ""}" data-metric-id="${metricId}" type="button">
      <span>${definition.label}</span>
      <strong>${compactCurrency(series[series.length - 1])}</strong>
      <small>${hasChildren ? `${childIds(metricId).length} inputs` : "Manual series"}</small>
    </button>
  `;
}

function renderDetail() {
  const metricId = state.selectedMetricId;
  const definition = metricDefinitions[metricId];
  const topDown = metricTopDownSeries(metricId);
  const bottomUp = metricBottomUpSeries(metricId);
  const reconciliation = metricReconciliation(metricId);
  const parents = parentIds(metricId).map(id => metricDefinitions[id].label);
  const children = childIds(metricId).map(id => metricDefinitions[id].label);

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

  const badge = document.getElementById("reconciliation-badge");
  badge.textContent = reconciliation.enabled
    ? reconciliation.matched ? "Matched" : `Mismatch ${formatCurrency(reconciliation.maxDelta)}`
    : "Leaf";
  badge.className = `reconciliation-badge ${reconciliation.enabled ? reconciliation.matched ? "matched" : "open" : ""}`;

  const series = [{
    name: "Top-Down",
    type: "line",
    data: topDown,
    symbolSize: 7,
    lineStyle: { color: "#2f6f73", width: 3 },
    itemStyle: { color: "#2f6f73" },
  }];
  if (bottomUp) {
    series.push({
      name: "Bottom-Up",
      type: "line",
      data: bottomUp,
      symbolSize: 6,
      lineStyle: { color: "#4f7fb8", width: 3 },
      itemStyle: { color: "#4f7fb8" },
    });
  }

  chart.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      valueFormatter: value => formatCurrency(value),
    },
    legend: { top: 0 },
    grid: { left: 12, right: 18, top: 42, bottom: 34, containLabel: true },
    xAxis: { type: "category", data: YEARS.map(String) },
    yAxis: { type: "value", axisLabel: { formatter: compactCurrency } },
    series,
  }, true);
}

function setTopToBottom(metricId) {
  const bottomUp = metricBottomUpSeries(metricId);
  if (!bottomUp) return;
  scenario.metrics[metricId].topDown = [...bottomUp];
  render();
}

function render() {
  renderMetricList();
  renderCanvas();
  renderDetail();
}

document.body.addEventListener("click", event => {
  const metricButton = event.target.closest("[data-metric-id]");
  if (metricButton) {
    state.selectedMetricId = metricButton.dataset.metricId;
    render();
    return;
  }
  if (event.target.closest("#set-top-to-bottom")) {
    setTopToBottom(state.selectedMetricId);
  }
});

window.addEventListener("resize", () => chart.resize());

render();
