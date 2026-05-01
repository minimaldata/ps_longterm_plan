(function () {
  const state = {
    selectedSourceId: "",
  };

  function label(value) {
    if (typeof window.displayLabel === "function") return window.displayLabel(value);
    return String(value);
  }

  function isAnonymized() {
    return typeof window.isAnonymizedView === "function" && window.isAnonymizedView();
  }

  function formatCell(value, format) {
    if (value === null || value === undefined || value === "") return "-";
    if (isAnonymized() && format !== "year" && format !== "cohort") return "Y";
    if (format === "year") return String(value);
    if (format === "cohort") return Number(value) === 0 ? "Legacy / Unknown" : String(value);
    if (format === "percent") return `${(Number(value) * 100).toFixed(1)}%`;
    if (format === "currency2") return `$${Number(value).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    if (format === "currency0") return `$${Math.round(Number(value)).toLocaleString("en-US")}`;
    if (format === "decimal") return Number(value).toFixed(2);
    if (format === "decimal4") return Number(value).toFixed(4);
    if (format === "integer") return Math.round(Number(value)).toLocaleString("en-US");
    return String(value);
  }

  function renderTable(source) {
    return `
      <div class="raw-data-table-wrap">
        <table class="raw-data-table">
          <thead>
            <tr>
              ${source.columns.map(column => `<th>${label(column.label)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${source.rows.map(row => `
              <tr>
                ${source.columns.map(column => `<td>${formatCell(row[column.key], column.format)}</td>`).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderEmpty(container) {
    container.innerHTML = `
      <div class="raw-data-empty">
        <h2>Raw Data Sources</h2>
        <p>No source data has been registered yet.</p>
      </div>
    `;
  }

  function render(containerId, sources) {
    const container = document.getElementById(containerId);
    if (!container) return;
    if (!Array.isArray(sources) || !sources.length) {
      renderEmpty(container);
      return;
    }

    if (!state.selectedSourceId || !sources.some(source => source.id === state.selectedSourceId)) {
      state.selectedSourceId = sources[0].id;
    }
    const selected = sources.find(source => source.id === state.selectedSourceId) || sources[0];

    container.innerHTML = `
      <div class="raw-data-layout">
        <aside class="raw-data-sidebar" aria-label="Raw data source list">
          ${sources.map(source => `
            <button
              type="button"
              class="${source.id === selected.id ? "active" : ""}"
              data-raw-source-id="${source.id}"
            >
              <strong>${label(source.title)}</strong>
              <span>${source.rows.length.toLocaleString("en-US")} rows</span>
            </button>
          `).join("")}
        </aside>
        <section class="raw-data-detail">
          <div class="raw-data-header">
            <div>
              <span>${label(selected.category || "Source Data")}</span>
              <h2>${label(selected.title)}</h2>
              <p>${label(selected.description || "")}</p>
            </div>
            <div class="raw-data-meta">
              <span>${selected.rows.length.toLocaleString("en-US")} rows</span>
              <span>${selected.columns.length.toLocaleString("en-US")} columns</span>
            </div>
          </div>
          ${renderTable(selected)}
        </section>
      </div>
    `;

    container.querySelectorAll("[data-raw-source-id]").forEach(button => {
      button.addEventListener("click", event => {
        state.selectedSourceId = event.currentTarget.dataset.rawSourceId;
        render(containerId, sources);
      });
    });
  }

  window.RawDataViewer = {
    render,
  };
})();
