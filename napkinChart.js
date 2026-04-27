// ----------------- START OF NapkinChart CLASS -----------------
class NapkinChart {
  /**
   * @param {string} domId - The ID of the DOM element where we render ECharts
   * @param {Object[]} initialLines - Array of lines to plot
   *    Each line: { name: string, color: string, data: [ [x1,y1], [x2,y2], ... ] }
   * @param {boolean} editMode - Whether the chart is currently in editable state
   * @param {Object} baseOption - The base option for the chart (fixed axis ranges, etc.)
   * @param {string} monotonic - "none", "ascending", or "descending"
   */
  constructor(domId, initialLines, editMode = true, baseOption = null, monotonic = 'none', enableZoomBar = true) {
    this.container = document.getElementById(domId);
    this.chart = echarts.init(this.container);

    // We add a global max X domain for dataZoom (e.g. 360)
    this.globalMaxX = 50;
    this.windowStartX = 0; // start visible window
    this.windowEndX = 50;  // end visible window

    // Store the lines
    this.lines = JSON.parse(JSON.stringify(initialLines || []));
    this.editMode = editMode;

    // For Undo/Redo
    this.history = [];
    this.historyIndex = -1;

    // Monotonic constraint setting
    this.monotonic = monotonic;

    // Optional callback for data changes (e.g., to refresh a table)
    this.onDataChanged = null;

    // Base chart option
    this.baseOption = baseOption || {
      animation: false,
      tooltip: { trigger: 'axis' },
      xAxis: {
        type: 'value',
        min: 0,
        max: 100,
      },
      yAxis: {
        type: 'value',
        min: 0,
        max: 100,
        name: 'Y',
        axisLabel: {
          formatter: function (value) {
            if (value >= 1e9) {
              return (value / 1e9).toFixed(1) + 'B';
            } else if (value >= 1e6) {
              return (value / 1e6).toFixed(1) + 'M';
            } else if (value >= 1e3) {
              return (value / 1e3).toFixed(1) + 'k';
            } else {
              return value;
            }
          }
        }
      },
      series: [],
    };

    // Toggle slider dataZoom
    this.enableZoomBar = enableZoomBar;

    // Initial render & event setup
    this._refreshChart();
    this._initEvents();

    // Handle window resize
    window.addEventListener('resize', () => this.chart.resize());
  }

  // A helper for any external code that wants to update the chart
  setOption(newOption) {
    this.chart.setOption(newOption, true);
  }

  resize() {
    // this.chart is the real ECharts instance
    this.chart.resize();
  }

  /** Toggle whether the chart is editable */
  setEditMode(mode) {
    this.editMode = mode;
  }

  /** Programmatically add a new line */
  addLine(lineObj) {
    this._pushHistory();
    this.lines.push(JSON.parse(JSON.stringify(lineObj)));
    this._refreshChart();
  }

  /** Remove a line by name (or index). Adjust as needed. */
  removeLineByName(lineName) {
    this._pushHistory();
    this.lines = this.lines.filter((ln) => ln.name !== lineName);
    this._refreshChart();
  }

  /** Return all lines data as JSON. */
  exportDataAsJSON() {
    return JSON.stringify(this.lines, null, 2);
  }

  /** Undo / Redo logic */
  undo() {
    if (this.historyIndex > 0) {
      this.historyIndex--;
      const state = JSON.parse(this.history[this.historyIndex]);
      this.lines = state;
      this._refreshChart();
    }
  }

  redo() {
    if (this.historyIndex < this.history.length - 1) {
      this.historyIndex++;
      const state = JSON.parse(this.history[this.historyIndex]);
      this.lines = state;
      this._refreshChart();
    }
  }

  // Called whenever we are about to change data
  _pushHistory() {
    // Trim any redo states if we are in the middle of the stack
    this.history = this.history.slice(0, this.historyIndex + 1);
    // Save the current state
    this.history.push(JSON.stringify(this.lines));
    this.historyIndex++;
  }

  _interpolateLine(data, minX, maxX) {
    // Sort points by x
    const sorted = data.slice().sort((a,b) => a[0] - b[0]);
    const result = [];
    for (let x = minX; x <= maxX; x++) {
      const y = this._piecewiseLinear(sorted, x);
      result.push([x, y]);
    }
    return result;
  }

  _piecewiseLinear(sortedData, x) {
    // If x <= the first data point's x, clamp:
    if (x <= sortedData[0][0]) {
      return sortedData[0][1];
    }
    // If x >= the last data point's x, clamp:
    if (x >= sortedData[sortedData.length - 1][0]) {
      return sortedData[sortedData.length - 1][1];
    }
    // Otherwise find segment:
    for (let i = 0; i < sortedData.length - 1; i++) {
      const [x1,y1] = sortedData[i];
      const [x2,y2] = sortedData[i+1];
      if (x1 <= x && x <= x2) {
        const ratio = (x - x1) / (x2 - x1);
        return y1 + ratio*(y2 - y1);
      }
    }
    // fallback if for some reason we didn't find a segment
    return sortedData[sortedData.length - 1][1];
  }

  _extractXFromDatum(datum) {
    if (Array.isArray(datum)) {
      return Number(datum[0]);
    }
    if (datum && Array.isArray(datum.value)) {
      return Number(datum.value[0]);
    }
    return NaN;
  }

  _normalizeEditRanges(ranges) {
    if (!Array.isArray(ranges)) {
      return null;
    }
    const normalized = [];
    ranges.forEach((range) => {
      if (!Array.isArray(range) || range.length < 2) return;
      const rawStart = Number(range[0]);
      const rawEnd = Number(range[1]);
      if (Number.isNaN(rawStart) || Number.isNaN(rawEnd)) return;
      const start = Math.min(rawStart, rawEnd);
      const end = Math.max(rawStart, rawEnd);
      normalized.push([start, end]);
    });
    return normalized;
  }

  _getEditRangesForOperation(line, operation = 'move') {
    if (!line || typeof line !== 'object') return null;

    let ranges = null;
    const domain = line.editDomain;
    if (domain && typeof domain === 'object') {
      const opKey = `${operation}X`;
      if (Array.isArray(domain[opKey])) {
        ranges = domain[opKey];
      } else if (Array.isArray(domain.x)) {
        ranges = domain.x;
      }
    }

    if (ranges === null) {
      if (Array.isArray(line.editableXRanges)) {
        ranges = line.editableXRanges;
      } else if (Array.isArray(line.editableRanges)) {
        ranges = line.editableRanges;
      }
    }

    return this._normalizeEditRanges(ranges);
  }

  _isXEditableForOperation(line, x, operation = 'move', precomputedRanges = null) {
    if (!line || line.editable === false) return false;
    if (!Number.isFinite(x)) return false;

    const ranges = precomputedRanges === null
      ? this._getEditRangesForOperation(line, operation)
      : precomputedRanges;

    // No point-level ranges configured => preserve legacy fully-editable behavior.
    if (ranges === null) return true;
    if (ranges.length === 0) return false;

    for (let i = 0; i < ranges.length; i++) {
      const [start, end] = ranges[i];
      if (x >= start && x <= end) {
        return true;
      }
    }
    return false;
  }

  _hasPointLevelEditConstraints(line) {
    return (
      this._getEditRangesForOperation(line, 'move') !== null ||
      this._getEditRangesForOperation(line, 'add') !== null ||
      this._getEditRangesForOperation(line, 'delete') !== null
    );
  }

  _resolveMaterializationBounds() {
    const axis = (this.baseOption && this.baseOption.xAxis) ? this.baseOption.xAxis : {};
    const axisMin = Number(axis.min);
    const axisMax = Number(axis.max);

    const xMin = Number.isFinite(axisMin)
      ? Math.ceil(axisMin)
      : (Number.isFinite(this.windowStartX) ? Math.ceil(this.windowStartX) : 0);
    const xMax = Number.isFinite(axisMax)
      ? Math.floor(axisMax)
      : (Number.isFinite(this.windowEndX) ? Math.floor(this.windowEndX) : Math.floor(this.globalMaxX));

    if (!Number.isFinite(xMin) || !Number.isFinite(xMax) || xMax < xMin) {
      return null;
    }

    return { xMin, xMax };
  }

  _materializeLockedDomainPoints() {
    if (!Array.isArray(this.lines) || this.lines.length === 0) return;

    const bounds = this._resolveMaterializationBounds();
    if (!bounds) return;
    const { xMin, xMax } = bounds;

    this.lines.forEach((line) => {
      if (!line || !Array.isArray(line.data) || line.data.length === 0) return;
      if (!this._hasPointLevelEditConstraints(line)) return;

      const sorted = line.data
        .filter((pt) => Array.isArray(pt) && pt.length >= 2 && Number.isFinite(Number(pt[0])) && Number.isFinite(Number(pt[1])))
        .map((pt) => [Number(pt[0]), Number(pt[1])])
        .sort((a, b) => a[0] - b[0]);
      if (!sorted.length) return;

      const dataMinX = Math.ceil(sorted[0][0]);
      const dataMaxX = Math.floor(sorted[sorted.length - 1][0]);
      const lockedMin = Math.max(xMin, dataMinX);
      const lockedMax = Math.min(xMax, dataMaxX);
      if (lockedMax < lockedMin) return;

      const byX = new Map();
      sorted.forEach(([px, py]) => byX.set(px, py));

      let changed = false;
      for (let x = lockedMin; x <= lockedMax; x++) {
        if (this._isXEditableForOperation(line, x, 'move')) continue;
        if (byX.has(x)) continue;
        byX.set(x, this._piecewiseLinear(sorted, x));
        changed = true;
      }

      if (changed) {
        line.data = Array.from(byX.entries()).sort((a, b) => a[0] - b[0]);
      }
    });
  }

  _clampXToEditableDomain(line, desiredX, operation, axisMin, axisMax, leftNeighborX = -Infinity, rightNeighborX = Infinity) {
    const ranges = this._getEditRangesForOperation(line, operation);
    if (ranges === null) return desiredX;

    const minAllowed = Math.ceil(Math.max(axisMin, leftNeighborX + 1));
    const maxAllowed = Math.floor(Math.min(axisMax, rightNeighborX - 1));
    if (!Number.isFinite(minAllowed) || !Number.isFinite(maxAllowed) || maxAllowed < minAllowed) {
      return null;
    }

    let bestX = null;
    let bestDist = Infinity;

    ranges.forEach(([start, end]) => {
      const lo = Math.ceil(Math.max(minAllowed, start));
      const hi = Math.floor(Math.min(maxAllowed, end));
      if (hi < lo) return;

      let candidate = desiredX;
      if (candidate < lo) candidate = lo;
      if (candidate > hi) candidate = hi;
      candidate = Math.round(candidate);
      if (candidate < lo) candidate = lo;
      if (candidate > hi) candidate = hi;

      const dist = Math.abs(candidate - desiredX);
      if (bestX === null || dist < bestDist || (dist === bestDist && candidate < bestX)) {
        bestX = candidate;
        bestDist = dist;
      }
    });

    return bestX;
  }

  /** Force re-render the chart from current lines data. */
  _refreshChart() {
    // We'll produce two ECharts series per user line.
    const seriesList = [];

    // 1) capture existing dataZoom so we preserve user's zoom
    this._captureZoomState();
    this._captureYAxisState(); // Capture Y-axis settings

    if (!this._isDragging) {
      this._materializeLockedDomainPoints();
    }

    // Make a copy of baseOption for final usage.
    const option = Object.assign({}, this.baseOption);

    // Decide the min & max for the X axis
    const xMin = 0;
    const xMax = this.globalMaxX; // e.g. 360

    // For each user line, create two series: 
    //  (1) "control" => actual points
    //  (2) "interpolated" => piecewise linear from xMin..xMax
    this.lines.forEach((ln) => {
      const moveRanges = this._getEditRangesForOperation(ln, 'move');
      const showControlSymbols = this.editMode && ln.editable !== false;
      const controlSymbolSize = showControlSymbols
        ? (
          moveRanges === null
            ? 10
            : ((value) => {
              const x = this._extractXFromDatum(value);
              if (!Number.isFinite(x)) return 10;
              return this._isXEditableForOperation(ln, x, 'move', moveRanges) ? 10 : 0;
            })
        )
        : 0;

      // Control series for the user's raw data
      seriesList.push({
        name: ln.name,
        type: 'line',
        data: ln.data,
        showSymbol: showControlSymbols,
        symbolSize: controlSymbolSize,
        lineStyle: {
          color: ln.color || '#000',
          width: 2
        },
        itemStyle: {
          color: ln.color || '#000'
        },
        tooltip: {
          show: false
        }
      });

      // Interpolated line
      const interpolatedData = this._interpolateLine(ln.data, xMin, xMax);
      seriesList.push({
        name: ln.name,
        type: 'line',
        data: interpolatedData,
        showSymbol: false,
        symbolSize: 0,
        itemStyle: {
          color: ln.color || '#000',
          opacity: 0.4
        },
        lineStyle: {
          color: ln.color || '#000',
          opacity: 0.4
        },
        tooltip: {
          show: true
        }
      });
    });

    // // Assign the new series list to the chart option
    option.series = seriesList;
    // this.chart.setOption(option, true);

    // 2) define dataZoom with our stored windowStartX..windowEndX
    if (this.enableZoomBar) {
      option.dataZoom = [{
        type: 'slider',
        xAxisIndex: 0,
        startValue: Math.ceil(this.windowStartX),
        endValue: Math.floor(this.windowEndX)
      }];
    }

    this.chart.setOption(option, true);

    // Fire the data-changed callback if it's defined
    if (typeof this.onDataChanged === 'function') {
      this.onDataChanged();
    }
  }

  /**
   * *Capture* current dataZoom from the chart so we can preserve it.
   */
  _captureZoomState() {
    // if no dataZoom yet, skip
    const opt = this.chart.getOption();
    if (!opt || !opt.dataZoom || !opt.dataZoom[0]) {
      return;
    }
    const dz = opt.dataZoom[0];
    if (typeof dz.startValue === 'number') {
      this.windowStartX = dz.startValue;
    }
    if (typeof dz.endValue === 'number') {
      this.windowEndX = dz.endValue;
    }
  }

  /**
   * *Capture* current Y-axis settings so we can preserve them.
   */
  _captureYAxisState() {
    const opt = this.chart.getOption();
    if (!opt || !opt.yAxis || !opt.yAxis[0]) {
      return;
    }
    const yAxis = opt.yAxis[0];
    if (typeof yAxis.min === 'number') {
      this.baseOption.yAxis.min = yAxis.min;
    }
    if (typeof yAxis.max === 'number') {
      this.baseOption.yAxis.max = yAxis.max;
    }
  }

  // --------------------------
  // INTERNAL EVENT HANDLING
  // --------------------------
  _initEvents() {
    const zr = this.chart.getZr();

    // For drag state
    this._isDragging = false;
    this._draggingLineIndex = null;
    this._draggingPointIndex = null;
    this._draggingSeriesIndex = null;
    this.chart.on('dataZoom', (params) => {
      // read the new window from chart options or from event params
      const opt = this.chart.getOption();
      if (opt.dataZoom && opt.dataZoom[0]) {
        const dz = opt.dataZoom[0];
        this.windowStartX = dz.startValue;
        this.windowEndX   = dz.endValue;
        }
    });
    

    zr.on('mousedown', (params) => {
      if (!this.editMode) return;
      if (params.event.shiftKey) {
        // SHIFT+Click => Possibly delete a point
        this._handleShiftClick(params);
        return;
      }
      this._handleMouseDown(params);
    });

    zr.on('mousemove', (params) => {
      if (!this.editMode) return;
      this._handleMouseMove(params);
    });

    zr.on('mouseup', (params) => {
      if (!this.editMode) return;
      this._handleMouseUp(params);
    });

    // "click" is fired if no drag occurred
    zr.on('click', (params) => {
      if (!this.editMode) return;
      if (this._isDragging) return;        // if we dragged a point, skip
      if (params.event.shiftKey) return;   // SHIFT logic handled in _handleShiftClick
      this._handleAddPoint(params);
    });
  }

  /**
   * SHIFT+Click on a point => remove it if line has > 2 points
   */
  _handleShiftClick(zrParams) {
    const pos = [zrParams.offsetX, zrParams.offsetY];
    const nearest = this._findNearestPoint(pos, 10, true, 'delete');
    if (!nearest) {
      return;
    }
    this._draggingSeriesIndex = nearest.lineIndex * 2;

    const { lineIndex, pointIndex } = nearest;
    const line = this.lines[lineIndex];
    const pointX = line && line.data && line.data[pointIndex] ? line.data[pointIndex][0] : NaN;
    if (!this._isXEditableForOperation(line, pointX, 'delete')) {
      return;
    }
    if (line.data.length <= 2) {
      // Do not allow removing if line would have < 2 points
      return;
    }

    this._pushHistory();
    line.data.splice(pointIndex, 1);
    this._refreshChart();
  }

  /**
   * MOUSEDOWN: check if user is on top of a point => begin dragging
   */
  _handleMouseDown(zrParams) {
    const pos = [zrParams.offsetX, zrParams.offsetY];
    const nearest = this._findNearestPoint(pos, 10, true, 'move');
    if (!nearest) return;

    // Begin drag
    this._isDragging = true;
    this._draggingLineIndex = nearest.lineIndex;
    this._draggingPointIndex = nearest.pointIndex;
    this._draggingSeriesIndex = nearest.lineIndex * 2;

    this._pushHistory();
  }

  /**
   * MOUSEMOVE: if dragging, update the point's position
   */
  _handleMouseMove(zrParams) {
    if (!this._isDragging) return;

    const pos = [zrParams.offsetX, zrParams.offsetY];
    let [chartX, chartY] = this.chart.convertFromPixel(
      { xAxisIndex: 0, yAxisIndex: 0 },
      pos
    );

    // Snap X to integer
    let snappedX = Math.round(chartX);

    // Clamp to axis
    const xMin = Math.ceil(this.windowStartX);
    const xMax = Math.floor(this.windowEndX);
    const yMin = this.baseOption.yAxis.min ?? 0;
    const yMax = this.baseOption.yAxis.max ?? 100;

    if (snappedX < xMin) snappedX = xMin;
    if (snappedX > xMax) snappedX = xMax;
    if (chartY < yMin) chartY = yMin;
    if (chartY > yMax) chartY = yMax;

    const line = this.lines[this._draggingLineIndex];
    const idx = this._draggingPointIndex;
    if (!line || !line.data || !line.data[idx]) return;

    const currentPointX = line.data[idx][0];
    if (!this._isXEditableForOperation(line, currentPointX, 'move')) {
      return;
    }

    // Preserve sorted X (avoid crossing neighbors)
    const leftNeighborX = idx > 0 ? line.data[idx - 1][0] : -Infinity;
    const rightNeighborX = idx < line.data.length - 1 ? line.data[idx + 1][0] : +Infinity;

    if (snappedX <= leftNeighborX) {
      snappedX = leftNeighborX + 1;
    } else if (snappedX >= rightNeighborX) {
      snappedX = rightNeighborX - 1;
    }

    const constrainedX = this._clampXToEditableDomain(
      line,
      snappedX,
      'move',
      xMin,
      xMax,
      leftNeighborX,
      rightNeighborX
    );
    if (!Number.isFinite(constrainedX)) {
      return;
    }
    snappedX = constrainedX;

    // Enforce monotonic constraints
    let newY = chartY;
    if (this.monotonic === 'ascending') {
      const leftNeighborY = idx > 0 ? line.data[idx - 1][1] : -Infinity;
      const rightNeighborY = idx < line.data.length - 1 ? line.data[idx + 1][1] : Infinity;
      if (newY < leftNeighborY) newY = leftNeighborY;
      if (newY > rightNeighborY) newY = rightNeighborY;
    } else if (this.monotonic === 'descending') {
      const leftNeighborY = idx > 0 ? line.data[idx - 1][1] : Infinity;
      const rightNeighborY = idx < line.data.length - 1 ? line.data[idx + 1][1] : -Infinity;
      if (newY > leftNeighborY) newY = leftNeighborY;
      if (newY < rightNeighborY) newY = rightNeighborY;
    }

    // Figure out how many digits to use
    const digits = getPrecisionForMaxY(yMax);
    // Round the Y
    newY = roundToScale(newY, digits);

    // Now store in line.data
    line.data[idx] = [snappedX, newY];
    // Re-sort, find new index
    line.data.sort((a, b) => a[0] - b[0]);
    this._draggingPointIndex = line.data.findIndex(
      (pt) => pt[0] === snappedX && pt[1] === newY
    );

    // Re-render
    this._refreshChart();

    // ----- Show/Update tooltip while dragging -----
    // If 'axis' tooltip, we can update the axis pointer:
    this.chart.dispatchAction({
      type: 'updateAxisPointer',
      xAxisIndex: 0,
      value: snappedX
    });

    // If you want an "item" tooltip for the exact data point:
    this.chart.dispatchAction({
      type: 'showTip',
      seriesIndex: this._draggingSeriesIndex,
      dataIndex: this._draggingPointIndex
    });
  }

  /**
   * MOUSEUP: finalize drag
   */
  _handleMouseUp(zrParams) {
    if (this._isDragging) {
      this._isDragging = false;
      this._draggingLineIndex = null;
      this._draggingPointIndex = null;
    }
  }

  /**
   * CLICK: if user clicked on blank area => add a new point near the nearest line
   */
  _handleAddPoint(zrParams) {
    const pos = [zrParams.offsetX, zrParams.offsetY];
    let [chartX, chartY] = this.chart.convertFromPixel(
      { xAxisIndex: 0, yAxisIndex: 0 },
      pos
    );

    // Hard-coded axis limits
    const xMin = this.windowStartX ?? 0;
    const xMax = this.windowEndX ?? 100;
    const yMin = this.baseOption.yAxis.min ?? 0;
    const yMax = this.baseOption.yAxis.max ?? 100;

    if (chartX < xMin || chartX > xMax || chartY < yMin || chartY > yMax) {
      return; // out of bounds => do nothing
    }

    const snappedX = Math.round(chartX);

    // Find nearest line at that X
    let bestDist = Infinity;
    let bestLineIndex = -1;
    this.lines.forEach((line, i) => {
      if (!this._isXEditableForOperation(line, snappedX, 'add')) return;
      const dist = this._verticalDistanceToLine(line.data, snappedX, chartY);
      if (dist < bestDist) {
        bestDist = dist;
        bestLineIndex = i;
      }
    });
    if (bestLineIndex === -1) return;

    if (!this._isXEditableForOperation(this.lines[bestLineIndex], snappedX, 'add')) {
      return; // Block add
    }

    const theLine = this.lines[bestLineIndex];
    // If a point at that X already exists => do not add
    if (theLine.data.some(pt => pt[0] === snappedX)) {
      return;
    }

    // Simulate adding the new point
    // Snap Y using same logic as drag: step based on yMax (3rd significant digit 0/5)
    const digits = getPrecisionForMaxY(yMax);
    const clampedY = this._clampNewPointY(theLine, snappedX, chartY);
    const snappedY = roundToScale(clampedY, digits);
    // Proceed with adding the point (no cross-check between lines)
    this._pushHistory();
    theLine.data.push([snappedX, snappedY]);
    theLine.data.sort((a, b) => a[0] - b[0]);
    this._refreshChart();
  }

  /**
   * Find nearest existing data point within pixelTolerance
   */
  _findNearestPoint(pixelPos, pixelTolerance, requireEditable = false, operation = 'move') {
    let nearest = null;
    let minDist = Infinity;
    // Iterate over lines
    for (let i = this.lines.length - 1; i >= 0; i--) {
      const line = this.lines[i];
      if (requireEditable && line && line.editable === false) {
        continue;
      }
      for (let j = 0; j < line.data.length; j++) {
        const pt = line.data[j];
        if (requireEditable && !this._isXEditableForOperation(line, pt[0], operation)) {
          continue;
        }
        const pixel = this.chart.convertToPixel({ xAxisIndex: 0, yAxisIndex: 0 }, pt);
        const dx = pixel[0] - pixelPos[0];
        const dy = pixel[1] - pixelPos[1];
        const dist = Math.sqrt(dx*dx + dy*dy);
        if (dist < minDist && dist <= pixelTolerance) {
          minDist = dist;
          nearest = { lineIndex: i, pointIndex: j, dist };
        }
      }
    }
    return nearest;
  }

  /**
   * Returns approximate vertical distance from (x,y) to line segments
   */
  _verticalDistanceToLine(lineData, x, y) {
    if (lineData.length === 0) return Infinity;
    if (x <= lineData[0][0]) {
      return Math.abs(y - lineData[0][1]);
    }
    if (x >= lineData[lineData.length - 1][0]) {
      return Math.abs(y - lineData[lineData.length - 1][1]);
    }
    for (let i = 0; i < lineData.length - 1; i++) {
      const [x1,y1] = lineData[i];
      const [x2,y2] = lineData[i+1];
      if (x1 <= x && x <= x2) {
        const t = (x - x1)/(x2 - x1);
        const interpY = y1 + t*(y2 - y1);
        return Math.abs(y - interpY);
      }
    }
    return Infinity;
  }

  /**
   * If monotonic is ascending or descending, clamp new Y
   */
  _clampNewPointY(line, newX, rawY) {
    if (this.monotonic === 'none' || line.data.length < 2) {
      return rawY;
    }
    const data = line.data;
    if (newX <= data[0][0]) {
      if (this.monotonic === 'ascending') {
        return Math.max(rawY, data[0][1]);
      } else {
        return Math.min(rawY, data[0][1]);
      }
    }
    if (newX >= data[data.length - 1][0]) {
      if (this.monotonic === 'ascending') {
        return Math.min(rawY, data[data.length - 1][1]);
      } else {
        return Math.max(rawY, data[data.length - 1][1]);
      }
    }
    for (let i = 0; i < data.length - 1; i++) {
      const [x1, y1] = data[i];
      const [x2, y2] = data[i+1];
      if (x1 <= newX && newX <= x2) {
        if (this.monotonic === 'ascending') {
          const lo = Math.min(y1, y2);
          const hi = Math.max(y1, y2);
          return Math.max(lo, Math.min(hi, rawY));
        } else {
          const lo = Math.min(y1, y2);
          const hi = Math.max(y1, y2);
          return Math.min(hi, Math.max(lo, rawY));
        }
      }
    }
    return rawY;
  }
}

/**
 * Rounds `value` to `digits` decimal places if digits >= 0,
 * or to tens/hundreds if digits < 0.
 * e.g. digits=-1 => round to nearest 10
 *      digits=-2 => round to nearest 100
 */
function roundToScale(value, expYMax) {
  // Snap based on yMax: force the 3rd significant digit of yMax to {0,5}.
  // expYMax should be floor(log10(yMax)).
  if (!isFinite(value) || !isFinite(expYMax)) return value;
  const step = 5 * (10 ** (expYMax - 2));
  if (step === 0) return value;
  return Math.round(value / step) * step;
}

/**
 * Decide how many digits to keep based on the max Y.
 */
function getPrecisionForMaxY(maxY) {
  // Return the exponent floor(log10(maxY)) to drive the 0/5 snap step in roundToScale.
  if (!isFinite(maxY) || maxY <= 0) return 0;
  return Math.floor(Math.log10(maxY));
}


/**
 * SegmentSliderChart: a minimal ECharts-based multi-segment slider.
 *
 * - xAxis in [0..dayCount].
 * - boundaries[]: sorted array of boundary positions => define segments.
 * - Each segment is drawn as a line from boundary[i]..boundary[i+1], each with a color in segmentColors[i].
 * - A scatter series shows the boundary knobs.
 * - SHIFT+click a knob => remove it (except pinned ends).
 * - Normal click in empty space => add boundary.
 * - Drag a knob horizontally => updates boundary in real-time.
 */
class SegmentSliderChart {
  /**
   * @param {string} domId - The container <div> ID for ECharts
   * @param {number} dayCount - The maximum domain on x-axis (e.g. 50, 60, etc.)
   * @param {number[]} boundaries - Initial boundary array, e.g. [0, 50]. Must contain at least [0,dayCount].
   * @param {string[]} segmentColors - optional array of color strings for each segment
   */
  constructor(domId, dayCount = 50, boundaries = [0, 50], segmentColors = []) {
    // The container for the ECharts instance
    this.container = document.getElementById(domId);
    this.dayCount = dayCount;

    // Boundaries: respect provided domain
    this.boundaries = boundaries.slice();
    if (this.boundaries.length === 0) {
      this.boundaries = [0, this.dayCount];
    }
    this.boundaries.sort((a, b) => a - b);
    this.domainMin = this.boundaries[0];
    this.domainMax = this.boundaries[this.boundaries.length - 1];
    if (!this.boundaries.includes(this.domainMin)) {
      this.boundaries.push(this.domainMin);
    }
    if (!this.boundaries.includes(this.domainMax)) {
      this.boundaries.push(this.domainMax);
    }
    this.boundaries.sort((a, b) => a - b);

    // Each segment i => boundaries[i]..boundaries[i+1], with color segmentColors[i]
    this.segmentColors = segmentColors.slice();

    // If we have (n-1) segments, ensure we have that many colors
    const segCount = this.boundaries.length - 1;
    while (this.segmentColors.length < segCount) {
      this.segmentColors.push(this._randomColor());
    }
    this.segmentIds = [];
    for (let i = 0; i < segCount; i++) {
      this.segmentIds.push(this._newSegmentId());
    }
    // Boundary ownership: internal boundaries belong to left segment by default
    this.boundaryOwners = new Array(this.boundaries.length).fill(null);
    for (let i = 1; i < this.boundaries.length - 1; i++) {
      this.boundaryOwners[i] = 'left';
    }

    // ECharts init
    this.chart = echarts.init(this.container);

    // Track which boundary is being dragged
    this._draggingIndex = -1;

    // Optional callbacks
    this.onChange = null;    // Fired whenever boundaries change
    this.onSelectSeg = null; // Fired if user clicks a segment
    this.onDataChanged = null; // Fired whenever data changes (similar to NapkinChart)

    // Attach event handlers (for dragging, SHIFT-click removal, etc.)
    this._initEvents();
    // Render first time
    this._refreshChart();
  }

  /**
   * Builds and sets the ECharts option from current boundaries + colors
   */
  _refreshChart() {
    // 1) Build line series for each segment
    const lineSeries = [];
    const n = this.boundaries.length;

    for (let i = 0; i < n - 1; i++) {
      const startX = this.boundaries[i];
      const endX = this.boundaries[i + 1];
      const col = this.segmentColors[i] || '#007bff';

      lineSeries.push({
        name: `Segment${i}`,
        type: 'line',
        data: [
          [startX, 0],
          [endX, 0],
        ],
        showSymbol: false,
        lineStyle: {
          width: 6,
          color: col,
        },
        tooltip: { show: false },
        z: 1,
      });
    }

    // 2) Build scatter series for boundary knobs
    const scatterData = [];
    for (let i = 0; i < n; i++) {
      if (i === 0 || i === n - 1) continue;
      const xVal = this.boundaries[i];
      let knobColor = '#007bff';
      // We color the knob based on which side owns the boundary
      if (i > 0) {
        const owner = this.boundaryOwners[i];
        if (owner === 'right') {
          knobColor = this.segmentColors[i] || this.segmentColors[i - 1] || '#007bff';
        } else {
          knobColor = this.segmentColors[i - 1] || '#007bff';
        }
      }
      scatterData.push({
        value: [xVal, 0],
        itemStyle: {
          color: knobColor,
          borderColor: '#fff',
          borderWidth: 2,
        },
      });
    }

    const knobSeries = {
      name: 'Boundaries',
      type: 'scatter',
      data: scatterData,
      symbolSize: 14,
      tooltip: { show: false },
      z: 3,
    };

    // Combine
    const seriesList = [...lineSeries, knobSeries];

    // 3) Basic ECharts option
    const xMin = this.boundaries.length ? this.boundaries[0] : 0;
    const xMax = this.boundaries.length ? this.boundaries[this.boundaries.length - 1] : this.dayCount;
    const option = {
      animation: false,
      xAxis: {
        type: 'value',
        min: this.domainMin,
        max: this.domainMax,
        axisLabel: {
          formatter: v => String(v),
        },
      },
      yAxis: {
        show: false,
        min: -1,
        max: 1,
      },
      grid: {
        left: 40,
        right: 20,
        top: 20,
        bottom: 40,
      },
      tooltip: { trigger: 'axis' },
      series: seriesList,
    };

    this.chart.setOption(option, true);

    // Fire the data-changed callback if it's defined
    if (typeof this.onDataChanged === 'function') {
      this.onDataChanged();
    }
  }

  /**
   * Low-level ZRender events for drag, SHIFT+click, etc.
   */
  _initEvents() {
    const zr = this.chart.getZr();

    zr.on('mousedown', (params) => {
      const [cx, cy] = this._pxToData(params.offsetX, params.offsetY);
      // find nearest boundary
      const near = this._findNearestBoundary(cx);
      if (near.dist < 0.5) {
        // user clicked a boundary knob
        if (params.event.shiftKey) {
          // SHIFT => remove boundary if not pinned
          this._removeBoundaryAtIndex(near.index);
        } else {
          // drag if not pinned
          if (
            near.index === 0 ||
            near.index === this.boundaries.length - 1
          ) {
            // pinned => do nothing
            return;
          }
          this._draggingIndex = near.index;
          this._dragStartDay = this.boundaries[near.index];
          this._dragMoved = false;
        }
      } else {
        // SHIFT => ignore or do something else
        if (!params.event.shiftKey) {
          // add boundary if within domain
          let day = Math.round(cx);
          if (day < this.domainMin) day = this.domainMin;
          if (day > this.domainMax) day = this.domainMax;
          if (!this.boundaries.includes(day)) {
            this._addBoundary(day);
          } else {
            // user might have clicked a segment
            const segIndex = this._findSegmentIndex(day);
            if (this.onSelectSeg) {
              this.onSelectSeg(segIndex);
            }
          }
        }
      }
    });

    zr.on('mousemove', (params) => {
      if (this._draggingIndex < 0) return;
      // we are dragging boundary this._draggingIndex
      const [cx, cy] = this._pxToData(params.offsetX, params.offsetY);
      let day = Math.round(cx);
      if (day < this.domainMin) day = this.domainMin;
      if (day > this.domainMax) day = this.domainMax;

      // avoid crossing neighbors
      const i = this._draggingIndex;
      const leftB = this.boundaries[i - 1];
      const rightB = this.boundaries[i + 1];
      if (day <= leftB) day = leftB + 1;
      if (day >= rightB) day = rightB - 1;

      if (this.boundaries[i] !== day) {
        this._dragMoved = true;
      }
      this.boundaries[i] = day;
      this._refreshChart();

      if (this.onChange) {
        this.onChange();
      }
    });

    zr.on('mouseup', (params) => {
      if (this._draggingIndex >= 0) {
        if (!this._dragMoved) {
          this._toggleBoundaryOwner(this._draggingIndex);
        }
        this._draggingIndex = -1;
        this._dragStartDay = null;
        this._dragMoved = false;
      }
    });
  }

  /**
   * Convert from pixel => [x, y] in data coords
   */
  _pxToData(px, py) {
    return this.chart.convertFromPixel({ xAxisIndex: 0, yAxisIndex: 0 }, [
      px,
      py,
    ]);
  }

  /**
   * Finds nearest boundary in x-dimension
   */
  _findNearestBoundary(x) {
    let bestDist = Infinity;
    let bestIndex = -1;
    for (let i = 0; i < this.boundaries.length; i++) {
      const bx = this.boundaries[i];
      const dist = Math.abs(bx - x);
      if (dist < bestDist) {
        bestDist = dist;
        bestIndex = i;
      }
    }
    return { index: bestIndex, dist: bestDist };
  }

  /**
   * find which segment day belongs to (0-based)
   */
  _findSegmentIndex(day) {
    for (let i = 0; i < this.boundaries.length - 1; i++) {
      if (day >= this.boundaries[i] && day <= this.boundaries[i + 1]) {
        return i;
      }
    }
    return -1;
  }

  /**
   * Insert a new boundary
   */
  _addBoundary(day) {
    this.boundaries.push(day);
    this.boundaries.sort((a, b) => a - b);

    // We'll generate a new color for the newly split segment
    const idx = this.boundaries.indexOf(day);
    const newCol = this._randomColor();
    // If we had n-1 segments, we now have n => insert color
    // E.g. if segment i was split, we keep the old color on the left,
    // and insert new color for the right.
    this.segmentColors.splice(idx, 0, newCol);
    this.segmentIds.splice(idx, 0, this._newSegmentId());
    this.boundaryOwners.splice(idx, 0, 'left');

    this._refreshChart();
    if (this.onChange) {
      this.onChange();
    }
  }

  /**
   * Remove boundary by index (unless pinned)
   */
  _removeBoundaryAtIndex(i) {
    if (i === 0 || i === this.boundaries.length - 1) {
      // pinned => do nothing
      return;
    }
    this.boundaries.splice(i, 1);
    // unify segments i-1 and i => keep left color, drop right color
    this.segmentColors.splice(i, 1);
    this.segmentIds.splice(i, 1);
    this.boundaryOwners.splice(i, 1);

    this._refreshChart();
    if (this.onChange) {
      this.onChange();
    }
  }

  /**
   * If user wants to remove boundary by value
   */
  removeBoundary(day) {
    const i = this.boundaries.indexOf(day);
    if (i >= 0) {
      this._removeBoundaryAtIndex(i);
    }
  }

  /**
   * Generate random color
   */
  _randomColor() {
    const r = Math.floor(Math.random() * 256);
    const g = Math.floor(Math.random() * 256);
    const b = Math.floor(Math.random() * 256);
    return `rgb(${r},${g},${b})`;
  }

  _newSegmentId() {
    const rand = Math.random().toString(36).slice(2, 8);
    return `seg-${Date.now().toString(36)}-${rand}`;
  }

  getSegments() {
    const segments = [];
    const owners = this.boundaryOwners || [];
    for (let i = 0; i < this.boundaries.length - 1; i++) {
      let start = this.boundaries[i];
      let end = this.boundaries[i + 1];
      if (i > 0 && owners[i] === 'left') {
        start = start + 1;
      }
      if (owners[i + 1] === 'right') {
        end = end - 1;
      }
      if (start > end) {
        start = end;
      }
      segments.push({
        id: this.segmentIds[i],
        start,
        end,
        color: this.segmentColors[i],
      });
    }
    return segments;
  }

  _toggleBoundaryOwner(i) {
    if (i <= 0 || i >= this.boundaries.length - 1) return;
    const current = this.boundaryOwners[i] || 'left';
    this.boundaryOwners[i] = current === 'left' ? 'right' : 'left';
    this._refreshChart();
    if (this.onChange) {
      this.onChange();
    }
  }
  colorWheelNonRandom() {
    const colors = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'];
    return colors[this.boundaries.length % colors.length];
  }
}

class NapkinChartArea extends NapkinChart {
  constructor(domId, initialLines, editMode = true, baseOption = null, monotonic = 'none') {
    super(domId, initialLines, editMode, baseOption, monotonic);
    this.topAreaColor = (baseOption && baseOption.topAreaColor) || '#dfe3ea';
    this.topAreaLabel = (baseOption && baseOption.topAreaLabel) || 'Top';
    this.areaTint = (baseOption && baseOption.areaTint) ?? 0.35;
  }

  _tintColor(color, tint = 0.35) {
    if (!color || typeof color !== 'string') return color;
    let r, g, b;
    if (color.startsWith('#')) {
      const hex = color.slice(1);
      if (hex.length === 3) {
        r = parseInt(hex[0] + hex[0], 16);
        g = parseInt(hex[1] + hex[1], 16);
        b = parseInt(hex[2] + hex[2], 16);
      } else if (hex.length === 6) {
        r = parseInt(hex.slice(0, 2), 16);
        g = parseInt(hex.slice(2, 4), 16);
        b = parseInt(hex.slice(4, 6), 16);
      } else {
        return color;
      }
    } else if (color.startsWith('rgb')) {
      const nums = color.match(/[\d.]+/g);
      if (!nums || nums.length < 3) return color;
      r = Number(nums[0]);
      g = Number(nums[1]);
      b = Number(nums[2]);
    } else {
      return color;
    }

    const mix = (c) => Math.round(c + (255 - c) * tint);
    return `rgb(${mix(r)},${mix(g)},${mix(b)})`;
  }

  _refreshChart() {
    const seriesList = [];

    this._captureZoomState();
    this._captureYAxisState();
    if (!this._isDragging) {
      this._materializeLockedDomainPoints();
    }

    const option = Object.assign({}, this.baseOption);
    option.tooltip = Object.assign({}, option.tooltip, {
      trigger: 'axis',
      axisPointer: { type: 'line' },
      formatter: (params) => {
        const order = Array.isArray(this._bandOrder) ? this._bandOrder : [];
        const bandSet = new Set(order);
        const bandItems = params.filter((p) => bandSet.has(p.seriesName));
        if (!bandItems.length) return '';
        const byName = {};
        bandItems.forEach((p) => {
          const val = Array.isArray(p.data) ? p.data[1] : p.value;
          byName[p.seriesName] = { val, marker: p.marker || '' };
        });
        const lines = [];
        let prev = 0;
        for (const name of order) {
          const entry = byName[name];
          if (!entry) continue;
          const height = (entry.val ?? 0) - prev;
          prev = entry.val ?? 0;
          const rounded = Number.isFinite(height) ? height.toFixed(3) : height;
          lines.push(`${entry.marker} ${name}: ${rounded}`);
        }
        const xLabel = bandItems[0].axisValueLabel ?? bandItems[0].axisValue ?? '';
        return [
          `x: ${xLabel}`,
          ...lines.reverse(),
        ].join('<br/>');
      }
    });
    const axisMin = this.baseOption?.xAxis?.min;
    const axisMax = this.baseOption?.xAxis?.max;
    const xMin = Number.isFinite(axisMin)
      ? axisMin
      : Number.isFinite(this.windowStartX)
        ? Math.ceil(this.windowStartX)
        : 0;
    const xMax = Number.isFinite(axisMax)
      ? axisMax
      : Number.isFinite(this.windowEndX)
        ? Math.floor(this.windowEndX)
        : this.globalMaxX;

    const yMin = this.baseOption.yAxis?.min ?? 0;
    const yMax = this.baseOption.yAxis?.max ?? 100;

    if (this.lines.length > 0 && isFinite(xMax) && xMax >= xMin) {
      const xGrid = [];
      for (let x = xMin; x <= xMax; x++) xGrid.push(x);
      const interpolated = this.lines.map(ln =>
        this._interpolateLine(ln.data, xMin, xMax).map(([x, y]) => y)
      );
      const order = interpolated
        .map((vals, i) => {
          const sum = vals.reduce((acc, v) => acc + (isFinite(v) ? v : 0), 0);
          const avg = vals.length ? sum / vals.length : 0;
          return { i, avg };
        })
        .sort((a, b) => b.avg - a.avg)
        .map(item => item.i);
      const orderedInterpolated = order.map(i => interpolated[i]);
      const orderedLines = order.map(i => this.lines[i]);

      const clamp = (value) => {
        if (!isFinite(value)) return 0;
        return Math.max(0, value);
      };

      const toBandSeries = (name, color, values, z) => ({
        name,
        type: 'line',
        data: xGrid.map((x, i) => [x, clamp(values[i])]),
        showSymbol: false,
        symbolSize: 0,
        lineStyle: { width: 0 },
        itemStyle: { color: color || '#999' },
        areaStyle: { opacity: 1, color: this._tintColor(color || '#999', this.areaTint) },
        tooltip: { show: true },
        emphasis: { disabled: true },
        silent: true,
        z,
      });

      const n = orderedInterpolated.length;
      const bands = [];
      // Bottom band: yMin -> lowest line
      bands.push({
        name: `${orderedLines[n - 1].name || 'Bottom'}`,
        color: orderedLines[n - 1].color,
        height: xGrid.map((x, i) => orderedInterpolated[n - 1][i] - yMin),
      });
      // Middle bands: between each pair of lines (bottom to top)
      for (let i = n - 2; i >= 0; i--) {
        bands.push({
          name: `${orderedLines[i].name || 'Band'}`,
          color: orderedLines[i].color,
          height: xGrid.map((x, idx) => orderedInterpolated[i][idx] - orderedInterpolated[i + 1][idx]),
        });
      }
      // Top band: top line -> yMax
      bands.push({
        name: this.topAreaLabel,
        color: this.topAreaColor,
        height: xGrid.map((x, i) => yMax - orderedInterpolated[0][i]),
      });

      const cumulative = [];
      let running = xGrid.map(() => 0);
      for (const band of bands) {
        running = running.map((v, i) => v + band.height[i]);
        cumulative.push(running);
      }

      // Draw from top to bottom so lower bands remain visible
      // Keep bands behind point symbols (control points use z=3)
      const zBase = -bands.length;
      for (let i = bands.length - 1; i >= 0; i--) {
        const z = zBase + (bands.length - 1 - i);
        seriesList.push(toBandSeries(bands[i].name, bands[i].color, cumulative[i], z));
      }

      this._bandOrder = bands.map((band) => band.name);
    }

    this._bandSeriesCount = this.lines.length > 0 ? this.lines.length + 1 : 0;

    // Line series (control + interpolated) so drag handles remain visible
    this.lines.forEach((ln) => {
      const moveRanges = this._getEditRangesForOperation(ln, 'move');
      const showControlSymbols = this.editMode && ln.editable !== false;
      const controlSymbolSize = showControlSymbols
        ? (
          moveRanges === null
            ? 10
            : ((value) => {
              const x = this._extractXFromDatum(value);
              if (!Number.isFinite(x)) return 10;
              return this._isXEditableForOperation(ln, x, 'move', moveRanges) ? 10 : 0;
            })
        )
        : 0;

      seriesList.push({
        name: ln.name,
        type: 'line',
        data: ln.data,
        showSymbol: showControlSymbols,
        symbolSize: controlSymbolSize,
        lineStyle: {
          color: ln.color || '#000',
          width: 2,
        },
        itemStyle: {
          color: ln.color || '#000',
        },
        tooltip: {
          show: false,
        },
        z: 3,
      });

      const interpolatedData = this._interpolateLine(ln.data, xMin, xMax);
      seriesList.push({
        name: ln.name,
        type: 'line',
        data: interpolatedData,
        showSymbol: false,
        symbolSize: 0,
        itemStyle: {
          color: ln.color || '#000',
          opacity: 0.4,
        },
        lineStyle: {
          color: ln.color || '#000',
          opacity: 0.4,
        },
        tooltip: {
          show: false,
        },
        z: 2,
      });
    });

    option.series = seriesList;

    if (this.enableZoomBar) {
      option.dataZoom = [{
        type: 'slider',
        xAxisIndex: 0,
        startValue: Math.ceil(this.windowStartX),
        endValue: Math.floor(this.windowEndX),
      }];
    }

    this.chart.setOption(option, true);

    if (typeof this.onDataChanged === 'function') {
      this.onDataChanged();
    }
  }

  _handleMouseDown(zrParams) {
    super._handleMouseDown(zrParams);
    if (this._isDragging) {
      const offset = this._bandSeriesCount || 0;
      this._draggingSeriesIndex = offset + (this._draggingLineIndex * 2);
    }
  }

  _handleMouseMove(zrParams) {
    if (!this._isDragging) return;

    // Prevent dragging if the current line is not editable
    if (
      this.lines[this._draggingLineIndex] &&
      this.lines[this._draggingLineIndex].editable === false
    ) {
      return;
    }

    const pos = [zrParams.offsetX, zrParams.offsetY];
    let [chartX, chartY] = this.chart.convertFromPixel(
      { xAxisIndex: 0, yAxisIndex: 0 },
      pos
    );

    // Snap X to integer
    let snappedX = Math.round(chartX);

    // Clamp to axis
    const xMin = Math.ceil(this.windowStartX);
    const xMax = Math.floor(this.windowEndX);
    const yMin = this.baseOption.yAxis.min ?? 0;
    const yMax = this.baseOption.yAxis.max ?? 100;

    if (snappedX < xMin) snappedX = xMin;
    if (snappedX > xMax) snappedX = xMax;
    if (chartY < yMin) chartY = yMin;
    if (chartY > yMax) chartY = yMax;

    const line = this.lines[this._draggingLineIndex];
    const idx = this._draggingPointIndex;
    if (!line || !line.data || !line.data[idx]) return;

    const currentPointX = line.data[idx][0];
    if (!this._isXEditableForOperation(line, currentPointX, 'move')) {
      return;
    }

    const currentX = line.data[idx][0];

    // If the point is at the left edge, always keep it at xMin
    if (currentX === xMin) {
      snappedX = xMin;
    }
    // If the point is at the right edge, always keep it at xMax
    if (currentX === xMax) {
      snappedX = xMax;
    }

    // Preserve sorted X (avoid crossing neighbors in X)
    const leftNeighborX = idx > 0 ? line.data[idx - 1][0] : -Infinity;
    const rightNeighborX = idx < line.data.length - 1 ? line.data[idx + 1][0] : Infinity;
    if (snappedX <= leftNeighborX) snappedX = leftNeighborX + 1;
    if (snappedX >= rightNeighborX) snappedX = rightNeighborX - 1;

    const constrainedX = this._clampXToEditableDomain(
      line,
      snappedX,
      'move',
      xMin,
      xMax,
      leftNeighborX,
      rightNeighborX
    );
    if (!Number.isFinite(constrainedX)) {
      return;
    }
    snappedX = constrainedX;

    // Enforce monotonic constraints (same as base)
    let newY = chartY;
    if (this.monotonic === 'ascending') {
      const leftNeighborY = idx > 0 ? line.data[idx - 1][1] : -Infinity;
      const rightNeighborY = idx < line.data.length - 1 ? line.data[idx + 1][1] : Infinity;
      if (newY < leftNeighborY) newY = leftNeighborY;
      if (newY > rightNeighborY) newY = rightNeighborY;
    } else if (this.monotonic === 'descending') {
      const leftNeighborY = idx > 0 ? line.data[idx - 1][1] : Infinity;
      const rightNeighborY = idx < line.data.length - 1 ? line.data[idx + 1][1] : -Infinity;
      if (newY > leftNeighborY) newY = leftNeighborY;
      if (newY < rightNeighborY) newY = rightNeighborY;
    }

    // Clamp to adjacent lines (dragged point only)
    const aboveLine = this._draggingLineIndex > 0 ? this.lines[this._draggingLineIndex - 1] : null;
    const belowLine = this._draggingLineIndex < this.lines.length - 1 ? this.lines[this._draggingLineIndex + 1] : null;
    let clampedToNeighbor = false;
    if (aboveLine) {
      const aboveSorted = aboveLine.data.slice().sort((a, b) => a[0] - b[0]);
      const yAbove = this._piecewiseLinear(aboveSorted, snappedX);
      if (newY > yAbove) {
        newY = yAbove;
        clampedToNeighbor = true;
      }
    }
    if (belowLine) {
      const belowSorted = belowLine.data.slice().sort((a, b) => a[0] - b[0]);
      const yBelow = this._piecewiseLinear(belowSorted, snappedX);
      if (newY < yBelow) {
        newY = yBelow;
        clampedToNeighbor = true;
      }
    }

    // Snap Y scale
    const digits = getPrecisionForMaxY(yMax);
    newY = clampedToNeighbor ? newY : roundToScale(newY, digits);

    // Simulate the new data (do NOT update the real data yet)
    const newData = line.data
      .map((pt, i) => i === idx ? [snappedX, newY] : pt)
      .sort((a, b) => a[0] - b[0]);
    const allLines = this.lines.map((ln, i) =>
      i === this._draggingLineIndex ? { ...ln, data: newData } : ln
    );

    // Interpolate all lines at all integer x
    const xGrid = [];
    for (let x = xMin; x <= xMax; x++) xGrid.push(x);
    const interpolated = allLines.map(ln =>
      this._interpolateLine(ln.data, xMin, xMax).map(([x, y]) => y)
    );

    // Check for crossing at any x (for any number of lines)
    for (let xi = 0; xi < xGrid.length; xi++) {
      for (let i = 0; i < interpolated.length - 1; i++) {
        if (interpolated[i][xi] < interpolated[i + 1][xi]) {
          // Block the drag if any line is less than the one below it
          return;
        }
      }
    }

    // Update data point
    line.data[idx] = [snappedX, newY];
    line.data.sort((a, b) => a[0] - b[0]);
    this._draggingPointIndex = line.data.findIndex(
      (pt) => pt[0] === snappedX && pt[1] === newY
    );

    this._refreshChart();

    this.chart.dispatchAction({
      type: 'updateAxisPointer',
      xAxisIndex: 0,
      value: snappedX
    });

    this.chart.dispatchAction({
      type: 'showTip',
      seriesIndex: this._draggingSeriesIndex,
      dataIndex: this._draggingPointIndex
    });

    this._ensureEdgePoints(line, xMin, xMax);
  }

  _handleAddPoint(zrParams) {
    const pos = [zrParams.offsetX, zrParams.offsetY];
    let [chartX, chartY] = this.chart.convertFromPixel(
      { xAxisIndex: 0, yAxisIndex: 0 },
      pos
    );

    // Hard-coded axis limits
    const xMin = this.windowStartX ?? 0;
    const xMax = this.windowEndX ?? 100;
    const yMin = this.baseOption.yAxis.min ?? 0;
    const yMax = this.baseOption.yAxis.max ?? 100;

    if (chartX < xMin || chartX > xMax || chartY < yMin || chartY > yMax) {
      return; // out of bounds => do nothing
    }

    const snappedX = Math.round(chartX);

    // Find nearest line at that X
    let bestDist = Infinity;
    let bestLineIndex = -1;
    this.lines.forEach((line, i) => {
      if (!this._isXEditableForOperation(line, snappedX, 'add')) return;
      const dist = this._verticalDistanceToLine(line.data, snappedX, chartY);
      if (dist < bestDist) {
        bestDist = dist;
        bestLineIndex = i;
      }
    });
    if (bestLineIndex === -1) return;

    if (!this._isXEditableForOperation(this.lines[bestLineIndex], snappedX, 'add')) {
      return; // Block add
    }

    const theLine = this.lines[bestLineIndex];
    // If a point at that X already exists => do not add
    if (theLine.data.some(pt => pt[0] === snappedX)) {
      return;
    }

    // Simulate adding the new point
    const digits = getPrecisionForMaxY(yMax);
    const snappedY = roundToScale(chartY, digits);
    const newData = [...theLine.data, [snappedX, snappedY]].sort((a, b) => a[0] - b[0]);
    const allLines = this.lines.map((ln, i) =>
      i === bestLineIndex ? { ...ln, data: newData } : ln
    );

    // Interpolate all lines at all integer x
    const xGrid = [];
    for (let x = xMin; x <= xMax; x++) xGrid.push(x);
    const interpolated = allLines.map(ln =>
      this._interpolateLine(ln.data, xMin, xMax).map(([x, y]) => y)
    );

    // Check for crossing at any x
    for (let xi = 0; xi < xGrid.length; xi++) {
      for (let i = 0; i < interpolated.length - 1; i++) {
        if (interpolated[i][xi] < interpolated[i + 1][xi]) {
          // Block the add if any line would cross
          return;
        }
      }
    }

    // If no crossings, proceed with adding the point
    this._pushHistory();
    theLine.data.push([snappedX, snappedY]);
    theLine.data.sort((a, b) => a[0] - b[0]);
    this._refreshChart();
  }

  _handleShiftClick(zrParams) {
    const pos = [zrParams.offsetX, zrParams.offsetY];
    const nearest = this._findNearestPoint(pos, 10, true, 'delete');
    if (!nearest) {
      return;
    }

    const { lineIndex, pointIndex } = nearest;
    const line = this.lines[lineIndex];
    const pointX = line && line.data && line.data[pointIndex] ? line.data[pointIndex][0] : NaN;
    if (!this._isXEditableForOperation(line, pointX, 'delete')) {
      return;
    }
    if (line.data.length <= 2) {
      // Do not allow removing if line would have < 2 points
      return;
    }

    // Simulate removing the point
    const newData = line.data.filter((_, i) => i !== pointIndex);
    const allLines = this.lines.map((ln, i) =>
      i === lineIndex ? { ...ln, data: newData } : ln
    );

    // Interpolate all lines at all integer x
    const xMin = Math.ceil(this.windowStartX);
    const xMax = Math.floor(this.windowEndX);
    const xGrid = [];
    for (let x = xMin; x <= xMax; x++) xGrid.push(x);
    const interpolated = allLines.map(ln =>
      this._interpolateLine(ln.data, xMin, xMax).map(([x, y]) => y)
    );

    // Check for crossing at any x
    for (let xi = 0; xi < xGrid.length; xi++) {
      for (let i = 0; i < interpolated.length - 1; i++) {
        if (interpolated[i][xi] < interpolated[i + 1][xi]) {
          // Block the delete if any line would cross
          return;
        }
      }
    }

    // If no crossings, proceed with removing the point
    this._pushHistory();
    line.data.splice(pointIndex, 1);
    this._refreshChart();
  }

  _ensureEdgePoints(line, xMin, xMax) {
    // Ensure a point at xMin
    if (!line.data.some(pt => Math.abs(pt[0] - xMin) < 1e-6)) {
      // Add a point at xMin with the y-value of the closest point
      const y = line.data.length > 0 ? line.data[0][1] : 0;
      line.data.unshift([xMin, y]);
    }
    // Ensure a point at xMax
    if (!line.data.some(pt => Math.abs(pt[0] - xMax) < 1e-6)) {
      // Add a point at xMax with the y-value of the closest point
      const y = line.data.length > 0 ? line.data[line.data.length - 1][1] : 0;
      line.data.push([xMax, y]);
    }
    // Sort the data by x
    line.data.sort((a, b) => a[0] - b[0]);
  }
}

// class NapkinChartArea extends NapkinChart {
//   constructor(domId, lines = [], editMode = true, baseOption = {}, theme = 'light') {
//     super(domId, lines, editMode, baseOption, theme);

//     // e.g. Let's store the lines in order from "bottom" to "top."
//     // If you want top→bottom, invert or rename as needed.
//     // We'll assume lines[0] is the bottom-most area, lines[1] is above it, etc.
//   }

//   /**
//    * Override the refresh logic to produce stacked-area series
//    */
//   _refreshChart() {
//     // 1) Build ECharts series from this.lines as normal, 
//     //    but add `stack` and `areaStyle`.
//     const seriesList = [];
    
//     for (let i = 0; i < this.lines.length; i++) {
//       const ln = this.lines[i];
//       // For user's drag handles, we can keep the showSymbol = this.editMode
//       seriesList.push({
//         name: ln.name || `Line${i}`,
//         type: 'line',
//         data: ln.data,              // same shape as the parent NapkinChart expects
//         showSymbol: !!this.editMode,
//         symbolSize: this.editMode ? 8 : 0,
//         stack: 'areaStack',         // All lines share the same stack name
//         areaStyle: { opacity: 0.4 },// Fills the area from this line down to the previous stacked line
//         lineStyle: {
//           color: ln.color || '#000'
//         },
//         // We could add other ECharts line properties (smooth, etc.) if desired
//       });
//     }

//     // 2) Build an updated option object
//     const option = Object.assign({}, this.baseOption);
//     option.series = seriesList;

//     // 3) Set the new option on the chart
//     //    Use "notMerge=true" so old series get fully replaced
//     this.chart.setOption(option, true);

//     // 4) Fire the data-changed callback if it's defined
//     if (typeof this.onDataChanged === 'function') {
//       this.onDataChanged();
//     }
//   }

//   /**
//    * We also override or hook the drag logic. 
//    * NapkinChart typically has something like `_onDragPoint(lineIndex, pointIndex, newX, newY)`
//    * We'll clamp so we don't cross the line above or below.
//    */
//   _onDragPoint(lineIndex, pointIndex, newX, newY) {
//     // 1) get references to the line being dragged, 
//     //    the line above it, and the line below it.
//     const currentLine = this.lines[lineIndex];
//     const aboveLine = (lineIndex < this.lines.length - 1) ? this.lines[lineIndex + 1] : null;
//     const belowLine = (lineIndex > 0) ? this.lines[lineIndex - 1] : null;

//     // 2) If we have an above line, clamp `newY` so it never exceeds the above line's Y at the same X
//     if (aboveLine) {
//       const yAbove = this._interpolateLine(aboveLine.data, newX);
//       // Ensure we don't cross: current line must remain <= yAbove
//       // If your "above" line is actually higher numeric y-values, you might invert the logic
//       if (newY > yAbove) {
//         newY = yAbove;
//       }
//     }

//     // 3) If we have a below line, clamp `newY` so it never goes below the below line's Y
//     if (belowLine) {
//       const yBelow = this._interpolateLine(belowLine.data, newX);
//       // Ensure we don't cross: current line must remain >= yBelow
//       if (newY < yBelow) {
//         newY = yBelow;
//       }
//     }

//     // 4) Now set the point's new position
//     currentLine.data[pointIndex][0] = newX; // X
//     currentLine.data[pointIndex][1] = newY; // Y

//     // 5) You might also want to sort the points by X again, recalc anything, etc.
//     currentLine.data.sort((a, b) => a[0] - b[0]);

//     // 6) Now call the original parent's or super's method that re-renders
//     this._refreshChart();
//   }

//   /**
//    * A helper function to find the Y-value on a line at a given X, 
//    * so we can clamp. This is the same logic you might already have
//    * in NapkinChart / your interpolation code.
//    */
//   _interpolateLine(dataPoints, x) {
//     // If x < dataPoints[0].x or x > dataPoints[last].x, clamp or return out-of-range
//     if (x <= dataPoints[0][0]) return dataPoints[0][1];
//     if (x >= dataPoints[dataPoints.length - 1][0]) return dataPoints[dataPoints.length - 1][1];

//     // Otherwise, find the segment
//     for (let i = 1; i < dataPoints.length; i++) {
//       const [x1, y1] = dataPoints[i - 1];
//       const [x2, y2] = dataPoints[i];
//       if (x >= x1 && x <= x2) {
//         // linear interpolation
//         const t = (x - x1) / (x2 - x1);
//         return y1 + t * (y2 - y1);
//       }
//     }
//     // Fallback
//     return dataPoints[dataPoints.length - 1][1];
//   }

//   /**
//    * If needed, override or hook the point-add/remove logic similarly,
//    * so that newly added points or removed points won't break the "no overlap" rule.
//    */
//   _onAddPoint(lineIndex, newX, newY) {
//     // same clamp logic: clamp newY so it doesn't cross lines above/below
//     // then push into line data, sort, refresh, etc.
//   }

//   _onRemovePoint(lineIndex, pointIndex) {
//     // typically no clamp needed, just remove
//     super._onRemovePoint(lineIndex, pointIndex);
//   }
// }
