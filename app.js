(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.ExcelTabB24 = factory();
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const DEFAULT_ROWS = 9;
  const DEFAULT_COLUMNS = 7;
  const STORAGE_KEY = "excel-tab-b24-grid-v1";
  const PENDING_DEAL_STORAGE_KEY = "excel-tab-b24-grid-pending-deal-v1";
  const FORMULA_STORAGE_KEY = "excel-tab-b24-saved-formulas-v1";
  const RECENT_FORMULA_STORAGE_KEY = "excel-tab-b24-recent-formulas-v1";
  const DEAL_STORAGE_KEY_PREFIX = "excel-tab-b24-grid-deal-v1";
  const FUNNEL_STORAGE_KEY_PREFIX = "excel-tab-b24-grid-funnel-v1";
  const SHEET_TYPE_DEAL = "deal";
  const SHEET_TYPE_FUNNEL = "funnel";
  const DISPLAY_VERSION = "v.47";
  const DISPLAY_TITLE = "Excel таблицы в сделке и экспорт данных из CRM";
  const SHARED_STORAGE_ENTITY = "exctabb24";
  const SHARED_STORAGE_PROPERTY = "DATA";
  const MAX_SHEETS_PER_GROUP = 7;
  const SHEET_LIST_STORAGE_KEY_SUFFIX = "sheet-list-v1";
  const DEFAULT_COLUMN_WIDTH = 132;
  const MAX_COLUMN_WIDTH = 420;
  const MIN_COLUMN_WIDTH = 90;
  const DEFAULT_ROW_HEIGHT = 34;
  const MAX_ROW_HEIGHT = 180;
  const MIN_ROW_HEIGHT = 28;
  const MAX_GRID_ROWS = 3000;
  const MAX_GRID_COLUMNS = 3000;
  const HISTORY_LIMIT = 15;
  const DELETE_CONFIRM_CELL_THRESHOLD = 18;
  const MAX_RECENT_FORMULAS = 5;
  const FILL_COLORS = ["", "#fff2cc", "#d9ead3", "#cfe2f3", "#f4cccc", "#eadcf8"];
  const FONT_WEIGHTS = ["", "600", "700", "800"];
  const FONT_STYLES = ["", "italic"];
  const FONT_SIZES = ["", "11pt", "13pt", "15pt", "18pt"];
  const HORIZONTAL_ALIGNMENTS = ["", "left", "center", "right"];
  const VERTICAL_ALIGNMENTS = ["", "top", "middle", "bottom"];
  const SELECTION_EDGE_CLASSES = [
    "selection-edge-top",
    "selection-edge-right",
    "selection-edge-bottom",
    "selection-edge-left",
  ];

  function createGrid(rows = DEFAULT_ROWS, columns = DEFAULT_COLUMNS) {
    return Array.from({ length: rows }, () => Array.from({ length: columns }, () => ""));
  }

  function addRow(grid) {
    const width = Math.max(1, grid[0] ? grid[0].length : DEFAULT_COLUMNS);
    return [...grid, Array.from({ length: width }, () => "")];
  }

  function addColumn(grid) {
    const source = grid.length ? grid : createGrid(DEFAULT_ROWS, DEFAULT_COLUMNS);
    const width = Math.max(1, ...source.map((row) => (Array.isArray(row) ? row.length : 0)));
    return source.map((row) => [
      ...(Array.isArray(row) ? row : []),
      ...Array.from({ length: Math.max(0, width - (Array.isArray(row) ? row.length : 0)) + 1 }, () => ""),
    ]);
  }

  function columnName(index) {
    let value = index + 1;
    let name = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }
    return name;
  }

  function columnIndexFromName(name) {
    return String(name || "")
      .toUpperCase()
      .split("")
      .reduce((index, letter) => index * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  }

  function cellKey(rowIndex, columnIndex) {
    return `${rowIndex}:${columnIndex}`;
  }

  function parseCellKey(key) {
    const [rowIndex, columnIndex] = String(key).split(":").map((part) => Number.parseInt(part, 10));
    return { rowIndex, columnIndex };
  }

  function getRangeCellKeys(start, end) {
    if (!start || !end) return [];

    const minRow = Math.min(start.rowIndex, end.rowIndex);
    const maxRow = Math.max(start.rowIndex, end.rowIndex);
    const minColumn = Math.min(start.columnIndex, end.columnIndex);
    const maxColumn = Math.max(start.columnIndex, end.columnIndex);
    const keys = [];

    for (let rowIndex = minRow; rowIndex <= maxRow; rowIndex += 1) {
      for (let columnIndex = minColumn; columnIndex <= maxColumn; columnIndex += 1) {
        keys.push(cellKey(rowIndex, columnIndex));
      }
    }

    return keys;
  }

  function getFilledCellKeys(grid) {
    const keys = [];
    grid.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        if (String(value || "").trim()) keys.push(cellKey(rowIndex, columnIndex));
      });
    });
    return keys;
  }

  function getSelectedColumns(selectedCells) {
    return Array.from(selectedCells)
      .map((key) => parseCellKey(key).columnIndex)
      .filter((columnIndex) => Number.isInteger(columnIndex) && columnIndex >= 0)
      .filter((columnIndex, index, columns) => columns.indexOf(columnIndex) === index)
      .sort((left, right) => left - right);
  }

  function getSelectedRows(selectedCells) {
    return Array.from(selectedCells)
      .map((key) => parseCellKey(key).rowIndex)
      .filter((rowIndex) => Number.isInteger(rowIndex) && rowIndex >= 0)
      .filter((rowIndex, index, rows) => rows.indexOf(rowIndex) === index)
      .sort((left, right) => left - right);
  }

  function measureColumnWidth(grid, columnIndex) {
    const maxLength = grid.reduce((max, row, rowIndex) => {
      const textLength = String(getCellDisplayValue(grid, rowIndex, columnIndex))
        .split(/\r?\n/)
        .reduce((lineMax, line) => Math.max(lineMax, line.length), 0);
      return Math.max(max, textLength);
    }, columnName(columnIndex).length);
    return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.ceil(maxLength * 8.5) + 42));
  }

  function getAutoFitColumnWidths(grid, selectedCells = null) {
    const columnCount = grid[0] ? grid[0].length : DEFAULT_COLUMNS;
    const columns =
      selectedCells && selectedCells.size
        ? getSelectedColumns(selectedCells)
        : Array.from({ length: columnCount }, (item, index) => index);

    return columns.reduce((widths, columnIndex) => {
      widths[columnIndex] = measureColumnWidth(grid, columnIndex);
      return widths;
    }, {});
  }

  function measureRowHeight(grid, rowIndex, columnWidths = [], selectedCells = null, wrappedCells = new Set()) {
    const row = grid[rowIndex] || [];
    const selectedColumns =
      selectedCells && selectedCells.size
        ? getSelectedColumns(new Set(Array.from(selectedCells).filter((key) => parseCellKey(key).rowIndex === rowIndex)))
        : row.map((cell, columnIndex) => columnIndex);
    const columns = selectedColumns.length ? selectedColumns : row.map((cell, columnIndex) => columnIndex);
    const maxLines = columns.reduce((lineCount, columnIndex) => {
      const value = String(getCellDisplayValue(grid, rowIndex, columnIndex) || "");
      const availableWidth = Math.max(24, (columnWidths[columnIndex] || DEFAULT_COLUMN_WIDTH) - 16);
      const wrappedLines = value.split(/\r?\n/).reduce((count, line) => {
        const estimatedLines = wrappedCells.has(cellKey(rowIndex, columnIndex))
          ? Math.max(1, Math.ceil(line.length * 8.5 / availableWidth))
          : 1;
        return count + estimatedLines;
      }, 0);
      return Math.max(lineCount, wrappedLines || 1);
    }, 1);
    return clampRowHeight(Math.ceil(maxLines * 20 + 14));
  }

  function getAutoFitRowHeights(grid, columnWidths = [], selectedCells = null, wrappedCells = new Set()) {
    const rows =
      selectedCells && selectedCells.size
        ? getSelectedRows(selectedCells)
        : Array.from({ length: grid.length }, (item, index) => index);

    return rows.reduce((heights, rowIndex) => {
      heights[rowIndex] = measureRowHeight(grid, rowIndex, columnWidths, selectedCells, wrappedCells);
      return heights;
    }, {});
  }

  function clampColumnWidth(width) {
    const parsed = Number.parseFloat(width);
    if (!Number.isFinite(parsed)) return DEFAULT_COLUMN_WIDTH;
    return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(parsed)));
  }

  function clampRowHeight(height) {
    const parsed = Number.parseFloat(height);
    if (!Number.isFinite(parsed)) return DEFAULT_ROW_HEIGHT;
    return Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round(parsed)));
  }

  function normalizeColumnWidths(widths = [], columnCount = DEFAULT_COLUMNS, fillMissing = false) {
    const source = Array.isArray(widths) ? widths : [];
    const length = fillMissing ? columnCount : Math.min(columnCount, source.length);
    return Array.from({ length }, (item, columnIndex) =>
      typeof source[columnIndex] === "undefined" ? DEFAULT_COLUMN_WIDTH : clampColumnWidth(source[columnIndex])
    );
  }

  function normalizeRowHeights(heights = [], rowCount = DEFAULT_ROWS, fillMissing = false) {
    const source = Array.isArray(heights) ? heights : [];
    const length = fillMissing ? rowCount : Math.min(rowCount, source.length);
    return Array.from({ length }, (item, rowIndex) =>
      typeof source[rowIndex] === "undefined" ? DEFAULT_ROW_HEIGHT : clampRowHeight(source[rowIndex])
    );
  }

  function isCopyShortcut(event) {
    if (!event || !(event.ctrlKey || event.metaKey)) return false;
    const key = String(event.key || "").toLowerCase();
    return event.code === "KeyC" || key === "c" || key === "с";
  }

  function isPasteShortcut(event) {
    if (!event || !(event.ctrlKey || event.metaKey)) return false;
    const key = String(event.key || "").toLowerCase();
    return event.code === "KeyV" || key === "v" || key === "м";
  }

  function getSortedCellKeys(keys) {
    return Array.from(keys).sort((left, right) => {
      const leftCell = parseCellKey(left);
      const rightCell = parseCellKey(right);
      if (leftCell.rowIndex !== rightCell.rowIndex) return leftCell.rowIndex - rightCell.rowIndex;
      return leftCell.columnIndex - rightCell.columnIndex;
    });
  }

  function parseCellNumber(value) {
    const normalized = String(value || "")
      .replace(/\s+/g, "")
      .replace(",", ".")
      .replace(/[^\d.+-]/g, "");
    if (!normalized || normalized === "-" || normalized === "+" || normalized === ".") return null;

    const parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function formatCalculationResult(value) {
    if (!Number.isFinite(value)) return "";
    const rounded = Math.round((value + Number.EPSILON) * 100000000) / 100000000;
    return Number.isInteger(rounded) ? String(rounded) : String(rounded).replace(".", ",");
  }

  function isFormula(value) {
    return String(value || "").trim().startsWith("=");
  }

  function parseFormulaReference(reference) {
    const match = String(reference || "").match(/^([A-Z]+)([1-9]\d*)$/i);
    if (!match) return null;

    return {
      columnIndex: columnIndexFromName(match[1]),
      rowIndex: Number.parseInt(match[2], 10) - 1,
    };
  }

  function shiftFormulaReferences(formula, rowOffset, columnOffset) {
    return String(formula || "").replace(/\b([A-Z]+)([1-9]\d*)\b/gi, (match, column, row) => {
      const columnIndex = columnIndexFromName(column) + columnOffset;
      const rowIndex = Number.parseInt(row, 10) - 1 + rowOffset;
      if (columnIndex < 0 || rowIndex < 0) return match;
      return `${columnName(columnIndex)}${rowIndex + 1}`;
    });
  }

  function remapFormulaReferences(formula, rowIndexMap = null, columnIndexMap = null) {
    if (!isFormula(formula)) return formula;

    return String(formula || "").replace(/\b([A-Z]+)([1-9]\d*)\b/gi, (match, column, row) => {
      const oldColumnIndex = columnIndexFromName(column);
      const oldRowIndex = Number.parseInt(row, 10) - 1;
      if (columnIndexMap && Object.prototype.hasOwnProperty.call(columnIndexMap, oldColumnIndex) && columnIndexMap[oldColumnIndex] === null) {
        return "#REF!";
      }
      if (rowIndexMap && Object.prototype.hasOwnProperty.call(rowIndexMap, oldRowIndex) && rowIndexMap[oldRowIndex] === null) {
        return "#REF!";
      }
      const nextColumnIndex =
        columnIndexMap && typeof columnIndexMap[oldColumnIndex] === "number" ? columnIndexMap[oldColumnIndex] : oldColumnIndex;
      const nextRowIndex =
        rowIndexMap && typeof rowIndexMap[oldRowIndex] === "number" ? rowIndexMap[oldRowIndex] : oldRowIndex;
      if (nextColumnIndex < 0 || nextRowIndex < 0) return match;
      return `${columnName(nextColumnIndex)}${nextRowIndex + 1}`;
    });
  }

  function appendFormulaReference(formula, reference) {
    const current = String(formula || "");
    const base = current.startsWith("=") ? current : "=";
    const trimmed = base.trimEnd();
    if (trimmed === "=") return `${base}${reference}`;
    if (/[+\-*/^(]\s*$/.test(trimmed)) return `${base}${reference}`;
    return `${base} + ${reference}`;
  }

  function normalizeSavedFormula(formula) {
    const value = formatFormulaInput(formula).trim();
    if (!value) return "";
    return value.startsWith("=") ? value : `=${value}`;
  }

  function sanitizeFormulaInput(value) {
    return String(value || "").replace(/[^\x09\x0A\x0D\x20-\x7E]/g, "");
  }

  function formatFormulaInput(value) {
    return sanitizeFormulaInput(value).replace(/[a-z]/g, (letter) => letter.toUpperCase());
  }

  function normalizeSavedFormulas(formulas) {
    if (!Array.isArray(formulas)) return [];

    const seen = new Set();
    return formulas.reduce((normalizedFormulas, formula) => {
      const normalized = normalizeSavedFormula(formula);
      if (!normalized || seen.has(normalized)) return normalizedFormulas;

      seen.add(normalized);
      normalizedFormulas.push(normalized);
      return normalizedFormulas;
    }, []);
  }

  function addSavedFormula(formulas, formula) {
    const normalized = normalizeSavedFormula(formula);
    if (!normalized) {
      return { error: "Введите формулу", formulas: normalizeSavedFormulas(formulas), formula: "" };
    }

    const savedFormulas = normalizeSavedFormulas(formulas);
    if (!savedFormulas.includes(normalized)) savedFormulas.push(normalized);
    return { error: "", formulas: savedFormulas, formula: normalized };
  }

  function removeSavedFormula(formulas, formula) {
    const normalized = normalizeSavedFormula(formula);
    return normalizeSavedFormulas(formulas).filter((savedFormula) => savedFormula !== normalized);
  }

  function normalizeRecentFormulas(formulas) {
    return normalizeSavedFormulas(formulas).slice(0, MAX_RECENT_FORMULAS);
  }

  function addRecentFormula(formulas, formula) {
    const normalized = normalizeSavedFormula(formula);
    if (!normalized || normalized === "=") return normalizeRecentFormulas(formulas);

    return [normalized, ...normalizeRecentFormulas(formulas).filter((savedFormula) => savedFormula !== normalized)].slice(
      0,
      MAX_RECENT_FORMULAS
    );
  }

  function loadSavedFormulas(storageKey = FORMULA_STORAGE_KEY) {
    try {
      const saved = window.localStorage.getItem(storageKey);
      return normalizeSavedFormulas(saved ? JSON.parse(saved) : []);
    } catch (error) {
      window.localStorage.removeItem(storageKey);
      return [];
    }
  }

  function saveSavedFormulas(formulas, storageKey = FORMULA_STORAGE_KEY) {
    window.localStorage.setItem(storageKey, JSON.stringify(normalizeSavedFormulas(formulas)));
  }

  function loadRecentFormulas(storageKey = RECENT_FORMULA_STORAGE_KEY) {
    try {
      const saved = window.localStorage.getItem(storageKey);
      return normalizeRecentFormulas(saved ? JSON.parse(saved) : []);
    } catch (error) {
      window.localStorage.removeItem(storageKey);
      return [];
    }
  }

  function saveRecentFormulas(formulas, storageKey = RECENT_FORMULA_STORAGE_KEY) {
    window.localStorage.setItem(storageKey, JSON.stringify(normalizeRecentFormulas(formulas)));
  }

  function applyFormulaToGridCell(grid, rowIndex, columnIndex, formula) {
    const normalized = normalizeSavedFormula(formula);
    if (!normalized) return { changed: false, error: "Выберите формулу", grid, formula: "" };
    if (!grid[rowIndex] || typeof grid[rowIndex][columnIndex] === "undefined") {
      return { changed: false, error: "Ячейка не найдена", grid, formula: normalized };
    }

    const nextGrid = grid.map((row) => [...row]);
    const changed = nextGrid[rowIndex][columnIndex] !== normalized;
    nextGrid[rowIndex][columnIndex] = normalized;
    return { changed, error: "", grid: nextGrid, formula: normalized };
  }

  function clearCellSelectionState(state, selectedCells) {
    const keys = selectedCells instanceof Set ? Array.from(selectedCells) : Array.from(selectedCells || []);
    const nextGrid = (state.grid || []).map((row) => [...row]);
    const nextCellFormats = { ...(state.cellFormats || {}) };
    const nextFieldBindings = { ...(state.fieldBindings || {}) };
    const nextWrappedCells = new Set(state.wrappedCells || []);
    let changed = false;

    keys.forEach((key) => {
      const { rowIndex, columnIndex } = parseCellKey(key);
      if (!nextGrid[rowIndex] || typeof nextGrid[rowIndex][columnIndex] === "undefined") return;

      if (nextGrid[rowIndex][columnIndex]) changed = true;
      nextGrid[rowIndex][columnIndex] = "";

      if (nextCellFormats[key]) changed = true;
      delete nextCellFormats[key];

      if (nextFieldBindings[key]) changed = true;
      delete nextFieldBindings[key];

      if (nextWrappedCells.has(key)) changed = true;
      nextWrappedCells.delete(key);
    });

    return {
      cellFormats: nextCellFormats,
      changed,
      fieldBindings: nextFieldBindings,
      grid: nextGrid,
      wrappedCells: nextWrappedCells,
    };
  }

  function getCellRangeBounds(selectedCells) {
    const cells = getSortedCellKeys(selectedCells).map(parseCellKey);
    if (!cells.length) return null;

    return cells.reduce(
      (bounds, cell) => ({
        maxColumnIndex: Math.max(bounds.maxColumnIndex, cell.columnIndex),
        maxRowIndex: Math.max(bounds.maxRowIndex, cell.rowIndex),
        minColumnIndex: Math.min(bounds.minColumnIndex, cell.columnIndex),
        minRowIndex: Math.min(bounds.minRowIndex, cell.rowIndex),
      }),
      {
        maxColumnIndex: cells[0].columnIndex,
        maxRowIndex: cells[0].rowIndex,
        minColumnIndex: cells[0].columnIndex,
        minRowIndex: cells[0].rowIndex,
      }
    );
  }

  function createCellClipboard(state, selectedCells) {
    const bounds = getCellRangeBounds(selectedCells);
    if (!bounds) return null;

    const rowCount = bounds.maxRowIndex - bounds.minRowIndex + 1;
    const columnCount = bounds.maxColumnIndex - bounds.minColumnIndex + 1;
    const values = Array.from({ length: rowCount }, (_, rowOffset) =>
      Array.from({ length: columnCount }, (_, columnOffset) => {
        const rowIndex = bounds.minRowIndex + rowOffset;
        const columnIndex = bounds.minColumnIndex + columnOffset;
        return (state.grid[rowIndex] && state.grid[rowIndex][columnIndex]) || "";
      })
    );
    const cellFormats = {};
    const fieldBindings = {};
    const wrappedCells = [];

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
      for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
        const sourceKey = cellKey(bounds.minRowIndex + rowOffset, bounds.minColumnIndex + columnOffset);
        const relativeKey = cellKey(rowOffset, columnOffset);
        if (state.cellFormats && state.cellFormats[sourceKey]) {
          cellFormats[relativeKey] = JSON.parse(JSON.stringify(state.cellFormats[sourceKey]));
        }
        if (state.fieldBindings && state.fieldBindings[sourceKey]) {
          fieldBindings[relativeKey] = state.fieldBindings[sourceKey];
        }
        if (state.wrappedCells && state.wrappedCells.has(sourceKey)) {
          wrappedCells.push(relativeKey);
        }
      }
    }

    return {
      cellFormats,
      columnCount,
      fieldBindings,
      originColumnIndex: bounds.minColumnIndex,
      originRowIndex: bounds.minRowIndex,
      rowCount,
      values,
      wrappedCells,
    };
  }

  function getClipboardText(clipboard) {
    if (!clipboard || !Array.isArray(clipboard.values)) return "";
    return clipboard.values
      .map((row) => row.map((value) => String(value || "").replace(/\r?\n/g, " ")).join("\t"))
      .join("\n");
  }

  function getClipboardTargetCellKeys(clipboard, targetRowIndex, targetColumnIndex) {
    if (!clipboard || !clipboard.rowCount || !clipboard.columnCount) return [];
    return getRangeCellKeys(
      { rowIndex: targetRowIndex, columnIndex: targetColumnIndex },
      {
        columnIndex: targetColumnIndex + clipboard.columnCount - 1,
        rowIndex: targetRowIndex + clipboard.rowCount - 1,
      }
    );
  }

  function ensureGridSize(grid, minRows, minColumns) {
    const nextGrid = grid.map((row) => [...row]);
    const currentColumns = Math.max(minColumns, ...nextGrid.map((row) => row.length), DEFAULT_COLUMNS);
    while (nextGrid.length < minRows) {
      nextGrid.push(Array.from({ length: currentColumns }, () => ""));
    }
    return nextGrid.map((row) => {
      const nextRow = [...row];
      while (nextRow.length < currentColumns) nextRow.push("");
      return nextRow;
    });
  }

  function pasteCellClipboard(state, clipboard, targetRowIndex, targetColumnIndex) {
    if (!clipboard || !Array.isArray(clipboard.values) || !clipboard.values.length) {
      return { changed: false, state };
    }

    const rowCount = clipboard.values.length;
    const columnCount = Math.max(...clipboard.values.map((row) => row.length));
    const nextGrid = ensureGridSize(state.grid || createGrid(), targetRowIndex + rowCount, targetColumnIndex + columnCount);
    const nextCellFormats = { ...(state.cellFormats || {}) };
    const nextFieldBindings = { ...(state.fieldBindings || {}) };
    const nextWrappedCells = new Set(state.wrappedCells || []);
    const formulaRowOffset = targetRowIndex - (clipboard.originRowIndex || 0);
    const formulaColumnOffset = targetColumnIndex - (clipboard.originColumnIndex || 0);
    const sourceWrappedCells = new Set(clipboard.wrappedCells || []);
    let changed = false;

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
      for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
        const targetKey = cellKey(targetRowIndex + rowOffset, targetColumnIndex + columnOffset);
        const relativeKey = cellKey(rowOffset, columnOffset);
        const sourceValue = (clipboard.values[rowOffset] && clipboard.values[rowOffset][columnOffset]) || "";
        const nextValue = isFormula(sourceValue) ? shiftFormulaReferences(sourceValue, formulaRowOffset, formulaColumnOffset) : sourceValue;

        if (nextGrid[targetRowIndex + rowOffset][targetColumnIndex + columnOffset] !== nextValue) changed = true;
        nextGrid[targetRowIndex + rowOffset][targetColumnIndex + columnOffset] = nextValue;

        if (clipboard.cellFormats && clipboard.cellFormats[relativeKey]) {
          const nextFormat = JSON.parse(JSON.stringify(clipboard.cellFormats[relativeKey]));
          if (JSON.stringify(nextCellFormats[targetKey] || {}) !== JSON.stringify(nextFormat)) changed = true;
          nextCellFormats[targetKey] = nextFormat;
        } else if (nextCellFormats[targetKey]) {
          changed = true;
          delete nextCellFormats[targetKey];
        }

        if (clipboard.fieldBindings && clipboard.fieldBindings[relativeKey]) {
          if (nextFieldBindings[targetKey] !== clipboard.fieldBindings[relativeKey]) changed = true;
          nextFieldBindings[targetKey] = clipboard.fieldBindings[relativeKey];
        } else if (nextFieldBindings[targetKey]) {
          changed = true;
          delete nextFieldBindings[targetKey];
        }

        if (sourceWrappedCells.has(relativeKey)) {
          if (!nextWrappedCells.has(targetKey)) changed = true;
          nextWrappedCells.add(targetKey);
        } else if (nextWrappedCells.has(targetKey)) {
          changed = true;
          nextWrappedCells.delete(targetKey);
        }
      }
    }

    return {
      changed,
      state: {
        cellFormats: nextCellFormats,
        fieldBindings: nextFieldBindings,
        grid: nextGrid,
        wrappedCells: nextWrappedCells,
      },
    };
  }

  function tokenizeFormula(expression) {
    const tokens = [];
    let index = 0;

    while (index < expression.length) {
      const char = expression[index];
      if (/\s/.test(char)) {
        index += 1;
        continue;
      }

      const rest = expression.slice(index);
      const reference = rest.match(/^[A-Z]+[1-9]\d*/i);
      if (reference) {
        tokens.push({ type: "reference", value: reference[0] });
        index += reference[0].length;
        continue;
      }

      const number = rest.match(/^(?:\d+(?:[.,]\d*)?|[.,]\d+)/);
      if (number) {
        tokens.push({ type: "number", value: parseCellNumber(number[0]) });
        index += number[0].length;
        continue;
      }

      if ("+-*/%^()".includes(char)) {
        tokens.push({ type: "operator", value: char });
        index += 1;
        continue;
      }

      return null;
    }

    return tokens;
  }

  function evaluateFormula(grid, rowIndex, columnIndex, visited = new Set()) {
    const key = cellKey(rowIndex, columnIndex);
    if (visited.has(key)) return { error: "circular", value: null };
    const raw = grid[rowIndex] && grid[rowIndex][columnIndex];
    const expression = String(raw || "").trim().slice(1);
    const tokens = tokenizeFormula(expression);
    if (!tokens || !tokens.length) return { error: "invalid", value: null };

    let position = 0;
    const nextVisited = new Set(visited);
    nextVisited.add(key);

    function peek() {
      return tokens[position];
    }

    function consume(value = null) {
      const token = tokens[position];
      if (!token || (value !== null && token.value !== value)) return null;
      position += 1;
      return token;
    }

    function readValue() {
      const token = peek();
      if (!token) return null;

      if (consume("+")) return readValue();
      if (consume("-")) {
        const value = readValue();
        return value === null ? null : -value;
      }

      if (consume("(")) {
        const value = readExpression();
        if (value === null || !consume(")")) return null;
        return value;
      }

      if (token.type === "number") {
        position += 1;
        return token.value;
      }

      if (token.type === "reference") {
        position += 1;
        const reference = parseFormulaReference(token.value);
        if (!reference) return null;
        const referencedValue = grid[reference.rowIndex] && grid[reference.rowIndex][reference.columnIndex];
        let number = null;
        if (isFormula(referencedValue)) {
          const result = evaluateFormula(grid, reference.rowIndex, reference.columnIndex, nextVisited);
          if (result.error) return null;
          number = parseCellNumber(result.value);
        } else {
          number = parseCellNumber(referencedValue);
        }
        return number === null ? 0 : number;
      }

      return null;
    }

    function readPower() {
      let value = readValue();
      if (value === null) return null;

      if (peek() && peek().value === "^") {
        consume("^");
        const right = readPower();
        if (right === null) return null;
        value = value ** right;
      }

      return value;
    }

    function readPercent() {
      let value = readPower();
      if (value === null) return null;

      while (peek() && peek().value === "%") {
        consume("%");
        value /= 100;
      }

      return value;
    }

    function readProduct() {
      let value = readPercent();
      if (value === null) return null;

      while (peek() && (peek().value === "*" || peek().value === "/")) {
        const operator = consume().value;
        const right = readPercent();
        if (right === null) return null;
        if (operator === "/" && right === 0) return null;
        if (operator === "*") value *= right;
        if (operator === "/") value /= right;
      }

      return value;
    }

    function readExpression() {
      let value = readProduct();
      if (value === null) return null;

      while (peek() && (peek().value === "+" || peek().value === "-")) {
        const operator = consume().value;
        const right = readProduct();
        if (right === null) return null;
        value = operator === "+" ? value + right : value - right;
      }

      return value;
    }

    const value = readExpression();
    if (value === null || position !== tokens.length || !Number.isFinite(value)) {
      return { error: "invalid", value: null };
    }

    return { error: "", value: formatCalculationResult(value) };
  }

  function getCellDisplayValue(grid, rowIndex, columnIndex, visited = new Set()) {
    const value = grid[rowIndex] && grid[rowIndex][columnIndex];
    if (!isFormula(value)) return value || "";
    if (String(value || "").trim() === "=") return value;

    const result = evaluateFormula(grid, rowIndex, columnIndex, visited);
    return result.error ? "#ОШИБКА" : result.value;
  }

  function calculateSelectedCells(grid, selectedCells, operation) {
    const values = getSortedCellKeys(selectedCells)
      .map((key) => {
        const { rowIndex, columnIndex } = parseCellKey(key);
        return parseCellNumber(getCellDisplayValue(grid, rowIndex, columnIndex));
      })
      .filter((value) => value !== null);

    if (!values.length) return { error: "Нет числовых значений", value: null };
    if (operation !== "add" && values.length < 2) {
      return { error: "Нужно минимум два числа", value: null };
    }

    let value = values[0];
    if (operation === "add") value = values.reduce((sum, item) => sum + item, 0);
    if (operation === "subtract") value = values.slice(1).reduce((result, item) => result - item, value);
    if (operation === "multiply") value = values.reduce((result, item) => result * item, 1);
    if (operation === "divide") {
      if (values.slice(1).some((item) => item === 0)) return { error: "Деление на ноль", value: null };
      value = values.slice(1).reduce((result, item) => result / item, value);
    }

    if (!Number.isFinite(value)) return { error: "Некорректный результат", value: null };
    return { error: "", value: formatCalculationResult(value) };
  }

  function normalizeCellFormat(format = {}) {
    return {
      fillColor: FILL_COLORS.includes(format.fillColor) ? format.fillColor : "",
      fontWeight: FONT_WEIGHTS.includes(format.fontWeight) ? format.fontWeight : "",
      fontStyle: FONT_STYLES.includes(format.fontStyle) ? format.fontStyle : "",
      fontSize: FONT_SIZES.includes(format.fontSize) ? format.fontSize : "",
      horizontalAlign: HORIZONTAL_ALIGNMENTS.includes(format.horizontalAlign) ? format.horizontalAlign : "",
      verticalAlign: VERTICAL_ALIGNMENTS.includes(format.verticalAlign) ? format.verticalAlign : "",
    };
  }

  function getUsedGridBounds(grid) {
    let lastRow = -1;
    let lastColumn = -1;

    grid.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        if (!String(value || "").trim()) return;
        lastRow = Math.max(lastRow, rowIndex);
        lastColumn = Math.max(lastColumn, columnIndex);
      });
    });

    if (lastRow < 0 || lastColumn < 0) return { lastColumn: 0, lastRow: 0 };
    return { lastColumn, lastRow };
  }

  function isCellKeyInsideBounds(key, rowCount, columnCount) {
    const { rowIndex, columnIndex } = parseCellKey(key);
    return rowIndex >= 0 && rowIndex < rowCount && columnIndex >= 0 && columnIndex < columnCount;
  }

  function countBoundFieldsOnSheet(fieldBindings = {}, rowCount = DEFAULT_ROWS, columnCount = DEFAULT_COLUMNS) {
    return Object.entries(fieldBindings || {}).reduce((count, [key, fieldId]) => {
      if (!fieldId || !isCellKeyInsideBounds(key, rowCount, columnCount)) return count;
      return count + 1;
    }, 0);
  }

  function getSelectionEdgeClassNames(selectedCells, key) {
    const keys = selectedCells instanceof Set ? selectedCells : new Set(selectedCells || []);
    if (!keys.has(key)) return [];

    const { rowIndex, columnIndex } = parseCellKey(key);
    return [
      !keys.has(cellKey(rowIndex - 1, columnIndex)) ? "selection-edge-top" : "",
      !keys.has(cellKey(rowIndex, columnIndex + 1)) ? "selection-edge-right" : "",
      !keys.has(cellKey(rowIndex + 1, columnIndex)) ? "selection-edge-bottom" : "",
      !keys.has(cellKey(rowIndex, columnIndex - 1)) ? "selection-edge-left" : "",
    ].filter(Boolean);
  }

  function filterCellKeyObjectByBounds(source, rowCount, columnCount) {
    return Object.entries(source || {}).reduce((filtered, [key, value]) => {
      if (isCellKeyInsideBounds(key, rowCount, columnCount)) filtered[key] = value;
      return filtered;
    }, {});
  }

  function getNormalizedSheetState(state, minRows = DEFAULT_ROWS, minColumns = DEFAULT_COLUMNS) {
    const sourceGrid = Array.isArray(state.grid) && state.grid.length ? state.grid : createGrid();
    const sourceColumnCount = Math.max(
      minColumns,
      ...sourceGrid.map((row) => (Array.isArray(row) ? row.length : 0))
    );
    const rowCount = Math.min(MAX_GRID_ROWS, Math.max(minRows, sourceGrid.length));
    const columnCount = Math.min(MAX_GRID_COLUMNS, sourceColumnCount);
    const grid = Array.from({ length: rowCount }, (item, rowIndex) => {
      const sourceRow = Array.isArray(sourceGrid[rowIndex]) ? sourceGrid[rowIndex] : [];
      return Array.from({ length: columnCount }, (cell, columnIndex) =>
        typeof sourceRow[columnIndex] === "undefined" ? "" : sourceRow[columnIndex]
      );
    });
    const wrappedCells = new Set(
      Array.from(state.wrappedCells || []).filter((key) => isCellKeyInsideBounds(key, rowCount, columnCount))
    );

    return {
      cellFormats: filterCellKeyObjectByBounds(state.cellFormats, rowCount, columnCount),
      columnWidths: normalizeColumnWidths(state.columnWidths, columnCount, true),
      fieldBindings: filterCellKeyObjectByBounds(state.fieldBindings, rowCount, columnCount),
      grid,
      rowHeights: normalizeRowHeights(state.rowHeights, rowCount, true),
      wrappedCells,
    };
  }

  function moveArrayItem(items = [], fromIndex, toIndex) {
    const source = Array.isArray(items) ? [...items] : [];
    if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || fromIndex >= source.length || toIndex >= source.length) {
      return source;
    }

    const [item] = source.splice(fromIndex, 1);
    source.splice(toIndex, 0, item);
    return source;
  }

  function getMoveIndexMap(length, fromIndex, toIndex) {
    const order = moveArrayItem(Array.from({ length }, (item, index) => index), fromIndex, toIndex);
    return order.reduce((map, oldIndex, newIndex) => {
      map[oldIndex] = newIndex;
      return map;
    }, {});
  }

  function getInsertIndexMap(length, afterIndex, count) {
    return Array.from({ length }, (item, index) => (index > afterIndex ? index + count : index));
  }

  function getDeleteIndexMap(length, deleteIndex) {
    return Array.from({ length }, (item, index) => {
      if (index === deleteIndex) return null;
      return index > deleteIndex ? index - 1 : index;
    });
  }

  function remapCellKeyCollection(source, rowIndexMap, columnIndexMap) {
    const keys = source instanceof Set ? Array.from(source) : Array.from(source || []);
    return keys.reduce((next, key) => {
        const { rowIndex, columnIndex } = parseCellKey(key);
        if (rowIndexMap && Object.prototype.hasOwnProperty.call(rowIndexMap, rowIndex) && rowIndexMap[rowIndex] === null) return next;
        if (
          columnIndexMap &&
          Object.prototype.hasOwnProperty.call(columnIndexMap, columnIndex) &&
          columnIndexMap[columnIndex] === null
        ) {
          return next;
        }
        const nextRowIndex = rowIndexMap && typeof rowIndexMap[rowIndex] === "number" ? rowIndexMap[rowIndex] : rowIndex;
        const nextColumnIndex =
          columnIndexMap && typeof columnIndexMap[columnIndex] === "number" ? columnIndexMap[columnIndex] : columnIndex;
        next.add(cellKey(nextRowIndex, nextColumnIndex));
        return next;
      }, new Set());
  }

  function remapCellKeyObject(source, rowIndexMap, columnIndexMap) {
    return Object.entries(source || {}).reduce((next, [key, value]) => {
      const { rowIndex, columnIndex } = parseCellKey(key);
      if (rowIndexMap && Object.prototype.hasOwnProperty.call(rowIndexMap, rowIndex) && rowIndexMap[rowIndex] === null) return next;
      if (
        columnIndexMap &&
        Object.prototype.hasOwnProperty.call(columnIndexMap, columnIndex) &&
        columnIndexMap[columnIndex] === null
      ) {
        return next;
      }
      const nextRowIndex = rowIndexMap && typeof rowIndexMap[rowIndex] === "number" ? rowIndexMap[rowIndex] : rowIndex;
      const nextColumnIndex =
        columnIndexMap && typeof columnIndexMap[columnIndex] === "number" ? columnIndexMap[columnIndex] : columnIndex;
      next[cellKey(nextRowIndex, nextColumnIndex)] = JSON.parse(JSON.stringify(value));
      return next;
    }, {});
  }

  function remapGridFormulas(grid, rowIndexMap = null, columnIndexMap = null) {
    return grid.map((row) =>
      row.map((value) => (isFormula(value) ? remapFormulaReferences(value, rowIndexMap, columnIndexMap) : value))
    );
  }

  function normalizeInsertCount(count) {
    const parsed = Number.parseInt(count, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
  }

  function insertRowsInSheetState(state, afterRowIndex, count = 1) {
    const source = getNormalizedSheetState(state, DEFAULT_ROWS, DEFAULT_COLUMNS);
    const rowCount = source.grid.length;
    const columnCount = source.grid[0] ? source.grid[0].length : DEFAULT_COLUMNS;
    if (rowCount >= MAX_GRID_ROWS) return source;
    const safeAfterIndex = Math.max(-1, Math.min(rowCount - 1, Number.parseInt(afterRowIndex, 10)));
    const insertCount = Math.min(normalizeInsertCount(count), MAX_GRID_ROWS - rowCount);
    const rowIndexMap = getInsertIndexMap(rowCount, safeAfterIndex, insertCount);
    const blankRows = Array.from({ length: insertCount }, () => Array.from({ length: columnCount }, () => ""));
    const grid = [
      ...source.grid.slice(0, safeAfterIndex + 1),
      ...blankRows,
      ...source.grid.slice(safeAfterIndex + 1),
    ];
    const rowHeights = [
      ...source.rowHeights.slice(0, safeAfterIndex + 1),
      ...Array.from({ length: insertCount }, () => DEFAULT_ROW_HEIGHT),
      ...source.rowHeights.slice(safeAfterIndex + 1),
    ];

    return {
      cellFormats: remapCellKeyObject(source.cellFormats, rowIndexMap, null),
      columnWidths: [...source.columnWidths],
      fieldBindings: remapCellKeyObject(source.fieldBindings, rowIndexMap, null),
      grid: remapGridFormulas(grid, rowIndexMap, null),
      rowHeights,
      wrappedCells: remapCellKeyCollection(source.wrappedCells, rowIndexMap, null),
    };
  }

  function insertColumnsInSheetState(state, afterColumnIndex, count = 1) {
    const source = getNormalizedSheetState(state, DEFAULT_ROWS, DEFAULT_COLUMNS);
    const rowCount = source.grid.length;
    const columnCount = source.grid[0] ? source.grid[0].length : DEFAULT_COLUMNS;
    if (columnCount >= MAX_GRID_COLUMNS) return source;
    const safeAfterIndex = Math.max(-1, Math.min(columnCount - 1, Number.parseInt(afterColumnIndex, 10)));
    const insertCount = Math.min(normalizeInsertCount(count), MAX_GRID_COLUMNS - columnCount);
    const columnIndexMap = getInsertIndexMap(columnCount, safeAfterIndex, insertCount);
    const grid = source.grid.map((row) => [
      ...row.slice(0, safeAfterIndex + 1),
      ...Array.from({ length: insertCount }, () => ""),
      ...row.slice(safeAfterIndex + 1),
    ]);
    const columnWidths = [
      ...source.columnWidths.slice(0, safeAfterIndex + 1),
      ...Array.from({ length: insertCount }, () => DEFAULT_COLUMN_WIDTH),
      ...source.columnWidths.slice(safeAfterIndex + 1),
    ];

    return {
      cellFormats: remapCellKeyObject(source.cellFormats, null, columnIndexMap),
      columnWidths,
      fieldBindings: remapCellKeyObject(source.fieldBindings, null, columnIndexMap),
      grid: remapGridFormulas(grid, null, columnIndexMap),
      rowHeights: [...source.rowHeights],
      wrappedCells: remapCellKeyCollection(source.wrappedCells, null, columnIndexMap),
    };
  }

  function moveRowInSheetState(state, fromRowIndex, toRowIndex) {
    const source = getNormalizedSheetState(state, DEFAULT_ROWS, DEFAULT_COLUMNS);
    const rowCount = source.grid.length;
    if (fromRowIndex === toRowIndex || fromRowIndex < 0 || toRowIndex < 0 || fromRowIndex >= rowCount || toRowIndex >= rowCount) {
      return source;
    }

    const rowIndexMap = getMoveIndexMap(rowCount, fromRowIndex, toRowIndex);
    return {
      cellFormats: remapCellKeyObject(source.cellFormats, rowIndexMap, null),
      columnWidths: [...source.columnWidths],
      fieldBindings: remapCellKeyObject(source.fieldBindings, rowIndexMap, null),
      grid: remapGridFormulas(moveArrayItem(source.grid, fromRowIndex, toRowIndex), rowIndexMap, null),
      rowHeights: moveArrayItem(source.rowHeights, fromRowIndex, toRowIndex),
      wrappedCells: remapCellKeyCollection(source.wrappedCells, rowIndexMap, null),
    };
  }

  function moveColumnInSheetState(state, fromColumnIndex, toColumnIndex) {
    const source = getNormalizedSheetState(state, DEFAULT_ROWS, DEFAULT_COLUMNS);
    const columnCount = source.grid[0] ? source.grid[0].length : DEFAULT_COLUMNS;
    if (
      fromColumnIndex === toColumnIndex ||
      fromColumnIndex < 0 ||
      toColumnIndex < 0 ||
      fromColumnIndex >= columnCount ||
      toColumnIndex >= columnCount
    ) {
      return source;
    }

    const columnIndexMap = getMoveIndexMap(columnCount, fromColumnIndex, toColumnIndex);
    const grid = source.grid.map((row) => moveArrayItem(row, fromColumnIndex, toColumnIndex));
    return {
      cellFormats: remapCellKeyObject(source.cellFormats, null, columnIndexMap),
      columnWidths: moveArrayItem(source.columnWidths, fromColumnIndex, toColumnIndex),
      fieldBindings: remapCellKeyObject(source.fieldBindings, null, columnIndexMap),
      grid: remapGridFormulas(grid, null, columnIndexMap),
      rowHeights: [...source.rowHeights],
      wrappedCells: remapCellKeyCollection(source.wrappedCells, null, columnIndexMap),
    };
  }

  function deleteRowInSheetState(state, rowIndex) {
    const source = getNormalizedSheetState(state, DEFAULT_ROWS, DEFAULT_COLUMNS);
    const rowCount = source.grid.length;
    const columnCount = source.grid[0] ? source.grid[0].length : DEFAULT_COLUMNS;
    if (rowIndex < 0 || rowIndex >= rowCount) return source;

    const rowIndexMap = getDeleteIndexMap(rowCount, rowIndex);
    const grid = source.grid.filter((row, index) => index !== rowIndex);
    const rowHeights = source.rowHeights.filter((height, index) => index !== rowIndex);
    while (grid.length < DEFAULT_ROWS) {
      grid.push(Array.from({ length: columnCount }, () => ""));
      rowHeights.push(DEFAULT_ROW_HEIGHT);
    }

    return {
      cellFormats: remapCellKeyObject(source.cellFormats, rowIndexMap, null),
      columnWidths: [...source.columnWidths],
      fieldBindings: remapCellKeyObject(source.fieldBindings, rowIndexMap, null),
      grid: remapGridFormulas(grid, rowIndexMap, null),
      rowHeights,
      wrappedCells: remapCellKeyCollection(source.wrappedCells, rowIndexMap, null),
    };
  }

  function deleteColumnInSheetState(state, columnIndex) {
    const source = getNormalizedSheetState(state, DEFAULT_ROWS, DEFAULT_COLUMNS);
    const columnCount = source.grid[0] ? source.grid[0].length : DEFAULT_COLUMNS;
    if (columnIndex < 0 || columnIndex >= columnCount) return source;

    const columnIndexMap = getDeleteIndexMap(columnCount, columnIndex);
    const grid = source.grid.map((row) => row.filter((cell, index) => index !== columnIndex));
    const columnWidths = source.columnWidths.filter((width, index) => index !== columnIndex);
    while ((grid[0] ? grid[0].length : 0) < DEFAULT_COLUMNS) {
      grid.forEach((row) => row.push(""));
      columnWidths.push(DEFAULT_COLUMN_WIDTH);
    }

    return {
      cellFormats: remapCellKeyObject(source.cellFormats, null, columnIndexMap),
      columnWidths,
      fieldBindings: remapCellKeyObject(source.fieldBindings, null, columnIndexMap),
      grid: remapGridFormulas(grid, null, columnIndexMap),
      rowHeights: [...source.rowHeights],
      wrappedCells: remapCellKeyCollection(source.wrappedCells, null, columnIndexMap),
    };
  }

  function getTrimmedSheetState(state, minRows = DEFAULT_ROWS, minColumns = DEFAULT_COLUMNS) {
    const sourceGrid = Array.isArray(state.grid) && state.grid.length ? state.grid : createGrid();
    const bounds = getUsedGridBounds(sourceGrid);
    const rowCount = Math.max(minRows, bounds.lastRow + 1);
    const columnCount = Math.max(minColumns, bounds.lastColumn + 1);
    const grid = Array.from({ length: rowCount }, (item, rowIndex) => {
      const sourceRow = Array.isArray(sourceGrid[rowIndex]) ? sourceGrid[rowIndex] : [];
      return Array.from({ length: columnCount }, (cell, columnIndex) =>
        typeof sourceRow[columnIndex] === "undefined" ? "" : sourceRow[columnIndex]
      );
    });
    const wrappedCells = new Set(
      Array.from(state.wrappedCells || []).filter((key) => isCellKeyInsideBounds(key, rowCount, columnCount))
    );

    return {
      cellFormats: filterCellKeyObjectByBounds(state.cellFormats, rowCount, columnCount),
      columnWidths: normalizeColumnWidths(state.columnWidths, columnCount),
      fieldBindings: filterCellKeyObjectByBounds(state.fieldBindings, rowCount, columnCount),
      grid,
      rowHeights: normalizeRowHeights(state.rowHeights, rowCount),
      wrappedCells,
    };
  }

  function getExportGrid(grid) {
    const bounds = getUsedGridBounds(grid);
    return grid
      .slice(0, bounds.lastRow + 1)
      .map((row) => row.slice(0, bounds.lastColumn + 1));
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function getExportCellStyle(format = {}, preserveText = false) {
    const normalizedFormat = normalizeCellFormat(format);
    const styles = [];
    if (preserveText) styles.push('mso-number-format:"\\@"');
    if (normalizedFormat.fillColor) styles.push(`background-color:${normalizedFormat.fillColor}`);
    if (normalizedFormat.fontWeight) styles.push(`font-weight:${normalizedFormat.fontWeight}`);
    if (normalizedFormat.fontStyle) styles.push(`font-style:${normalizedFormat.fontStyle}`);
    if (normalizedFormat.fontSize) styles.push(`font-size:${normalizedFormat.fontSize}`);
    if (normalizedFormat.horizontalAlign) styles.push(`text-align:${normalizedFormat.horizontalAlign}`);
    if (normalizedFormat.verticalAlign) styles.push(`vertical-align:${normalizedFormat.verticalAlign}`);
    return styles.join(";");
  }

  function parseExportNumber(value) {
    const normalized = String(value || "")
      .trim()
      .replace(/\s+/g, "")
      .replace(",", ".");
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(normalized)) return null;

    const number = Number.parseFloat(normalized);
    return Number.isFinite(number) ? normalized : null;
  }

  function getExportCellType(grid, rowIndex, columnIndex) {
    const value = grid[rowIndex] && grid[rowIndex][columnIndex];
    if (isFormula(value)) return "formula";
    return parseExportNumber(getCellDisplayValue(grid, rowIndex, columnIndex)) === null ? "text" : "number";
  }

  function getExportCellContent(grid, rowIndex, columnIndex) {
    const value = grid[rowIndex] && grid[rowIndex][columnIndex];
    if (isFormula(value)) return value;

    const displayValue = getCellDisplayValue(grid, rowIndex, columnIndex);
    return parseExportNumber(displayValue) || displayValue;
  }

  function pxToPt(value) {
    const number = Number.parseFloat(value);
    if (!Number.isFinite(number) || number <= 0) return "";
    return `${Math.round(number * 0.75 * 100) / 100}pt`;
  }

  function getExportColumnWidths(columnCount, columnWidths = []) {
    return Array.from({ length: columnCount }, (_, columnIndex) => {
      const width = columnWidths[columnIndex] || DEFAULT_COLUMN_WIDTH;
      return pxToPt(width) || pxToPt(DEFAULT_COLUMN_WIDTH);
    });
  }

  function buildExcelHtml(grid, cellFormats = {}, options = {}) {
    const rows = getExportGrid(grid);
    const columnCount = rows[0] ? rows[0].length : 0;
    const colgroup = getExportColumnWidths(columnCount, options.columnWidths)
      .map((width) => `<col style='width:${width}'>`)
      .join("");
    const body = rows
      .map((row, rowIndex) => {
        const cells = row
          .map((value, columnIndex) => {
            const exportType = getExportCellType(grid, rowIndex, columnIndex);
            const style = getExportCellStyle(cellFormats[cellKey(rowIndex, columnIndex)], exportType === "text");
            return `<td${style ? ` style='${style}'` : ""}>${escapeHtml(getExportCellContent(grid, rowIndex, columnIndex))}</td>`;
          })
          .join("");
        return `<tr>${cells}</tr>`;
      })
      .join("");

    return `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Type" content="application/vnd.ms-excel; charset=utf-8">
  <!--[if gte mso 9]><xml>
    <x:ExcelWorkbook>
      <x:ExcelWorksheets>
        <x:ExcelWorksheet>
          <x:Name>Excel таблицы в сделке и экспорт данных из CRM</x:Name>
          <x:WorksheetOptions>
            <x:DisplayGridlines/>
          </x:WorksheetOptions>
        </x:ExcelWorksheet>
      </x:ExcelWorksheets>
    </x:ExcelWorkbook>
  </xml><![endif]-->
  <style>
    html, body { background: transparent; margin: 0; padding: 0; }
    table { border-collapse: collapse; }
    tr { height: 25.5pt; }
    td { border: 1.5pt solid #7f7f7f; height: 25.5pt; white-space: pre-wrap; }
  </style>
</head>
<body>
  <table><colgroup>${colgroup}</colgroup>${body}</table>
</body>
</html>`;
  }

  function getExportFileName(dealId) {
    const suffix = dealId ? ` deal ${dealId}` : "";
    return `Excel таблицы в сделке и экспорт данных из CRM${suffix}.xls`;
  }

  function parsePlacementOptions(raw) {
    if (!raw) return {};
    if (typeof raw === "object") return raw;
    try {
      return JSON.parse(raw);
    } catch (error) {
      return {};
    }
  }

  function collectNestedValues(input, keys, values = []) {
    if (!input || typeof input !== "object") return values;

    Object.entries(input).forEach(([key, value]) => {
      if (keys.includes(key)) values.push(value);

      if (typeof value === "string" && /PLACEMENT_OPTIONS|options/i.test(key)) {
        collectNestedValues(parsePlacementOptions(value), keys, values);
        return;
      }

      if (value && typeof value === "object") collectNestedValues(value, keys, values);
    });

    return values;
  }

  function extractDealId(input = {}) {
    const idKeys = [
      "dealId",
      "DEAL_ID",
      "entityId",
      "ENTITY_ID",
      "entityValueId",
      "ENTITY_VALUE_ID",
      "ownerId",
      "OWNER_ID",
      "VALUE_ID",
      "valueId",
      "ID",
      "id",
    ];
    const directUrlDealId = [input.currentUrl, input.referrer, input.url]
      .map((value) => extractDealIdFromText(value))
      .find((value) => Number.isInteger(value));

    if (Number.isInteger(directUrlDealId)) return directUrlDealId;

    const candidates = [...idKeys.map((key) => input[key]), ...collectNestedValues(input, idKeys)];

    for (const candidate of candidates) {
      const value = Number.parseInt(candidate, 10);
      if (Number.isInteger(value) && value > 0) return value;
    }

    const nestedTextDealId = [input.URI, input.PLACEMENT_OPTIONS, input.placement_options]
      .map((value) => extractDealIdFromText(value))
      .find((value) => Number.isInteger(value));

    if (Number.isInteger(nestedTextDealId)) return nestedTextDealId;

    return null;
  }

  function extractDealIdFromText(value) {
    if (typeof value !== "string") return null;
    const match = value.match(/\/crm\/deal\/(?:details|show)\/(\d+)\//i);
    if (!match) return null;
    const dealId = Number.parseInt(match[1], 10);
    return Number.isInteger(dealId) && dealId > 0 ? dealId : null;
  }

  function getPlacementInfo() {
    if (!window.BX24 || !window.BX24.placement || typeof window.BX24.placement.info !== "function") {
      return Promise.resolve({});
    }

    return new Promise((resolve) => {
      let resolved = false;
      const finish = (info) => {
        if (resolved) return;
        resolved = true;
        resolve(info || {});
      };

      try {
        const syncInfo = window.BX24.placement.info((callbackInfo) => finish(callbackInfo));
        if (syncInfo && typeof syncInfo === "object") finish(syncInfo);
      } catch (error) {
        finish({});
      }

      window.setTimeout(() => finish({}), 1200);
    });
  }

  function normalizeFields(fields = {}, deal = {}) {
    return Object.entries(fields)
      .map(([id, meta]) => ({
        id,
        title: meta.formLabel || meta.title || meta.listLabel || id,
        type: meta.type || "",
        value: deal[id],
      }))
      .sort((left, right) => left.title.localeCompare(right.title, "ru"));
  }

  function normalizeIdList(value) {
    if (Array.isArray(value)) return value.flatMap(normalizeIdList);
    if (value && typeof value === "object") {
      if ("ID" in value) return normalizeIdList(value.ID);
      if ("id" in value) return normalizeIdList(value.id);
      if ("VALUE" in value) return normalizeIdList(value.VALUE);
      if ("value" in value) return normalizeIdList(value.value);
    }

    const id = Number.parseInt(value, 10);
    return Number.isInteger(id) && id > 0 ? [id] : [];
  }

  function compactName(parts) {
    return parts.map((part) => String(part || "").trim()).filter(Boolean).join(" ");
  }

  function formatUser(user) {
    if (!user || typeof user !== "object") return "";
    return compactName([user.LAST_NAME, user.NAME, user.SECOND_NAME]) || String(user.ID || "");
  }

  function formatContact(contact) {
    if (!contact || typeof contact !== "object") return "";
    return compactName([contact.LAST_NAME, contact.NAME, contact.SECOND_NAME]) || String(contact.ID || "");
  }

  function formatCompany(company) {
    if (!company || typeof company !== "object") return "";
    return company.TITLE || String(company.ID || "");
  }

  function getDealStageEntityId(categoryId) {
    const id = Number.parseInt(categoryId, 10);
    return Number.isInteger(id) && id > 0 ? `DEAL_STAGE_${id}` : "DEAL_STAGE";
  }

  function findStatusName(statuses, statusId) {
    if (!Array.isArray(statuses)) return "";
    const status = statuses.find((item) => String(item.STATUS_ID) === String(statusId));
    return status ? status.NAME || status.NAME_INIT || String(status.STATUS_ID || "") : "";
  }

  function findCategoryName(categories, categoryId) {
    const normalizedCategoryId = normalizeCategoryId(categoryId);
    if (normalizedCategoryId === null) return "";
    if (normalizedCategoryId === 0) return "\u041e\u0431\u0449\u0430\u044f \u0432\u043e\u0440\u043e\u043d\u043a\u0430";
    if (!Array.isArray(categories)) return "";

    const category = categories.find((item) => {
      const id = item.ID ?? item.id;
      return String(id) === String(normalizedCategoryId);
    });
    if (!category) return "";

    return category.NAME || category.name || category.TITLE || category.title || String(category.ID || category.id || "");
  }

  function normalizeCategoryId(categoryId) {
    const id = Number.parseInt(categoryId, 10);
    return Number.isInteger(id) && id >= 0 ? id : null;
  }

  function applyDisplayValues(fields, displayValues = {}) {
    return fields.map((field) => ({
      ...field,
      value: Object.prototype.hasOwnProperty.call(displayValues, field.id) ? displayValues[field.id] : field.value,
    }));
  }

  function applyFieldBindings(grid, fieldBindings = {}, fields = []) {
    const valuesByFieldId = fields.reduce((values, field) => {
      values[field.id] = formatDealFieldValue(field.value, field.type);
      return values;
    }, {});
    let changed = false;
    const nextGrid = grid.map((row, rowIndex) =>
      row.map((value, columnIndex) => {
        const fieldId = fieldBindings[cellKey(rowIndex, columnIndex)];
        if (!fieldId || !Object.prototype.hasOwnProperty.call(valuesByFieldId, fieldId)) return value;

        const nextValue = valuesByFieldId[fieldId];
        if (nextValue !== value) changed = true;
        return nextValue;
      })
    );

    return { changed, grid: nextGrid };
  }

  function inferFieldBindingsFromGrid(grid, fieldBindings = {}, fields = []) {
    const currentBindings = { ...(fieldBindings || {}) };
    const valueToFieldIds = fields.reduce((map, field) => {
      const value = formatDealFieldValue(field.value, field.type).trim();
      if (!value) return map;
      if (!map[value]) map[value] = [];
      map[value].push(field.id);
      return map;
    }, {});
    const usedFieldIds = new Set(Object.values(currentBindings).filter(Boolean).map(String));
    let changed = false;

    grid.forEach((row, rowIndex) => {
      row.forEach((cellValue, columnIndex) => {
        const key = cellKey(rowIndex, columnIndex);
        if (currentBindings[key]) return;

        const value = String(cellValue || "").trim();
        const fieldIds = valueToFieldIds[value] || [];
        if (fieldIds.length !== 1) return;

        const [fieldId] = fieldIds;
        if (usedFieldIds.has(String(fieldId))) return;
        currentBindings[key] = fieldId;
        usedFieldIds.add(String(fieldId));
        changed = true;
      });
    });

    return { changed, fieldBindings: currentBindings };
  }

  function padDatePart(value) {
    return String(value).padStart(2, "0");
  }

  function parseBitrixDateValue(value) {
    const text = String(value || "").trim();
    const match = text.match(/(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2})(?::\d{2}(?:\.\d+)?)?(?:([+-])(\d{2}):?(\d{2})|Z)?)?/);
    if (!match) return null;

    const [, year, month, day, hour = "00", minute = "00", offsetSign, offsetHour, offsetMinute] = match;
    const offset =
      match[0].endsWith("Z") && !offsetSign
        ? "+0ч"
        : offsetSign && offsetHour
          ? `${offsetSign}${Number.parseInt(offsetHour, 10)}${offsetMinute && offsetMinute !== "00" ? `:${offsetMinute}` : ""}ч`
          : "";

    return {
      date: `${padDatePart(day)}.${padDatePart(month)}.${year}`,
      time: `${padDatePart(hour)}:${padDatePart(minute)}`,
      offset,
    };
  }

  function formatDealDateValue(value, fieldType) {
    const parsed = parseBitrixDateValue(value);
    if (!parsed) return String(value);
    if (String(fieldType || "").toLowerCase() === "datetime") return `${parsed.time}${parsed.offset ? ` (${parsed.offset})` : ""} ${parsed.date}г.`;
    return parsed.date;
  }

  function formatDealFieldValue(value, fieldType = "") {
    const normalizedFieldType = String(fieldType || "").toLowerCase();
    if (value === null || typeof value === "undefined") return "";
    if (Array.isArray(value)) return value.map((item) => formatDealFieldValue(item, normalizedFieldType)).filter(Boolean).join(", ");
    if (typeof value === "object") {
      if ("VALUE" in value) return formatDealFieldValue(value.VALUE, normalizedFieldType);
      if ("value" in value) return formatDealFieldValue(value.value, normalizedFieldType);
      return JSON.stringify(value);
    }
    if (normalizedFieldType === "date" || normalizedFieldType === "datetime") return formatDealDateValue(value, normalizedFieldType);
    return String(value);
  }

  function callMethod(method, params = {}) {
    return new Promise((resolve, reject) => {
      if (!window.BX24 || typeof window.BX24.callMethod !== "function") {
        reject(new Error("Bitrix24 SDK is not available"));
        return;
      }

      window.BX24.callMethod(method, params, (result) => {
        if (result.error()) {
          reject(new Error(result.error_description() || result.error()));
          return;
        }

        resolve(result.data());
      });
    });
  }

  async function callOptionalMethod(method, params = {}, fallback = null) {
    try {
      return await callMethod(method, params);
    } catch (error) {
      return fallback;
    }
  }

  function getSharedItemPropertyValue(item, propertyName = SHARED_STORAGE_PROPERTY) {
    const properties = (item && (item.PROPERTY_VALUES || item.propertyValues || item.properties)) || {};
    const value = properties[propertyName] || properties[propertyName.toLowerCase()];
    if (Array.isArray(value)) return value[0] || "";
    if (value && typeof value === "object") return value.VALUE || value.value || "";
    return typeof value === "undefined" || value === null ? "" : String(value);
  }

  async function ensureSharedStorageEntity() {
    await callMethod("entity.add", {
      ENTITY: SHARED_STORAGE_ENTITY,
      NAME: "Excel Tab B24 shared storage",
      ACCESS: { AU: "W" },
    }).catch(() => null);

    await callMethod("entity.update", {
      ENTITY: SHARED_STORAGE_ENTITY,
      ACCESS: { AU: "W" },
    }).catch(() => null);

    await callMethod("entity.item.property.add", {
      ENTITY: SHARED_STORAGE_ENTITY,
      PROPERTY: SHARED_STORAGE_PROPERTY,
      NAME: "Serialized data",
      TYPE: "S",
    }).catch(() => null);

    await callMethod("entity.item.get", {
      ENTITY: SHARED_STORAGE_ENTITY,
      FILTER: { NAME: "__healthcheck__" },
    });
  }

  async function getSharedStorageItem(storageKey) {
    const result = await callMethod("entity.item.get", {
      ENTITY: SHARED_STORAGE_ENTITY,
      FILTER: { NAME: storageKey },
    });
    const items = Array.isArray(result) ? result : result && Array.isArray(result.items) ? result.items : [];
    return items[0] || null;
  }

  async function loadSharedStorageValue(storageKey) {
    const item = await getSharedStorageItem(storageKey);
    return item ? getSharedItemPropertyValue(item) : "";
  }

  async function saveSharedStorageValue(storageKey, serializedValue) {
    const item = await getSharedStorageItem(storageKey);
    const fields = {
      ENTITY: SHARED_STORAGE_ENTITY,
      NAME: storageKey,
      PROPERTY_VALUES: {
        [SHARED_STORAGE_PROPERTY]: String(serializedValue || ""),
      },
    };

    if (item && (item.ID || item.id)) {
      await callMethod("entity.item.update", {
        ...fields,
        ID: item.ID || item.id,
      });
      return;
    }

    await callMethod("entity.item.add", fields);
  }

  function shouldResolveUserField(field) {
    return (
      field.id === "ASSIGNED_BY_ID" ||
      field.id === "CREATED_BY_ID" ||
      field.id === "MODIFY_BY_ID" ||
      field.id === "MOVED_BY_ID" ||
      field.id === "LAST_ACTIVITY_BY" ||
      field.type === "user"
    );
  }

  async function loadDisplayValues(fields, deal) {
    const displayValues = {};
    const userIds = new Set();

    fields.forEach((field) => {
      if (shouldResolveUserField(field)) {
        normalizeIdList(deal[field.id]).forEach((id) => userIds.add(id));
      }
    });

    await Promise.all(
      Array.from(userIds).map(async (id) => {
        const users = await callOptionalMethod(
          "user.get",
          { FILTER: { ID: id }, select: ["ID", "NAME", "LAST_NAME", "SECOND_NAME"] },
          []
        );
        const user = Array.isArray(users) ? users[0] : users;
        const name = formatUser(user);
        if (!name) return;

        fields.forEach((field) => {
          if (shouldResolveUserField(field) && normalizeIdList(deal[field.id]).includes(id)) {
            displayValues[field.id] = name;
          }
        });
      })
    );

    const contactIds = normalizeIdList(deal.CONTACT_ID || deal.CONTACT_IDS);
    if (contactIds.length) {
      const contacts = await Promise.all(
        contactIds.map((id) => callOptionalMethod("crm.contact.get", { id }, null))
      );
      const contactNames = contacts.map(formatContact).filter(Boolean);
      if (contactNames.length) displayValues.CONTACT_ID = contactNames.join(", ");
    }

    const companyIds = normalizeIdList(deal.COMPANY_ID);
    if (companyIds.length) {
      const companies = await Promise.all(
        companyIds.map((id) => callOptionalMethod("crm.company.get", { id }, null))
      );
      const companyNames = companies.map(formatCompany).filter(Boolean);
      if (companyNames.length) displayValues.COMPANY_ID = companyNames.join(", ");
    }

    if (deal.CATEGORY_ID !== null && typeof deal.CATEGORY_ID !== "undefined") {
      const categories = await callOptionalMethod("crm.dealcategory.list", {}, []);
      const categoryName = findCategoryName(categories, deal.CATEGORY_ID);
      if (categoryName) displayValues.CATEGORY_ID = categoryName;
    }

    if (deal.STAGE_ID) {
      const statuses = await callOptionalMethod(
        "crm.status.list",
        { filter: { ENTITY_ID: getDealStageEntityId(deal.CATEGORY_ID) } },
        []
      );
      const stageName = findStatusName(statuses, deal.STAGE_ID);
      if (stageName) displayValues.STAGE_ID = stageName;
    }

    return displayValues;
  }

  function getGridStorageKey(dealId) {
    const normalizedDealId = Number.parseInt(dealId, 10);
    if (Number.isInteger(normalizedDealId) && normalizedDealId > 0) {
      return `${DEAL_STORAGE_KEY_PREFIX}-${normalizedDealId}`;
    }

    return PENDING_DEAL_STORAGE_KEY;
  }

  function getFunnelStorageKey(categoryId) {
    const normalizedCategoryId = normalizeCategoryId(categoryId);
    if (normalizedCategoryId !== null) {
      return `${FUNNEL_STORAGE_KEY_PREFIX}-${normalizedCategoryId}`;
    }

    return `${FUNNEL_STORAGE_KEY_PREFIX}-local`;
  }

  function getBaseSheetStorageKey(sheetType, dealId, categoryId) {
    return sheetType === SHEET_TYPE_FUNNEL ? getFunnelStorageKey(categoryId) : getGridStorageKey(dealId);
  }

  function getSheetStorageKey(sheetType, dealId, categoryId, sheetIndex = 0) {
    const baseKey = getBaseSheetStorageKey(sheetType, dealId, categoryId);
    const normalizedSheetIndex = Number.parseInt(sheetIndex, 10);
    if (Number.isInteger(normalizedSheetIndex) && normalizedSheetIndex > 0) {
      return `${baseKey}-sheet-${normalizedSheetIndex + 1}`;
    }
    return baseKey;
  }

  function getSheetListStorageKey(sheetType, dealId, categoryId) {
    return `${getBaseSheetStorageKey(sheetType, dealId, categoryId)}-${SHEET_LIST_STORAGE_KEY_SUFFIX}`;
  }

  function getDefaultSheetTitle(sheetIndex) {
    return `Лист ${sheetIndex + 1}`;
  }

  function normalizeSheetList(sheetList = []) {
    const normalized = Array.isArray(sheetList)
      ? sheetList
          .slice(0, MAX_SHEETS_PER_GROUP)
          .map((sheet, index) => ({
            title: String((sheet && sheet.title) || getDefaultSheetTitle(index)).trim() || getDefaultSheetTitle(index),
          }))
      : [];

    return normalized.length ? normalized : [{ title: getDefaultSheetTitle(0) }];
  }

  function isPlainObjectEmpty(value) {
    return !value || typeof value !== "object" || Object.keys(value).length === 0;
  }

  function isSheetStateEmpty(state) {
    if (!state) return true;
    const gridIsEmpty = !Array.isArray(state.grid)
      ? true
      : state.grid.every((row) =>
          !Array.isArray(row) || row.every((value) => String(value === null || typeof value === "undefined" ? "" : value).trim() === "")
        );
    const wrappedCells = state.wrappedCells instanceof Set ? Array.from(state.wrappedCells) : state.wrappedCells || [];

    return (
      gridIsEmpty &&
      isPlainObjectEmpty(state.cellFormats) &&
      isPlainObjectEmpty(state.fieldBindings) &&
      (!Array.isArray(wrappedCells) || wrappedCells.length === 0)
    );
  }

  function pruneEmptySheetList(sheetList = [], sheetStates = []) {
    const sheets = normalizeSheetList(sheetList);
    const kept = sheets
      .map((sheet, index) => ({ index, sheet, state: sheetStates[index] || getEmptySheetState() }))
      .filter((item) => !isSheetStateEmpty(item.state));

    if (!kept.length) {
      return {
        changed: sheets.length > 1,
        oldToNewIndex: { 0: 0 },
        sheets: [{ title: getDefaultSheetTitle(0) }],
        states: [sheetStates[0] || getEmptySheetState()],
      };
    }

    const oldToNewIndex = kept.reduce((indexMap, item, newIndex) => {
      indexMap[item.index] = newIndex;
      return indexMap;
    }, {});

    return {
      changed: kept.length !== sheets.length || kept.some((item, newIndex) => item.index !== newIndex),
      oldToNewIndex,
      sheets: kept.map((item, newIndex) => ({ title: getDefaultSheetTitle(newIndex) })),
      states: kept.map((item) => item.state),
    };
  }

  function parseSheetListData(rawValue) {
    try {
      const parsed = typeof rawValue === "string" ? JSON.parse(rawValue) : rawValue;
      return normalizeSheetList(parsed && Array.isArray(parsed.sheets) ? parsed.sheets : []);
    } catch (error) {
      return normalizeSheetList();
    }
  }

  function serializeSheetList(sheetList) {
    return JSON.stringify({ sheets: normalizeSheetList(sheetList) });
  }

  function loadSheetList(storageKey) {
    try {
      const saved = window.localStorage.getItem(storageKey);
      return parseSheetListData(saved);
    } catch (error) {
      window.localStorage.removeItem(storageKey);
      return normalizeSheetList();
    }
  }

  function saveSheetList(sheetList, storageKey) {
    window.localStorage.setItem(storageKey, serializeSheetList(sheetList));
  }

  function loadGrid(storageKey = STORAGE_KEY) {
    try {
      const saved = window.localStorage.getItem(storageKey);
      const parsed = saved ? JSON.parse(saved) : null;
      if (Array.isArray(parsed) && parsed.length && Array.isArray(parsed[0])) return parsed;
      if (parsed && Array.isArray(parsed.grid) && parsed.grid.length && Array.isArray(parsed.grid[0])) {
        return parsed.grid;
      }
    } catch (error) {
      window.localStorage.removeItem(storageKey);
    }
    return createGrid();
  }

  function saveGrid(grid, storageKey = STORAGE_KEY) {
    window.localStorage.setItem(storageKey, JSON.stringify(grid));
  }

  function parseSheetStateData(rawValue) {
    try {
      const parsed = typeof rawValue === "string" ? JSON.parse(rawValue) : rawValue;
      if (Array.isArray(parsed) && parsed.length && Array.isArray(parsed[0])) {
        return { cellFormats: {}, columnWidths: [], fieldBindings: {}, grid: parsed, rowHeights: [], wrappedCells: new Set() };
      }
      if (parsed && Array.isArray(parsed.grid) && parsed.grid.length && Array.isArray(parsed.grid[0])) {
        const rowCount = parsed.grid.length;
        const columnCount = Math.max(DEFAULT_COLUMNS, ...parsed.grid.map((row) => (Array.isArray(row) ? row.length : 0)));
        return {
          cellFormats: parsed.cellFormats && typeof parsed.cellFormats === "object" ? parsed.cellFormats : {},
          columnWidths: normalizeColumnWidths(parsed.columnWidths, columnCount),
          fieldBindings: parsed.fieldBindings && typeof parsed.fieldBindings === "object" ? parsed.fieldBindings : {},
          grid: parsed.grid,
          rowHeights: normalizeRowHeights(parsed.rowHeights, rowCount),
          wrappedCells: new Set(Array.isArray(parsed.wrappedCells) ? parsed.wrappedCells : []),
        };
      }
    } catch (error) {
      return null;
    }

    return null;
  }

  function getEmptySheetState() {
    return { cellFormats: {}, columnWidths: [], fieldBindings: {}, grid: createGrid(), rowHeights: [], wrappedCells: new Set() };
  }

  function loadSheetState(storageKey = STORAGE_KEY) {
    try {
      const saved = window.localStorage.getItem(storageKey);
      return parseSheetStateData(saved) || getEmptySheetState();
    } catch (error) {
      window.localStorage.removeItem(storageKey);
    }

    return getEmptySheetState();
  }

  function serializeSheetState(state) {
    return JSON.stringify({
      cellFormats: state.cellFormats || {},
      columnWidths: state.columnWidths || [],
      fieldBindings: state.fieldBindings || {},
      grid: state.grid,
      rowHeights: state.rowHeights || [],
      wrappedCells: Array.from(state.wrappedCells || []),
    });
  }

  function saveSheetState(state, storageKey = STORAGE_KEY) {
    window.localStorage.setItem(storageKey, serializeSheetState(state));
  }

  function cloneSheetSnapshot(state) {
    return {
      cellFormats: JSON.parse(JSON.stringify(state.cellFormats || {})),
      columnWidths: Array.isArray(state.columnWidths) ? [...state.columnWidths] : [],
      fieldBindings: JSON.parse(JSON.stringify(state.fieldBindings || {})),
      grid: JSON.parse(JSON.stringify(state.grid || createGrid())),
      rowHeights: Array.isArray(state.rowHeights) ? [...state.rowHeights] : [],
      wrappedCells: Array.from(state.wrappedCells || []),
    };
  }

  function areSheetSnapshotsEqual(left, right) {
    return JSON.stringify(left) === JSON.stringify(right);
  }

  function bootBrowserApp() {
    const table = document.getElementById("sheet");
    const appShell = document.querySelector(".app-shell");
    const topbar = document.querySelector(".topbar");
    const appTitle = document.querySelector(".topbar h1");
    const gridFrame = document.querySelector(".grid-frame");
    const sheetSwitcher = document.querySelector(".sheet-switcher");
    const cellTemplate = document.getElementById("cellTemplate");
    const popover = document.getElementById("fieldPopover");
    const fieldPopoverClose = document.getElementById("fieldPopoverClose");
    const fieldList = document.getElementById("fieldList");
    const fieldSearch = document.getElementById("fieldSearch");
    const formulaSuggestions = document.getElementById("formulaSuggestions");
    const gridStatus = document.getElementById("gridStatus");
    const versionStatus = document.getElementById("versionStatus");
    const dealContext = document.getElementById("dealContext");
    const dealSheetButton = document.getElementById("dealSheetButton");
    const funnelSheetButton = document.getElementById("funnelSheetButton");
    const dealSheetTabs = document.getElementById("dealSheetTabs");
    const funnelSheetTabs = document.getElementById("funnelSheetTabs");
    const addDealSheetButton = document.getElementById("addDealSheetButton");
    const addFunnelSheetButton = document.getElementById("addFunnelSheetButton");
    const addRowButton = document.getElementById("addRowButton");
    const addColumnButton = document.getElementById("addColumnButton");
    const reloadFieldsButton = document.getElementById("reloadFieldsButton");
    const selectFilledButton = document.getElementById("selectFilledButton");
    const selectionActions = document.getElementById("selectionActions");
    const copyCellsButton = document.getElementById("copyCellsButton");
    const pasteCellsButton = document.getElementById("pasteCellsButton");
    const wrapTextButton = document.getElementById("wrapTextButton");
    const clearSelectionButton = document.getElementById("clearSelectionButton");
    const undoButton = document.getElementById("undoButton");
    const redoButton = document.getElementById("redoButton");
    const autoFitButton = document.getElementById("autoFitButton");
    const fillColorSelect = document.getElementById("fillColorSelect");
    const fontSizeSelect = document.getElementById("fontSizeSelect");
    const boldButton = document.getElementById("boldButton");
    const italicButton = document.getElementById("italicButton");
    const alignTopButton = document.getElementById("alignTopButton");
    const alignMiddleButton = document.getElementById("alignMiddleButton");
    const alignBottomButton = document.getElementById("alignBottomButton");
    const alignLeftButton = document.getElementById("alignLeftButton");
    const alignCenterButton = document.getElementById("alignCenterButton");
    const alignRightButton = document.getElementById("alignRightButton");
    const exportExcelButton = document.getElementById("exportExcelButton");
    const formulaLibraryButton = document.getElementById("formulaLibraryButton");
    const formulaModal = document.getElementById("formulaModal");
    const formulaModalClose = document.getElementById("formulaModalClose");
    const formulaList = document.getElementById("formulaList");
    const formulaInput = document.getElementById("formulaInput");
    const saveFormulaButton = document.getElementById("saveFormulaButton");
    const applyFormulaButton = document.getElementById("applyFormulaButton");
    const cancelFormulaButton = document.getElementById("cancelFormulaButton");
    const formulaModalStatus = document.getElementById("formulaModalStatus");
    const deleteConfirmModal = document.getElementById("deleteConfirmModal");
    const deleteConfirmClose = document.getElementById("deleteConfirmClose");
    const confirmDeleteButton = document.getElementById("confirmDeleteButton");
    const cancelDeleteButton = document.getElementById("cancelDeleteButton");
    const helpButton = document.getElementById("helpButton");
    const helpModal = document.getElementById("helpModal");
    const helpModalClose = document.getElementById("helpModalClose");
    let supportWidgetLoaded = false;

    if (appTitle) appTitle.textContent = DISPLAY_TITLE;

    function resizeBitrixFrameToContent() {
      if (!window.BX24 || typeof window.BX24.resizeWindow !== "function" || !appShell) return;
      window.requestAnimationFrame(() => {
        const rect = appShell.getBoundingClientRect();
        const width = Math.ceil(rect.width || window.innerWidth || document.documentElement.clientWidth);
        const height = Math.ceil(rect.height);
        window.BX24.resizeWindow(width, height);
      });
    }

    if (topbar && reloadFieldsButton && exportExcelButton) {
      topbar.insertBefore(exportExcelButton, reloadFieldsButton);
    }
    if (versionStatus) versionStatus.textContent = DISPLAY_VERSION;

    if (gridFrame && sheetSwitcher && gridStatus) {
      const bottomPanel = document.createElement("div");
      const bottomStatusGroup = document.createElement("div");
      bottomPanel.className = "bottom-panel";
      bottomStatusGroup.className = "bottom-status-group";
      bottomPanel.appendChild(sheetSwitcher);
      bottomStatusGroup.appendChild(gridStatus);
      if (versionStatus) bottomStatusGroup.appendChild(versionStatus);
      bottomPanel.appendChild(bottomStatusGroup);
      gridFrame.insertAdjacentElement("afterend", bottomPanel);
    }

    let activeSheetType = SHEET_TYPE_DEAL;
    let dealCategoryId = null;
    let dealCategoryName = "";
    let dealTitle = "";
    let storageKey = getSheetStorageKey(activeSheetType, null, null);
    let sheetState = loadSheetState(storageKey);
    let grid = sheetState.grid;
    let cellFormats = sheetState.cellFormats;
    let wrappedCells = sheetState.wrappedCells;
    let columnWidths = sheetState.columnWidths;
    let rowHeights = sheetState.rowHeights;
    let fieldBindings = sheetState.fieldBindings;
    let currentCell = null;
    let dealFields = [];
    let dealId = null;
    let popoverAnchor = null;
    let selectedCells = new Set();
    let selectionAnchor = null;
    let formulaSourceCell = null;
    let formulaEditCell = null;
    let formulaEditInput = null;
    let formulaReferencePointerHandled = false;
    let savedFormulas = loadSavedFormulas();
    let selectedSavedFormula = savedFormulas[0] || "";
    let recentFormulas = loadRecentFormulas();
    let formulaSuggestionInput = null;
    let gridContextMenu = null;
    let headerDrag = null;
    let dragSelection = null;
    let suppressSelectionClick = false;
    let gridResize = null;
    let undoStack = [];
    let redoStack = [];
    let cellClipboard = null;
    let activeSheetIndex = 0;
    let dealSheets = normalizeSheetList();
    let funnelSheets = normalizeSheetList();
    const sheetHistories = {};
    const sharedStorageWriteQueues = {};
    let suppressCellBlurCommit = false;
    let sharedStorageReady = false;
    let isRestoringHistory = false;

    function getCurrentSheetSnapshot() {
      return cloneSheetSnapshot({ cellFormats, columnWidths, fieldBindings, grid, rowHeights, wrappedCells });
    }

    function applySheetState(nextState) {
      cellFormats = nextState.cellFormats || {};
      columnWidths = nextState.columnWidths || [];
      fieldBindings = nextState.fieldBindings || {};
      grid = nextState.grid || createGrid();
      rowHeights = nextState.rowHeights || [];
      wrappedCells = nextState.wrappedCells instanceof Set ? nextState.wrappedCells : new Set(nextState.wrappedCells || []);
      selectedCells = new Set(Array.from(selectedCells).filter((key) => isCellKeyInsideBounds(key, grid.length, grid[0].length)));
      currentCell = null;
      formulaSourceCell = null;
      formulaEditCell = null;
      formulaEditInput = null;
    }

    function updateHistoryButtons() {
      if (undoButton) undoButton.disabled = undoStack.length === 0;
      if (redoButton) redoButton.disabled = redoStack.length === 0;
    }

    function bindHistoryToStorageKey() {
      if (!sheetHistories[storageKey]) {
        sheetHistories[storageKey] = { redo: [], undo: [] };
      }
      undoStack = sheetHistories[storageKey].undo;
      redoStack = sheetHistories[storageKey].redo;
      updateHistoryButtons();
    }

    function pushUndoSnapshot(snapshot) {
      const last = undoStack[undoStack.length - 1];
      if (last && areSheetSnapshotsEqual(last, snapshot)) return;
      undoStack.push(snapshot);
      if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
    }

    function queueSharedStorageSave(key, serializedValue) {
      if (!sharedStorageReady || !key) return;
      sharedStorageWriteQueues[key] = (sharedStorageWriteQueues[key] || Promise.resolve())
        .catch(() => null)
        .then(() => saveSharedStorageValue(key, serializedValue))
        .catch((error) => {
          if (window.console && typeof window.console.warn === "function") {
            window.console.warn("Failed to save shared sheet data", error);
          }
        });
    }

    function saveSharedSheetState(state, key) {
      queueSharedStorageSave(key, serializeSheetState(state));
    }

    function saveSharedSheetList(sheetList, key) {
      queueSharedStorageSave(key, serializeSheetList(sheetList));
    }

    function persistSheetState(options = {}) {
      if (!isRestoringHistory && !options.skipHistory) {
        const previousSnapshot = cloneSheetSnapshot(loadSheetState(storageKey));
        const currentSnapshot = getCurrentSheetSnapshot();
        if (!areSheetSnapshotsEqual(previousSnapshot, currentSnapshot)) {
          pushUndoSnapshot(previousSnapshot);
          redoStack = [];
        }
      }
      const nextState = { cellFormats, columnWidths, fieldBindings, grid, rowHeights, wrappedCells };
      saveSheetState(nextState, storageKey);
      saveSharedSheetState(nextState, storageKey);
      updateHistoryButtons();
    }

    function restoreSheetSnapshot(snapshot) {
      cellFormats = JSON.parse(JSON.stringify(snapshot.cellFormats || {}));
      columnWidths = Array.isArray(snapshot.columnWidths) ? [...snapshot.columnWidths] : [];
      fieldBindings = JSON.parse(JSON.stringify(snapshot.fieldBindings || {}));
      grid = JSON.parse(JSON.stringify(snapshot.grid || createGrid()));
      rowHeights = Array.isArray(snapshot.rowHeights) ? [...snapshot.rowHeights] : [];
      wrappedCells = new Set(Array.isArray(snapshot.wrappedCells) ? snapshot.wrappedCells : []);
      currentCell = null;
      formulaSourceCell = null;
      formulaEditCell = null;
      formulaEditInput = null;
      selectedCells = new Set(Array.from(selectedCells).filter((key) => isCellKeyInsideBounds(key, grid.length, grid[0].length)));
      isRestoringHistory = true;
      persistSheetState({ skipHistory: true });
      isRestoringHistory = false;
      renderGrid();
      updateHistoryButtons();
    }

    function undoSheetState() {
      const snapshot = undoStack.pop();
      if (!snapshot) return;
      redoStack.push(getCurrentSheetSnapshot());
      if (redoStack.length > HISTORY_LIMIT) redoStack.shift();
      restoreSheetSnapshot(snapshot);
    }

    function redoSheetState() {
      const snapshot = redoStack.pop();
      if (!snapshot) return;
      undoStack.push(getCurrentSheetSnapshot());
      if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
      restoreSheetSnapshot(snapshot);
    }

    function updateDealContext() {
      if (!dealContext) return;

      if (activeSheetType === SHEET_TYPE_FUNNEL) {
        dealContext.textContent = "Общая таблица сделок";
        return;
      }

      if (!dealId) {
        dealContext.textContent = "Таблица сделки не определена. Обновите вкладку после полного открытия карточки.";
        return;
      }

      const title = dealTitle || `ID ${dealId}`;
      dealContext.textContent = `Таблица сделки "${title}"`;
    }

    function applyLoadedSheetState(nextState) {
      grid = nextState.grid;
      cellFormats = nextState.cellFormats;
      wrappedCells = nextState.wrappedCells;
      columnWidths = nextState.columnWidths;
      rowHeights = nextState.rowHeights;
      fieldBindings = nextState.fieldBindings;
      currentCell = null;
      selectedCells = new Set();
      selectionAnchor = null;
      formulaSourceCell = null;
      formulaEditCell = null;
      formulaEditInput = null;
    }

    async function loadSharedSheetStateIntoCurrentSheet(expectedStorageKey) {
      if (!sharedStorageReady || !expectedStorageKey) return false;

      try {
        const serialized = await loadSharedStorageValue(expectedStorageKey);
        if (!serialized) {
          saveSharedSheetState({ cellFormats, columnWidths, fieldBindings, grid, rowHeights, wrappedCells }, expectedStorageKey);
          return false;
        }

        const remoteState = parseSheetStateData(serialized);
        if (!remoteState || storageKey !== expectedStorageKey) return false;
        if (currentCell && currentCell.input && document.activeElement === currentCell.input) return false;

        applyLoadedSheetState(remoteState);
        saveSheetState(remoteState, expectedStorageKey);
        renderGrid();
        return true;
      } catch (error) {
        return false;
      }
    }

    async function loadSheetStateForCleanup(sheetType, sheetIndex, useRemote) {
      const nextStorageKey = getSheetStorageKey(sheetType, dealId, dealCategoryId, sheetIndex);
      if (sheetType === activeSheetType && sheetIndex === activeSheetIndex) return getCurrentSheetSnapshot();

      const localState = loadSheetState(nextStorageKey);
      if (!useRemote || !sharedStorageReady) return localState;

      try {
        const serialized = await loadSharedStorageValue(nextStorageKey);
        const remoteState = parseSheetStateData(serialized);
        if (remoteState) {
          saveSheetState(remoteState, nextStorageKey);
          return remoteState;
        }
      } catch (error) {
        return localState;
      }

      return localState;
    }

    function getActiveSheetList(sheetType) {
      return sheetType === SHEET_TYPE_FUNNEL ? funnelSheets : dealSheets;
    }

    function setActiveSheetList(sheetType, sheets) {
      const normalized = normalizeSheetList(sheets);
      if (sheetType === SHEET_TYPE_FUNNEL) funnelSheets = normalized;
      else dealSheets = normalized;
      return normalized;
    }

    function getSheetGroupContext(sheetType) {
      return {
        addButton: sheetType === SHEET_TYPE_FUNNEL ? addFunnelSheetButton : addDealSheetButton,
        sheets: getActiveSheetList(sheetType),
        tabList: sheetType === SHEET_TYPE_FUNNEL ? funnelSheetTabs : dealSheetTabs,
      };
    }

    function isSheetGroupAvailable(sheetType) {
      return sheetType !== SHEET_TYPE_FUNNEL || (dealId && dealCategoryId !== null);
    }

    function renderSheetGroup(sheetType) {
      const context = getSheetGroupContext(sheetType);
      if (!context.tabList) return;

      const isAvailable = isSheetGroupAvailable(sheetType);
      context.tabList.innerHTML = "";
      context.sheets.forEach((sheet, index) => {
        const button = document.createElement("button");
        button.className = "sheet-switch-button";
        button.type = "button";
        button.textContent = sheet.title || getDefaultSheetTitle(index);
        button.disabled = !isAvailable;
        const isActive = isAvailable && activeSheetType === sheetType && activeSheetIndex === index;
        button.classList.toggle("is-active", isActive);
        button.setAttribute("aria-pressed", String(isActive));
        button.addEventListener("click", () => switchSheet(sheetType, index));
        context.tabList.appendChild(button);
      });

      if (context.addButton) {
        context.addButton.disabled = !isAvailable || context.sheets.length >= MAX_SHEETS_PER_GROUP;
        context.addButton.title =
          context.sheets.length >= MAX_SHEETS_PER_GROUP
            ? `Можно добавить не больше ${MAX_SHEETS_PER_GROUP} листов`
            : "";
      }
    }

    function updateSheetModeControls() {
      renderSheetGroup(SHEET_TYPE_DEAL);
      renderSheetGroup(SHEET_TYPE_FUNNEL);
      updateDealContext();
    }

    async function pruneEmptySheetsInGroup(sheetType, options = {}) {
      if (!isSheetGroupAvailable(sheetType)) return false;

      const sheets = getActiveSheetList(sheetType);
      if (sheets.length <= 1) return false;

      const states = await Promise.all(
        sheets.map((sheet, index) => loadSheetStateForCleanup(sheetType, index, Boolean(options.useRemote)))
      );
      const pruned = pruneEmptySheetList(sheets, states);
      if (!pruned.changed) return false;

      const sheetListKey = getSheetListStorageKey(sheetType, dealId, dealCategoryId);
      const emptyState = getEmptySheetState();
      pruned.states.forEach((state, index) => {
        const nextStorageKey = getSheetStorageKey(sheetType, dealId, dealCategoryId, index);
        saveSheetState(state, nextStorageKey);
        saveSharedSheetState(state, nextStorageKey);
      });
      for (let index = pruned.states.length; index < sheets.length; index += 1) {
        const staleStorageKey = getSheetStorageKey(sheetType, dealId, dealCategoryId, index);
        saveSheetState(emptyState, staleStorageKey);
        saveSharedSheetState(emptyState, staleStorageKey);
      }

      setActiveSheetList(sheetType, pruned.sheets);
      saveSheetList(pruned.sheets, sheetListKey);
      saveSharedSheetList(pruned.sheets, sheetListKey);

      if (activeSheetType === sheetType) {
        activeSheetIndex =
          typeof pruned.oldToNewIndex[activeSheetIndex] === "number"
            ? pruned.oldToNewIndex[activeSheetIndex]
            : Math.min(activeSheetIndex, pruned.sheets.length - 1);
        storageKey = getSheetStorageKey(activeSheetType, dealId, dealCategoryId, activeSheetIndex);
        bindHistoryToStorageKey();
      }

      if (options.render) {
        loadActiveSheetState();
      } else {
        updateSheetModeControls();
      }

      return true;
    }

    function pruneEmptySheetsInGroupBeforeExit(sheetType) {
      if (!isSheetGroupAvailable(sheetType)) return false;

      const sheets = getActiveSheetList(sheetType);
      if (sheets.length <= 1) return false;

      const states = sheets.map((sheet, index) => {
        const nextStorageKey = getSheetStorageKey(sheetType, dealId, dealCategoryId, index);
        return sheetType === activeSheetType && index === activeSheetIndex ? getCurrentSheetSnapshot() : loadSheetState(nextStorageKey);
      });
      const pruned = pruneEmptySheetList(sheets, states);
      if (!pruned.changed) return false;

      const sheetListKey = getSheetListStorageKey(sheetType, dealId, dealCategoryId);
      const emptyState = getEmptySheetState();
      pruned.states.forEach((state, index) => {
        saveSheetState(state, getSheetStorageKey(sheetType, dealId, dealCategoryId, index));
      });
      for (let index = pruned.states.length; index < sheets.length; index += 1) {
        saveSheetState(emptyState, getSheetStorageKey(sheetType, dealId, dealCategoryId, index));
      }

      setActiveSheetList(sheetType, pruned.sheets);
      saveSheetList(pruned.sheets, sheetListKey);
      return true;
    }

    async function pruneEmptySheets(options = {}) {
      const changed = await Promise.all([
        pruneEmptySheetsInGroup(SHEET_TYPE_DEAL, options),
        pruneEmptySheetsInGroup(SHEET_TYPE_FUNNEL, options),
      ]);
      return changed.some(Boolean);
    }

    function pruneEmptySheetsBeforeExit() {
      persistSheetState({ skipHistory: true });
      pruneEmptySheetsInGroupBeforeExit(SHEET_TYPE_DEAL);
      pruneEmptySheetsInGroupBeforeExit(SHEET_TYPE_FUNNEL);
    }

    function loadSheetLists() {
      dealSheets = loadSheetList(getSheetListStorageKey(SHEET_TYPE_DEAL, dealId, dealCategoryId));
      funnelSheets = loadSheetList(getSheetListStorageKey(SHEET_TYPE_FUNNEL, dealId, dealCategoryId));
      if (activeSheetIndex >= getActiveSheetList(activeSheetType).length) activeSheetIndex = 0;
    }

    async function syncSheetListFromSharedStorage(sheetType) {
      if (!sharedStorageReady || !isSheetGroupAvailable(sheetType)) return false;
      const key = getSheetListStorageKey(sheetType, dealId, dealCategoryId);

      try {
        const serialized = await loadSharedStorageValue(key);
        if (!serialized) {
          saveSharedSheetList(getActiveSheetList(sheetType), key);
          return false;
        }

        const nextSheets = parseSheetListData(serialized);
        setActiveSheetList(sheetType, nextSheets);
        saveSheetList(nextSheets, key);
        if (activeSheetType === sheetType && activeSheetIndex >= nextSheets.length) activeSheetIndex = 0;
        updateSheetModeControls();
        return true;
      } catch (error) {
        return false;
      }
    }

    async function syncSheetListsFromSharedStorage() {
      await Promise.all([
        syncSheetListFromSharedStorage(SHEET_TYPE_DEAL),
        syncSheetListFromSharedStorage(SHEET_TYPE_FUNNEL),
      ]);
      await pruneEmptySheets({ render: true, useRemote: true });
    }

    async function initializeSharedStorage() {
      try {
        await ensureSharedStorageEntity();
        sharedStorageReady = true;
        return true;
      } catch (error) {
        sharedStorageReady = false;
        if (window.console && typeof window.console.warn === "function") {
          window.console.warn("Shared Bitrix24 storage is unavailable", error);
        }
        return false;
      }
    }

    function loadActiveSheetState() {
      storageKey = getSheetStorageKey(activeSheetType, dealId, dealCategoryId, activeSheetIndex);
      bindHistoryToStorageKey();
      sheetState = loadSheetState(storageKey);
      applyLoadedSheetState(sheetState);
      renderGrid();
      updateSheetModeControls();
      loadSharedSheetStateIntoCurrentSheet(storageKey);
    }

    function refreshFieldBoundCells(options = {}) {
      if (!dealFields.length) return;

      const inferredBindings = inferFieldBindingsFromGrid(grid, fieldBindings, dealFields);
      if (inferredBindings.changed) fieldBindings = inferredBindings.fieldBindings;
      const updatedBindings = applyFieldBindings(grid, fieldBindings, dealFields);
      if (!updatedBindings.changed && !inferredBindings.changed) return;

      grid = updatedBindings.grid;
      if (options.deferPersist) return;

      persistSheetState();
      renderGrid();
    }

    function compactGridToFilledBounds(options = {}) {
      const nextState = getTrimmedSheetState({ cellFormats, columnWidths, fieldBindings, grid, rowHeights, wrappedCells });
      const changed =
        JSON.stringify(grid) !== JSON.stringify(nextState.grid) ||
        JSON.stringify(cellFormats) !== JSON.stringify(nextState.cellFormats) ||
        JSON.stringify(fieldBindings) !== JSON.stringify(nextState.fieldBindings) ||
        JSON.stringify(rowHeights) !== JSON.stringify(nextState.rowHeights) ||
        JSON.stringify(columnWidths) !== JSON.stringify(nextState.columnWidths) ||
        Array.from(wrappedCells).join("\n") !== Array.from(nextState.wrappedCells).join("\n");

      if (!changed) return false;

      grid = nextState.grid;
      cellFormats = nextState.cellFormats;
      fieldBindings = nextState.fieldBindings;
      columnWidths = nextState.columnWidths;
      rowHeights = nextState.rowHeights;
      wrappedCells = nextState.wrappedCells;
      currentCell = null;
      selectedCells = new Set(Array.from(selectedCells).filter((key) => isCellKeyInsideBounds(key, grid.length, grid[0].length)));
      if (!options.deferPersist) {
        persistSheetState();
        renderGrid();
      }
      return true;
    }

    function switchSheet(sheetType, sheetIndex = 0) {
      if (!isSheetGroupAvailable(sheetType)) return;
      if (sheetType === activeSheetType && sheetIndex === activeSheetIndex) return;
      if (sheetIndex < 0 || sheetIndex >= getActiveSheetList(sheetType).length) return;
      if (sheetType === SHEET_TYPE_FUNNEL && (!dealId || dealCategoryId === null)) return;

      persistSheetState();
      activeSheetType = sheetType;
      activeSheetIndex = sheetIndex;
      closeFieldPopover();
      closeFormulaSuggestions();
      loadActiveSheetState();
      refreshFieldBoundCells();
    }

    function addSheetToGroup(sheetType) {
      if (!isSheetGroupAvailable(sheetType)) return;
      const sheets = getActiveSheetList(sheetType);
      if (sheets.length >= MAX_SHEETS_PER_GROUP) return;

      persistSheetState();
      const nextSheets = setActiveSheetList(sheetType, [...sheets, { title: getDefaultSheetTitle(sheets.length) }]);
      const sheetListKey = getSheetListStorageKey(sheetType, dealId, dealCategoryId);
      saveSheetList(nextSheets, sheetListKey);
      saveSharedSheetList(nextSheets, sheetListKey);
      activeSheetType = sheetType;
      activeSheetIndex = nextSheets.length - 1;
      closeFieldPopover();
      closeFormulaSuggestions();
      loadActiveSheetState();
    }

    function updateSelectionActions() {
      if (selectionActions) selectionActions.hidden = selectedCells.size === 0;
      if (copyCellsButton) copyCellsButton.disabled = selectedCells.size === 0;
      if (pasteCellsButton) pasteCellsButton.disabled = !cellClipboard;
      if (clearSelectionButton) clearSelectionButton.disabled = selectedCells.size === 0;
      if (!fontSizeSelect || !selectedCells.size) return;

      const selectedFormats = Array.from(selectedCells).map((key) => normalizeCellFormat(cellFormats[key] || {}).fontSize || "13pt");
      const firstSize = selectedFormats[0] || "13pt";
      fontSizeSelect.value = selectedFormats.every((size) => size === firstSize) ? firstSize : "";
    }

    function copySelectedCells() {
      if (!selectedCells.size) return;
      flushCurrentCellInput();
      const clipboardKeys = getClipboardSelectionKeys();
      cellClipboard = createCellClipboard({ cellFormats, fieldBindings, grid, wrappedCells }, clipboardKeys);
      if (window.getSelection) window.getSelection().removeAllRanges();
      updateSelectionActions();
      return cellClipboard;
    }

    function getPasteTargetCell() {
      if (currentCell) return { rowIndex: currentCell.rowIndex, columnIndex: currentCell.columnIndex };
      const bounds = getCellRangeBounds(selectedCells);
      return bounds ? { rowIndex: bounds.minRowIndex, columnIndex: bounds.minColumnIndex } : { rowIndex: 0, columnIndex: 0 };
    }

    function pasteCopiedCells() {
      if (!cellClipboard) return;
      const target = getPasteTargetCell();
      const result = pasteCellClipboard({ cellFormats, fieldBindings, grid, wrappedCells }, cellClipboard, target.rowIndex, target.columnIndex);
      if (!result.changed) return;

      grid = result.state.grid;
      cellFormats = result.state.cellFormats;
      fieldBindings = result.state.fieldBindings;
      wrappedCells = result.state.wrappedCells;
      columnWidths = normalizeColumnWidths(columnWidths, grid[0] ? grid[0].length : DEFAULT_COLUMNS, true);
      rowHeights = normalizeRowHeights(rowHeights, grid.length, true);
      selectedCells = new Set(getClipboardTargetCellKeys(cellClipboard, target.rowIndex, target.columnIndex));
      selectionAnchor = target;
      persistSheetState();
      suppressCellBlurCommit = true;
      renderGrid();
      suppressCellBlurCommit = false;
      focusCell(target.rowIndex, target.columnIndex);
    }

    function shouldUseCellClipboardShortcut(event) {
      if ((formulaModal && !formulaModal.hidden) || (deleteConfirmModal && !deleteConfirmModal.hidden)) return false;
      const target = event.target;
      if (!target || typeof target.closest !== "function") return true;
      if (target.closest(".field-popover") || target.closest(".formula-suggestions")) return false;
      return true;
    }

    function getClipboardSelectionKeys() {
      const keys = new Set(selectedCells);
      table.querySelectorAll("td.is-selected[data-cell-key]").forEach((cell) => {
        keys.add(cell.dataset.cellKey);
      });
      return keys;
    }

    function fitCellInputHeight(input) {
      if (!input) return;
      input.style.height = "auto";
      input.style.height = `${clampRowHeight(Math.max(input.scrollHeight, MIN_ROW_HEIGHT))}px`;
    }

    function fitWrappedRowHeight(rowIndex) {
      rowHeights[rowIndex] = measureRowHeight(grid, rowIndex, columnWidths, null, wrappedCells);
      const row = table.querySelector(`tr[data-row="${rowIndex}"]`);
      if (row) row.style.height = `${rowHeights[rowIndex]}px`;
    }

    function setFormulaModalStatus(message) {
      if (formulaModalStatus) formulaModalStatus.textContent = message || "";
    }

    function renderFormulaList() {
      if (!formulaList) return;

      formulaList.innerHTML = "";
      if (!savedFormulas.length) {
        const empty = document.createElement("div");
        empty.className = "formula-empty";
        empty.textContent = "Нет сохранённых формул";
        formulaList.appendChild(empty);
        return;
      }

      savedFormulas.forEach((formula) => {
        const row = document.createElement("div");
        row.className = "formula-row";
        row.classList.toggle("is-selected", formula === selectedSavedFormula);

        const button = document.createElement("button");
        button.className = "formula-option";
        button.type = "button";
        button.textContent = formula;
        button.addEventListener("click", () => {
          selectedSavedFormula = formula;
          if (formulaInput) formulaInput.value = formula;
          setFormulaModalStatus("");
          renderFormulaList();
        });

        const deleteButton = document.createElement("button");
        deleteButton.className = "formula-delete-button";
        deleteButton.type = "button";
        deleteButton.setAttribute("aria-label", `Удалить формулу ${formula}`);
        deleteButton.innerHTML = [
          '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">',
          '<path d="M9 3h6l1 2h4v2H4V5h4l1-2Zm-2 6h10l-.7 11H7.7L7 9Zm3 2v7h2v-7h-2Zm4 0v7h2v-7h-2Z"></path>',
          "</svg>",
        ].join("");
        deleteButton.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          savedFormulas = removeSavedFormula(savedFormulas, formula);
          if (selectedSavedFormula === formula) selectedSavedFormula = savedFormulas[0] || "";
          saveSavedFormulas(savedFormulas);
          if (formulaInput) formulaInput.value = selectedSavedFormula;
          setFormulaModalStatus("Формула удалена");
          renderFormulaList();
        });

        row.appendChild(button);
        row.appendChild(deleteButton);
        formulaList.appendChild(row);
      });
    }

    function getFormulaTargetCell() {
      if (currentCell && selectedCells.has(cellKey(currentCell.rowIndex, currentCell.columnIndex))) {
        return { rowIndex: currentCell.rowIndex, columnIndex: currentCell.columnIndex };
      }

      const firstKey = selectedCells.values().next().value;
      return firstKey ? parseCellKey(firstKey) : null;
    }

    function focusCell(rowIndex, columnIndex) {
      const input = table.querySelector(
        `.cell-input[data-row="${rowIndex}"][data-column="${columnIndex}"]`
      );
      if (!input) return;

      setCurrentCell(input, rowIndex, columnIndex);
      input.focus();
    }

    function openFormulaModal() {
      if (!formulaModal || !selectedCells.size) return;

      closeFieldPopover();
      closeFormulaSuggestions();
      selectedSavedFormula = selectedSavedFormula || savedFormulas[0] || "";
      if (formulaInput) formulaInput.value = selectedSavedFormula;
      setFormulaModalStatus("");
      renderFormulaList();
      formulaModal.hidden = false;
      window.setTimeout(() => {
        const selectedOption = formulaList && formulaList.querySelector(".formula-option.is-selected");
        if (selectedOption) selectedOption.focus();
        else if (formulaInput) formulaInput.focus();
      }, 0);
    }

    function closeFormulaModal() {
      if (!formulaModal) return;
      formulaModal.hidden = true;
      setFormulaModalStatus("");
    }

    function openDeleteConfirmModal() {
      if (!deleteConfirmModal) return;
      closeFieldPopover();
      closeFormulaModal();
      deleteConfirmModal.hidden = false;
      window.setTimeout(() => {
        if (confirmDeleteButton) confirmDeleteButton.focus();
      }, 0);
    }

    function closeDeleteConfirmModal() {
      if (!deleteConfirmModal) return;
      deleteConfirmModal.hidden = true;
    }

    function loadSupportWidget() {
      if (supportWidgetLoaded) return;
      supportWidgetLoaded = true;

      const script = document.createElement("script");
      script.async = true;
      script.src = `https://cdn-ru.bitrix24.ru/b31051/crm/site_button/loader_9_no7zeu.js?${(Date.now() / 60000) | 0}`;
      const anchor = document.getElementsByTagName("script")[0];
      anchor.parentNode.insertBefore(script, anchor);
    }

    function openHelpModal() {
      if (!helpModal) return;
      closeFieldPopover();
      closeFormulaModal();
      closeDeleteConfirmModal();
      helpModal.hidden = false;
      loadSupportWidget();
      window.setTimeout(() => {
        if (helpModalClose) helpModalClose.focus();
      }, 0);
    }

    function closeHelpModal() {
      if (!helpModal) return;
      helpModal.hidden = true;
    }

    function saveFormulaFromModal() {
      if (!formulaInput) return;

      const result = addSavedFormula(savedFormulas, formulaInput.value);
      savedFormulas = result.formulas;
      selectedSavedFormula = result.formula || selectedSavedFormula;
      saveSavedFormulas(savedFormulas);
      if (result.formula) rememberRecentFormula(result.formula);
      renderFormulaList();
      setFormulaModalStatus(result.error || "Формула сохранена");
    }

    function handleFormulaInput() {
      if (!formulaInput) return;

      const currentValue = formulaInput.value;
      const sanitizedValue = formatFormulaInput(currentValue);
      if (currentValue === sanitizedValue) return;

      const selectionStart = formulaInput.selectionStart || sanitizedValue.length;
      const removedBeforeCursor = currentValue.slice(0, selectionStart).length - sanitizeFormulaInput(currentValue.slice(0, selectionStart)).length;
      formulaInput.value = sanitizedValue;
      const nextPosition = Math.max(0, selectionStart - removedBeforeCursor);
      formulaInput.setSelectionRange(nextPosition, nextPosition);
      setFormulaModalStatus("Формулы вводятся английскими буквами, цифрами и символами формул. Буквы автоматически переводятся в верхний регистр.");
    }

    function applySelectedFormulaToCell() {
      const target = getFormulaTargetCell();
      if (!target) {
        setFormulaModalStatus("Выберите ячейку");
        return;
      }

      const result = applyFormulaToGridCell(grid, target.rowIndex, target.columnIndex, selectedSavedFormula);
      if (result.error) {
        setFormulaModalStatus(result.error);
        return;
      }

      grid = result.grid;
      delete fieldBindings[cellKey(target.rowIndex, target.columnIndex)];
      rememberRecentFormula(result.formula);
      formulaSourceCell = { columnIndex: target.columnIndex, rowIndex: target.rowIndex };
      formulaEditCell = null;
      formulaEditInput = null;
      persistSheetState();
      closeFormulaModal();
      renderGrid();
      focusCell(target.rowIndex, target.columnIndex);
    }

    function clearSelectedCells() {
      if (!selectedCells.size) return;

      const result = clearCellSelectionState({ cellFormats, fieldBindings, grid, wrappedCells }, selectedCells);
      grid = result.grid;
      cellFormats = result.cellFormats;
      fieldBindings = result.fieldBindings;
      wrappedCells = result.wrappedCells;
      formulaSourceCell = null;
      formulaEditCell = null;
      formulaEditInput = null;

      if (result.changed) persistSheetState();
      closeDeleteConfirmModal();
      renderGrid();
      focusFirstSelectedCell();
    }

    function requestClearSelectedCells() {
      if (!selectedCells.size) return;
      if (selectedCells.size > DELETE_CONFIRM_CELL_THRESHOLD) {
        openDeleteConfirmModal();
        return;
      }

      clearSelectedCells();
    }

    function paintSelection() {
      table.querySelectorAll("td[data-cell-key]").forEach((cell) => {
        applySelectionCellClasses(cell, cell.dataset.cellKey);
      });
      updateSelectionActions();
    }

    function applySelectionCellClasses(cell, key) {
      const isSelected = selectedCells.has(key);
      cell.classList.toggle("is-selected", isSelected);
      SELECTION_EDGE_CLASSES.forEach((className) => cell.classList.remove(className));
      if (!isSelected) return;

      getSelectionEdgeClassNames(selectedCells, key).forEach((className) => cell.classList.add(className));
    }

    function setSelectedCells(keys) {
      selectedCells = new Set(keys);
      paintSelection();
    }

    function selectCell(rowIndex, columnIndex, options = {}) {
      const key = cellKey(rowIndex, columnIndex);
      if (options.shiftKey && selectionAnchor) {
        setSelectedCells(getRangeCellKeys(selectionAnchor, { rowIndex, columnIndex }));
      } else if (options.toggleKey) {
        const next = new Set(selectedCells);
        const added = !next.has(key);
        if (next.has(key)) next.delete(key);
        else next.add(key);
        selectedCells = next;
        selectionAnchor = { rowIndex, columnIndex };
        paintSelection();
        if (added) applyFormulaToCells([key]);
      } else {
        selectionAnchor = { rowIndex, columnIndex };
        setSelectedCells([key]);
      }
    }

    function beginDragSelection(event, input, rowIndex, columnIndex) {
      if (event.button !== 0 || event.shiftKey || event.ctrlKey || event.metaKey) return false;
      if (event.target && event.target.closest(".field-picker-button")) return false;
      if (handleFormulaReferencePointer(event, rowIndex, columnIndex)) return true;
      if (event.target && event.target.closest(".cell-input")) {
        setCurrentCell(input, rowIndex, columnIndex);
        if (!selectedCells.has(cellKey(rowIndex, columnIndex))) selectCell(rowIndex, columnIndex);
        selectionAnchor = { rowIndex, columnIndex };
        dragSelection = {
          anchor: { rowIndex, columnIndex },
          fromInput: true,
          moved: false,
          pointerId: event.pointerId,
        };
        return false;
      }

      event.preventDefault();
      setCurrentCell(input, rowIndex, columnIndex);
      input.focus({ preventScroll: true });
      selectionAnchor = { rowIndex, columnIndex };
      dragSelection = {
        anchor: { rowIndex, columnIndex },
        fromInput: false,
        moved: false,
        pointerId: event.pointerId,
      };
      if (event.currentTarget && typeof event.currentTarget.setPointerCapture === "function") {
        event.currentTarget.setPointerCapture(event.pointerId);
      }
      if (window.getSelection) window.getSelection().removeAllRanges();
      setSelectedCells([cellKey(rowIndex, columnIndex)]);
      return true;
    }

    function autoFitSheetSize(targetCells = null) {
      const widths = getAutoFitColumnWidths(grid, targetCells);
      Object.entries(widths).forEach(([columnIndex, width]) => {
        columnWidths[Number.parseInt(columnIndex, 10)] = width;
      });
      const heights = getAutoFitRowHeights(grid, columnWidths, targetCells, wrappedCells);
      Object.entries(heights).forEach(([rowIndex, height]) => {
        rowHeights[Number.parseInt(rowIndex, 10)] = height;
      });
    }

    function updateDragSelection(rowIndex, columnIndex) {
      if (!dragSelection) return;
      const target = { rowIndex, columnIndex };
      if (!isSameCell(dragSelection.anchor, target)) {
        dragSelection.moved = true;
        document.body.classList.add("is-drag-selecting-grid");
      }
      setSelectedCells(getRangeCellKeys(dragSelection.anchor, target));
    }

    function updateDragSelectionFromPointer(event) {
      if (!dragSelection || event.pointerId !== dragSelection.pointerId) return;
      const element = document.elementFromPoint(event.clientX, event.clientY);
      const cell = element && typeof element.closest === "function" ? element.closest("td[data-cell-key]") : null;
      if (!cell || !table.contains(cell)) return;
      const target = parseCellKey(cell.dataset.cellKey);
      if (dragSelection.fromInput && isSameCell(dragSelection.anchor, target)) return;
      updateDragSelection(target.rowIndex, target.columnIndex);
      if (window.getSelection) window.getSelection().removeAllRanges();
      event.preventDefault();
    }

    function endDragSelection(event) {
      if (!dragSelection) return;
      if (event && event.pointerId !== dragSelection.pointerId) return;
      suppressSelectionClick = dragSelection.moved;
      document.body.classList.remove("is-drag-selecting-grid");
      if (dragSelection.moved && window.getSelection) window.getSelection().removeAllRanges();
      dragSelection = null;
    }

    function beginColumnResize(event, columnIndex) {
      if (event.button !== 0) return;
      gridResize = {
        index: columnIndex,
        pointerId: event.pointerId,
        startSize: columnWidths[columnIndex] || DEFAULT_COLUMN_WIDTH,
        startX: event.clientX,
        type: "column",
      };
      document.body.classList.add("is-resizing-grid", "is-resizing-column");
      if (typeof event.currentTarget.setPointerCapture === "function") {
        event.currentTarget.setPointerCapture(event.pointerId);
      }
      event.preventDefault();
      event.stopPropagation();
    }

    function beginRowResize(event, rowIndex) {
      if (event.button !== 0) return;
      gridResize = {
        index: rowIndex,
        pointerId: event.pointerId,
        startSize: rowHeights[rowIndex] || DEFAULT_ROW_HEIGHT,
        startY: event.clientY,
        type: "row",
      };
      document.body.classList.add("is-resizing-grid", "is-resizing-row");
      if (typeof event.currentTarget.setPointerCapture === "function") {
        event.currentTarget.setPointerCapture(event.pointerId);
      }
      event.preventDefault();
      event.stopPropagation();
    }

    function updateGridResize(event) {
      if (!gridResize || event.pointerId !== gridResize.pointerId) return;

      if (gridResize.type === "column") {
        const width = clampColumnWidth(gridResize.startSize + event.clientX - gridResize.startX);
        columnWidths[gridResize.index] = width;
        const col = table.querySelector(`col[data-column="${gridResize.index}"]`);
        if (col) col.style.width = `${width}px`;
      } else {
        const height = clampRowHeight(gridResize.startSize + event.clientY - gridResize.startY);
        rowHeights[gridResize.index] = height;
        const row = table.querySelector(`tr[data-row="${gridResize.index}"]`);
        if (row) row.style.height = `${height}px`;
      }

      event.preventDefault();
    }

    function endGridResize(event) {
      if (!gridResize || event.pointerId !== gridResize.pointerId) return;
      persistSheetState();
      document.body.classList.remove("is-resizing-grid", "is-resizing-column", "is-resizing-row");
      gridResize = null;
      resizeBitrixFrameToContent();
      event.preventDefault();
    }

    function clearHeaderDropTarget() {
      table.querySelectorAll(".is-drop-target").forEach((element) => element.classList.remove("is-drop-target"));
    }

    function getHeaderDragTarget(event) {
      const element = document.elementFromPoint(event.clientX, event.clientY);
      if (!element || !headerDrag) return null;
      const selector = headerDrag.type === "column" ? "th[data-column-heading]" : "th[data-row-heading]";
      const header = typeof element.closest === "function" ? element.closest(selector) : null;
      if (!header || !table.contains(header)) return null;
      const value = headerDrag.type === "column" ? header.dataset.columnHeading : header.dataset.rowHeading;
      const index = Number.parseInt(value, 10);
      return Number.isFinite(index) ? { header, index } : null;
    }

    function beginHeaderDrag(event, type, index) {
      if (event.button !== 0 || event.target.closest(".column-resize-handle, .row-resize-handle")) return;
      closeGridContextMenu();
      headerDrag = {
        index,
        moved: false,
        pointerId: event.pointerId,
        targetIndex: index,
        type,
      };
      document.body.classList.add("is-dragging-header");
      if (typeof event.currentTarget.setPointerCapture === "function") {
        event.currentTarget.setPointerCapture(event.pointerId);
      }
      event.preventDefault();
    }

    function updateHeaderDrag(event) {
      if (!headerDrag || event.pointerId !== headerDrag.pointerId) return;
      const target = getHeaderDragTarget(event);
      if (!target) return;

      headerDrag.targetIndex = target.index;
      if (target.index !== headerDrag.index) {
        headerDrag.moved = true;
      }
      clearHeaderDropTarget();
      if (headerDrag.moved) target.header.classList.add("is-drop-target");
      if (window.getSelection) window.getSelection().removeAllRanges();
      event.preventDefault();
    }

    function endHeaderDrag(event) {
      if (!headerDrag || event.pointerId !== headerDrag.pointerId) return;
      const finishedDrag = headerDrag;
      headerDrag = null;
      clearHeaderDropTarget();
      document.body.classList.remove("is-dragging-header");
      if (!finishedDrag.moved || finishedDrag.index === finishedDrag.targetIndex) return;

      suppressSelectionClick = true;
      if (finishedDrag.type === "column") moveColumn(finishedDrag.index, finishedDrag.targetIndex);
      else moveRow(finishedDrag.index, finishedDrag.targetIndex);
      event.preventDefault();
      event.stopPropagation();
    }

    function setCurrentCell(input, rowIndex, columnIndex) {
      currentCell = { input, rowIndex, columnIndex };
      if (isFormula(grid[rowIndex] && grid[rowIndex][columnIndex])) {
        formulaSourceCell = { columnIndex, rowIndex };
        formulaEditCell = { columnIndex, rowIndex };
        formulaEditInput = input;
      }
    }

    function isSameCell(left, right) {
      return (
        left &&
        right &&
        left.rowIndex === right.rowIndex &&
        left.columnIndex === right.columnIndex
      );
    }

    function stopFormulaEditingIfNeeded(rowIndex, columnIndex) {
      if (isSameCell(formulaEditCell, { rowIndex, columnIndex })) return;
      formulaEditCell = null;
      formulaEditInput = null;
    }

    function insertFormulaReference(rowIndex, columnIndex) {
      if (!formulaEditCell || !formulaEditInput) return false;
      if (isSameCell(formulaEditCell, { rowIndex, columnIndex })) return false;

      const sourceValue = grid[formulaEditCell.rowIndex] && grid[formulaEditCell.rowIndex][formulaEditCell.columnIndex];
      if (!isFormula(sourceValue)) return false;

      const reference = `${columnName(columnIndex)}${rowIndex + 1}`;
      const formula = appendFormulaReference(sourceValue, reference);
      grid[formulaEditCell.rowIndex][formulaEditCell.columnIndex] = formula;
      formulaEditInput.value = formula;
      formulaEditInput.focus();
      persistSheetState();
      return true;
    }

    function handleFormulaReferencePointer(event, rowIndex, columnIndex) {
      if (event.type === "click" && formulaReferencePointerHandled) {
        formulaReferencePointerHandled = false;
        event.preventDefault();
        event.stopPropagation();
        return true;
      }

      if (!insertFormulaReference(rowIndex, columnIndex)) return false;

      if (event.type === "mousedown" || event.type === "pointerdown") formulaReferencePointerHandled = true;
      event.preventDefault();
      event.stopPropagation();
      return true;
    }

    function applyFormulaToCells(keys) {
      if (!formulaSourceCell) return;

      const sourceKey = cellKey(formulaSourceCell.rowIndex, formulaSourceCell.columnIndex);
      const sourceFormula = grid[formulaSourceCell.rowIndex] && grid[formulaSourceCell.rowIndex][formulaSourceCell.columnIndex];
      if (!selectedCells.has(sourceKey) || !isFormula(sourceFormula)) return;

      let changed = false;
      keys.forEach((key) => {
        if (key === sourceKey || !selectedCells.has(key)) return;

        const target = parseCellKey(key);
        const rowOffset = target.rowIndex - formulaSourceCell.rowIndex;
        const columnOffset = target.columnIndex - formulaSourceCell.columnIndex;
        const formula = shiftFormulaReferences(sourceFormula, rowOffset, columnOffset);
        if (!grid[target.rowIndex] || grid[target.rowIndex][target.columnIndex] === formula) return;

        grid[target.rowIndex][target.columnIndex] = formula;
        delete fieldBindings[key];
        changed = true;
      });

      if (!changed) return;

      persistSheetState();
      renderGrid();
    }

    function commitCellInput(input, rowIndex, columnIndex) {
      const key = cellKey(rowIndex, columnIndex);
      const previousValue = grid[rowIndex] && grid[rowIndex][columnIndex];
      const nextValue = isFormula(input.value) ? formatFormulaInput(input.value) : input.value;
      if (input.value !== nextValue) {
        const selectionStart = input.selectionStart || nextValue.length;
        const removedBeforeCursor = input.value.slice(0, selectionStart).length - sanitizeFormulaInput(input.value.slice(0, selectionStart)).length;
        input.value = nextValue;
        const nextPosition = Math.max(0, selectionStart - removedBeforeCursor);
        input.setSelectionRange(nextPosition, nextPosition);
      }

      grid[rowIndex][columnIndex] = nextValue;
      if (previousValue !== nextValue) delete fieldBindings[key];
      if (isFormula(nextValue)) {
        formulaSourceCell = { columnIndex, rowIndex };
        formulaEditCell = { columnIndex, rowIndex };
        formulaEditInput = input;
      } else {
        stopFormulaEditingIfNeeded(rowIndex, columnIndex);
        closeFormulaSuggestions();
      }
      persistSheetState();
    }

    function flushCurrentCellInput() {
      if (!currentCell || !currentCell.input) return;
      const { input, rowIndex, columnIndex } = currentCell;
      if (document.body.contains(input)) commitCellInput(input, rowIndex, columnIndex);
    }

    function updateCellFormulaSuggestions(input, rowIndex, columnIndex) {
      if (isFormula(input.value)) {
        renderFormulaSuggestions(input, rowIndex, columnIndex);
      } else {
        closeFormulaSuggestions();
      }
    }

    function finalizeCellInput(input, rowIndex, columnIndex) {
      commitCellInput(input, rowIndex, columnIndex);
      if (isFormula(grid[rowIndex] && grid[rowIndex][columnIndex])) {
        rememberRecentFormula(grid[rowIndex][columnIndex]);
      }
    }

    function updateGridStatus() {
      const rows = grid.length;
      const columns = grid[0] ? grid[0].length : 0;
      const groupLabel = activeSheetType === SHEET_TYPE_FUNNEL ? "Общие" : "Сделка";
      const activeSheet = getActiveSheetList(activeSheetType)[activeSheetIndex] || { title: getDefaultSheetTitle(activeSheetIndex) };
      gridStatus.textContent = `${groupLabel}, ${activeSheet.title}: ${rows} строк, ${columns} столбцов`;
    }

    function renderGrid() {
      table.innerHTML = "";
      const columnCount = grid[0] ? grid[0].length : DEFAULT_COLUMNS;
      const colgroup = document.createElement("colgroup");
      const headingColumn = document.createElement("col");
      headingColumn.style.width = "46px";
      colgroup.appendChild(headingColumn);

      for (let column = 0; column < columnCount; column += 1) {
        const col = document.createElement("col");
        col.dataset.column = String(column);
        col.style.width = `${clampColumnWidth(columnWidths[column])}px`;
        colgroup.appendChild(col);
      }

      table.appendChild(colgroup);
      const thead = document.createElement("thead");
      const headRow = document.createElement("tr");
      const corner = document.createElement("th");
      corner.className = "corner-heading";
      corner.title = "Выделить всю таблицу";
      corner.addEventListener("click", () => {
        const allCells = getRangeCellKeys(
          { rowIndex: 0, columnIndex: 0 },
          { rowIndex: grid.length - 1, columnIndex: columnCount - 1 }
        );
        selectionAnchor = { rowIndex: 0, columnIndex: 0 };
        setSelectedCells(allCells);
      });
      headRow.appendChild(corner);

      for (let column = 0; column < columnCount; column += 1) {
        const th = document.createElement("th");
        const label = document.createElement("span");
        const resizeHandle = document.createElement("span");
        th.dataset.columnHeading = String(column);
        label.textContent = columnName(column);
        resizeHandle.className = "column-resize-handle";
        resizeHandle.setAttribute("aria-hidden", "true");
        th.appendChild(label);
        th.appendChild(resizeHandle);
        th.title = "Выделить столбец";
        resizeHandle.addEventListener("pointerdown", (event) => beginColumnResize(event, column));
        th.addEventListener("pointerdown", (event) => beginHeaderDrag(event, "column", column));
        th.addEventListener("contextmenu", (event) => showGridContextMenu(event, "column", column));
        th.addEventListener("click", (event) => {
          if (suppressSelectionClick) {
            suppressSelectionClick = false;
            event.preventDefault();
            event.stopPropagation();
            return;
          }
          const columnCells = grid.map((row, rowIndex) => cellKey(rowIndex, column));
          if (event.ctrlKey || event.metaKey) {
            const next = new Set(selectedCells);
            columnCells.forEach((key) => next.add(key));
            selectedCells = next;
            paintSelection();
            applyFormulaToCells(columnCells);
          } else {
            selectionAnchor = { rowIndex: 0, columnIndex: column };
            setSelectedCells(columnCells);
          }
        });
        headRow.appendChild(th);
      }

      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      grid.forEach((row, rowIndex) => {
        const tr = document.createElement("tr");
        tr.dataset.row = String(rowIndex);
        tr.style.height = `${clampRowHeight(rowHeights[rowIndex])}px`;
        const heading = document.createElement("th");
        heading.className = "row-heading";
        heading.dataset.rowHeading = String(rowIndex);
        const headingLabel = document.createElement("span");
        const resizeHandle = document.createElement("span");
        headingLabel.textContent = String(rowIndex + 1);
        resizeHandle.className = "row-resize-handle";
        resizeHandle.setAttribute("aria-hidden", "true");
        heading.appendChild(headingLabel);
        heading.appendChild(resizeHandle);
        heading.title = "Выделить строку";
        resizeHandle.addEventListener("pointerdown", (event) => beginRowResize(event, rowIndex));
        heading.addEventListener("pointerdown", (event) => beginHeaderDrag(event, "row", rowIndex));
        heading.addEventListener("contextmenu", (event) => showGridContextMenu(event, "row", rowIndex));
        heading.addEventListener("click", (event) => {
          if (suppressSelectionClick) {
            suppressSelectionClick = false;
            event.preventDefault();
            event.stopPropagation();
            return;
          }
          const rowCells = row.map((value, columnIndex) => cellKey(rowIndex, columnIndex));
          if (event.ctrlKey || event.metaKey) {
            const next = new Set(selectedCells);
            rowCells.forEach((key) => next.add(key));
            selectedCells = next;
            paintSelection();
            applyFormulaToCells(rowCells);
          } else {
            selectionAnchor = { rowIndex, columnIndex: 0 };
            setSelectedCells(rowCells);
          }
        });
        tr.appendChild(heading);

        row.forEach((value, columnIndex) => {
          const td = document.createElement("td");
          const key = cellKey(rowIndex, columnIndex);
          td.dataset.cellKey = key;
          applySelectionCellClasses(td, key);
          td.classList.toggle("is-wrapped", wrappedCells.has(key));
          const format = normalizeCellFormat(cellFormats[key] || {});
          if (format.verticalAlign) td.classList.add(`align-vertical-${format.verticalAlign}`);
          if (format.fillColor) td.style.backgroundColor = format.fillColor;
          const fragment = cellTemplate.content.cloneNode(true);
          const input = fragment.querySelector(".cell-input");
          const picker = fragment.querySelector(".field-picker-button");
          input.value = isFormula(value) ? getCellDisplayValue(grid, rowIndex, columnIndex) : value;
          if (format.fillColor) input.style.backgroundColor = format.fillColor;
          if (format.fontWeight) input.style.fontWeight = format.fontWeight;
          if (format.fontStyle) input.style.fontStyle = format.fontStyle;
          if (format.fontSize) input.style.fontSize = format.fontSize;
          if (format.horizontalAlign) input.style.textAlign = format.horizontalAlign;
          if (format.verticalAlign) {
            td.style.verticalAlign = format.verticalAlign;
          }
          input.dataset.row = String(rowIndex);
          input.dataset.column = String(columnIndex);
          picker.dataset.row = String(rowIndex);
          picker.dataset.column = String(columnIndex);
          td.addEventListener("click", (event) => {
            if (suppressSelectionClick) {
              suppressSelectionClick = false;
              event.preventDefault();
              event.stopPropagation();
              return;
            }
            if (handleFormulaReferencePointer(event, rowIndex, columnIndex)) return;
            setCurrentCell(input, rowIndex, columnIndex);
            selectCell(rowIndex, columnIndex, {
              shiftKey: event.shiftKey,
              toggleKey: event.ctrlKey || event.metaKey,
            });
          });
          td.addEventListener("pointerdown", (event) => beginDragSelection(event, input, rowIndex, columnIndex));
          td.addEventListener("pointerenter", () => updateDragSelection(rowIndex, columnIndex));
          input.addEventListener("click", (event) => {
            if (suppressSelectionClick) {
              suppressSelectionClick = false;
              event.preventDefault();
            }
            event.stopPropagation();
          });
          input.addEventListener("focus", () => {
            setCurrentCell(input, rowIndex, columnIndex);
            if (isFormula(grid[rowIndex][columnIndex])) input.value = grid[rowIndex][columnIndex];
            if (!selectedCells.has(key)) selectCell(rowIndex, columnIndex);
            updateCellFormulaSuggestions(input, rowIndex, columnIndex);
          });
          input.addEventListener("blur", () => {
            if (suppressCellBlurCommit || !document.body.contains(input)) {
              formulaEditCell = null;
              formulaEditInput = null;
              closeFormulaSuggestions();
              return;
            }
            finalizeCellInput(input, rowIndex, columnIndex);
            formulaEditCell = null;
            formulaEditInput = null;
            closeFormulaSuggestions();
            if (isFormula(grid[rowIndex][columnIndex])) renderGrid();
          });
          input.addEventListener("input", () => {
            commitCellInput(input, rowIndex, columnIndex);
            if (td.classList.contains("is-wrapped")) {
              fitWrappedRowHeight(rowIndex);
              persistSheetState();
            } else if (format.verticalAlign) {
              fitCellInputHeight(input);
            }
            updateCellFormulaSuggestions(input, rowIndex, columnIndex);
          });
          input.addEventListener("keydown", (event) => {
            if (event.key !== "Enter" || event.shiftKey) return;
            event.preventDefault();
            finalizeCellInput(input, rowIndex, columnIndex);
            closeFormulaSuggestions();
            input.blur();
          });
          picker.addEventListener("click", (event) => {
            if (!popover.hidden && popoverAnchor === event.currentTarget) {
              closeFieldPopover();
              return;
            }

            setCurrentCell(input, rowIndex, columnIndex);
            openFieldPopover(event.currentTarget);
          });
          td.appendChild(fragment);
          tr.appendChild(td);
        });

        tbody.appendChild(tr);
      });

      table.appendChild(tbody);
      table.querySelectorAll("td[class*='align-vertical-']:not(.is-wrapped) .cell-input").forEach(fitCellInputHeight);
      updateGridStatus();
      updateSelectionActions();
      resizeBitrixFrameToContent();
    }

    function renderFieldList(filter = "") {
      const query = filter.trim().toLowerCase();
      const visibleFields = dealFields.filter((field) => {
        const haystack = `${field.title} ${field.id}`.toLowerCase();
        return haystack.includes(query);
      });

      fieldList.innerHTML = "";
      if (!visibleFields.length) {
        const empty = document.createElement("div");
        empty.className = "field-option";
        empty.textContent = "Поля не найдены";
        fieldList.appendChild(empty);
        return;
      }

      visibleFields.forEach((field) => {
        const button = document.createElement("button");
        button.className = "field-option";
        button.type = "button";
        button.innerHTML = `<span></span><small></small>`;
        button.querySelector("span").textContent = field.title;
        button.querySelector("small").textContent = field.id;
        button.addEventListener("click", () => {
          if (!currentCell || !currentCell.input) return;
          const formatted = formatDealFieldValue(field.value, field.type);
          currentCell.input.value = formatted;
          grid[currentCell.rowIndex][currentCell.columnIndex] = formatted;
          fieldBindings[cellKey(currentCell.rowIndex, currentCell.columnIndex)] = field.id;
          persistSheetState();
          closeFieldPopover();
          currentCell.input.focus();
        });
        fieldList.appendChild(button);
      });
    }

    function openFieldPopover(anchor) {
      popoverAnchor = anchor;
      renderFieldList(fieldSearch.value);
      const rect = anchor.getBoundingClientRect();
      popover.hidden = false;
      popover.style.left = `${Math.min(rect.left, window.innerWidth - 300)}px`;
      popover.style.top = `${Math.min(rect.bottom + 4, window.innerHeight - 310)}px`;
      fieldSearch.focus();
    }

    function closeFieldPopover() {
      popover.hidden = true;
      popoverAnchor = null;
      fieldSearch.value = "";
    }

    function closeFormulaSuggestions() {
      if (!formulaSuggestions) return;
      formulaSuggestions.hidden = true;
      formulaSuggestions.innerHTML = "";
      formulaSuggestionInput = null;
    }

    function rememberRecentFormula(formula) {
      const nextRecentFormulas = addRecentFormula(recentFormulas, formula);
      if (nextRecentFormulas.join("\n") === recentFormulas.join("\n")) return;

      recentFormulas = nextRecentFormulas;
      saveRecentFormulas(recentFormulas);
    }

    function renderFormulaSuggestions(input, rowIndex, columnIndex) {
      if (!formulaSuggestions || !input || !isFormula(input.value)) {
        closeFormulaSuggestions();
        return;
      }

      const suggestionFormulas = normalizeSavedFormulas([...recentFormulas, ...savedFormulas]).slice(0, MAX_RECENT_FORMULAS);
      if (!suggestionFormulas.length) {
        closeFormulaSuggestions();
        return;
      }

      formulaSuggestionInput = input;
      formulaSuggestions.innerHTML = "";
      suggestionFormulas.forEach((formula) => {
        const button = document.createElement("button");
        button.className = "formula-suggestion-option";
        button.type = "button";
        button.textContent = formula;
        button.addEventListener("mousedown", (event) => {
          event.preventDefault();
        });
        button.addEventListener("click", () => {
          input.value = formula;
          commitCellInput(input, rowIndex, columnIndex);
          rememberRecentFormula(formula);
          closeFormulaSuggestions();
          input.focus();
        });
        formulaSuggestions.appendChild(button);
      });

      const rect = input.getBoundingClientRect();
      formulaSuggestions.style.left = `${Math.min(rect.left, window.innerWidth - 320)}px`;
      formulaSuggestions.style.top = `${Math.min(rect.bottom + 4, window.innerHeight - 220)}px`;
      formulaSuggestions.style.width = `${Math.max(rect.width, 240)}px`;
      formulaSuggestions.hidden = false;
    }

    async function resolveDealContext() {
      const urlParams = Object.fromEntries(new URLSearchParams(window.location.search).entries());
      const queryPlacementOptions = parsePlacementOptions(
        urlParams.PLACEMENT_OPTIONS || urlParams.placement_options || urlParams.options
      );
      let placementOptions = { ...queryPlacementOptions };

      const info = await getPlacementInfo();
      placementOptions = {
        ...placementOptions,
        ...info,
        ...(info.options ? parsePlacementOptions(info.options) : {}),
        ...(info.OPTIONS ? parsePlacementOptions(info.OPTIONS) : {}),
        ...(info.PLACEMENT_OPTIONS ? parsePlacementOptions(info.PLACEMENT_OPTIONS) : {}),
      };

      dealId = extractDealId({
        ...urlParams,
        ...placementOptions,
        queryPlacementOptions,
        currentUrl: window.location.href,
        referrer: document.referrer,
        url: document.referrer || window.location.href,
      });
      dealTitle = "";

      if (dealContext) {
        dealContext.textContent = dealId
          ? `Таблица сделки "ID ${dealId}"`
          : "Таблица сделки не определена. Обновите вкладку после полного открытия карточки.";
      }

      if (!dealId) activeSheetType = SHEET_TYPE_DEAL;
      loadSheetLists();
      loadActiveSheetState();
      await syncSheetListsFromSharedStorage();
      loadActiveSheetState();
    }

    async function loadDealFields(options = {}) {
      const compactAfterLoad = Boolean(options.compactAfterLoad);
      if (!dealId) {
        return;
      }

      try {
        const fields = await callMethod("crm.deal.fields");
        const deal = await callMethod("crm.deal.get", { id: dealId });
        const nextCategoryId = normalizeCategoryId(deal.CATEGORY_ID);
        const normalizedFields = normalizeFields(fields, deal);
        const displayValues = await loadDisplayValues(normalizedFields, deal);
        const previousFunnelStorageKey = getSheetListStorageKey(SHEET_TYPE_FUNNEL, dealId, dealCategoryId);
        dealCategoryId = nextCategoryId;
        dealCategoryName = displayValues.CATEGORY_ID || "";
        dealTitle = formatDealFieldValue(deal.TITLE) || `ID ${dealId}`;
        if (previousFunnelStorageKey !== getSheetListStorageKey(SHEET_TYPE_FUNNEL, dealId, dealCategoryId)) {
          loadSheetLists();
        }
        await syncSheetListsFromSharedStorage();
        if (storageKey !== getSheetStorageKey(activeSheetType, dealId, dealCategoryId, activeSheetIndex)) {
          loadActiveSheetState();
        } else {
          updateSheetModeControls();
        }
        dealFields = applyDisplayValues(normalizedFields, displayValues);
        refreshFieldBoundCells({ deferPersist: compactAfterLoad });
        if (compactAfterLoad) {
          compactGridToFilledBounds({ deferPersist: true });
          autoFitSheetSize();
          persistSheetState();
          renderGrid();
        }
      } catch (error) {
        if (window.console && typeof window.console.warn === "function") {
          window.console.warn("Deal fields loading failed", error);
        }
      }
    }

    function focusFirstSelectedCell() {
      const firstKey = selectedCells.values().next().value;
      if (!firstKey) return;

      const { rowIndex, columnIndex } = parseCellKey(firstKey);
      const input = table.querySelector(
        `.cell-input[data-row="${rowIndex}"][data-column="${columnIndex}"]`
      );
      if (input) {
        setCurrentCell(input, rowIndex, columnIndex);
        input.focus();
      }
    }

    function selectAllCells() {
      const columnCount = grid[0] ? grid[0].length : DEFAULT_COLUMNS;
      const keys = getRangeCellKeys(
        { rowIndex: 0, columnIndex: 0 },
        { rowIndex: grid.length - 1, columnIndex: columnCount - 1 }
      );
      selectionAnchor = { rowIndex: 0, columnIndex: 0 };
      setSelectedCells(keys);
      focusFirstSelectedCell();
    }

    function commitSheetStateChange(nextState, nextSelectionKeys = []) {
      applySheetState(nextState);
      selectedCells = new Set(nextSelectionKeys);
      persistSheetState();
      renderGrid();
      if (selectedCells.size) focusFirstSelectedCell();
    }

    function insertRowsAfter(rowIndex, count = 1) {
      const nextState = insertRowsInSheetState(getCurrentSheetSnapshot(), rowIndex, count);
      const columnCount = nextState.grid[0] ? nextState.grid[0].length : DEFAULT_COLUMNS;
      const firstInsertedRow = Math.min(rowIndex + 1, nextState.grid.length - 1);
      commitSheetStateChange(
        nextState,
        getRangeCellKeys(
          { rowIndex: firstInsertedRow, columnIndex: 0 },
          { rowIndex: Math.min(firstInsertedRow + normalizeInsertCount(count) - 1, nextState.grid.length - 1), columnIndex: columnCount - 1 }
        )
      );
    }

    function insertColumnsAfter(columnIndex, count = 1) {
      const nextState = insertColumnsInSheetState(getCurrentSheetSnapshot(), columnIndex, count);
      const firstInsertedColumn = Math.min(columnIndex + 1, (nextState.grid[0] ? nextState.grid[0].length : DEFAULT_COLUMNS) - 1);
      commitSheetStateChange(
        nextState,
        getRangeCellKeys(
          { rowIndex: 0, columnIndex: firstInsertedColumn },
          { rowIndex: nextState.grid.length - 1, columnIndex: Math.min(firstInsertedColumn + normalizeInsertCount(count) - 1, nextState.grid[0].length - 1) }
        )
      );
    }

    function moveRow(fromRowIndex, toRowIndex) {
      const nextState = moveRowInSheetState(getCurrentSheetSnapshot(), fromRowIndex, toRowIndex);
      const columnCount = nextState.grid[0] ? nextState.grid[0].length : DEFAULT_COLUMNS;
      commitSheetStateChange(
        nextState,
        getRangeCellKeys({ rowIndex: toRowIndex, columnIndex: 0 }, { rowIndex: toRowIndex, columnIndex: columnCount - 1 })
      );
    }

    function moveColumn(fromColumnIndex, toColumnIndex) {
      const nextState = moveColumnInSheetState(getCurrentSheetSnapshot(), fromColumnIndex, toColumnIndex);
      commitSheetStateChange(
        nextState,
        getRangeCellKeys({ rowIndex: 0, columnIndex: toColumnIndex }, { rowIndex: nextState.grid.length - 1, columnIndex: toColumnIndex })
      );
    }

    function deleteRow(rowIndex) {
      const nextState = deleteRowInSheetState(getCurrentSheetSnapshot(), rowIndex);
      const columnCount = nextState.grid[0] ? nextState.grid[0].length : DEFAULT_COLUMNS;
      const selectedRowIndex = Math.min(rowIndex, nextState.grid.length - 1);
      commitSheetStateChange(
        nextState,
        getRangeCellKeys({ rowIndex: selectedRowIndex, columnIndex: 0 }, { rowIndex: selectedRowIndex, columnIndex: columnCount - 1 })
      );
    }

    function deleteColumn(columnIndex) {
      const nextState = deleteColumnInSheetState(getCurrentSheetSnapshot(), columnIndex);
      const selectedColumnIndex = Math.min(columnIndex, (nextState.grid[0] ? nextState.grid[0].length : DEFAULT_COLUMNS) - 1);
      commitSheetStateChange(
        nextState,
        getRangeCellKeys({ rowIndex: 0, columnIndex: selectedColumnIndex }, { rowIndex: nextState.grid.length - 1, columnIndex: selectedColumnIndex })
      );
    }

    function closeGridContextMenu() {
      if (!gridContextMenu) return;
      gridContextMenu.remove();
      gridContextMenu = null;
    }

    function showGridContextMenu(event, type, index) {
      event.preventDefault();
      event.stopPropagation();
      closeGridContextMenu();

      gridContextMenu = document.createElement("div");
      gridContextMenu.className = "grid-context-menu";
      gridContextMenu.style.left = `${event.clientX}px`;
      gridContextMenu.style.top = `${event.clientY}px`;

      const actions =
        type === "row"
          ? [
              { label: "Добавить строку", run: () => insertRowsAfter(index, 1) },
              { label: "Добавить 4 строки", run: () => insertRowsAfter(index, 4) },
              { label: "Удалить строку", run: () => deleteRow(index) },
            ]
          : [
              { label: "Добавить столбец", run: () => insertColumnsAfter(index, 1) },
              { label: "Добавить 4 столбца", run: () => insertColumnsAfter(index, 4) },
              { label: "Удалить столбец", run: () => deleteColumn(index) },
            ];

      actions.forEach((action) => {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = action.label;
        button.addEventListener("click", () => {
          closeGridContextMenu();
          action.run();
        });
        gridContextMenu.appendChild(button);
      });

      document.body.appendChild(gridContextMenu);
    }

    function toggleWrapSelectedCells() {
      if (!selectedCells.size) return;

      const keys = Array.from(selectedCells);
      const shouldWrap = keys.some((key) => !wrappedCells.has(key));
      keys.forEach((key) => {
        if (shouldWrap) wrappedCells.add(key);
        else wrappedCells.delete(key);
      });
      if (shouldWrap) autoFitSheetSize(new Set(keys));

      persistSheetState();
      renderGrid();
      focusFirstSelectedCell();
    }

    function applyCellFormatToSelection(updateFormat) {
      if (!selectedCells.size) return;

      selectedCells.forEach((key) => {
        const nextFormat = normalizeCellFormat(updateFormat({ ...(cellFormats[key] || {}) }));
        const hasFormat = nextFormat.fillColor ||
          nextFormat.fontWeight ||
          nextFormat.fontStyle ||
          nextFormat.fontSize ||
          nextFormat.horizontalAlign ||
          nextFormat.verticalAlign;
        if (hasFormat) cellFormats[key] = nextFormat;
        else delete cellFormats[key];
      });

      persistSheetState();
      renderGrid();
      focusFirstSelectedCell();
    }

    function setSelectedFillColor(color) {
      applyCellFormatToSelection((format) => ({
        ...format,
        fillColor: color || "",
      }));
    }

    function setSelectedFontWeight(weight) {
      applyCellFormatToSelection((format) => ({
        ...format,
        fontWeight: weight || "",
      }));
    }

    function setSelectedFontSize(size) {
      applyCellFormatToSelection((format) => ({
        ...format,
        fontSize: size || "",
      }));
    }

    function toggleSelectedBold() {
      if (!selectedCells.size) return;

      const shouldEnable = Array.from(selectedCells).some((key) => !(cellFormats[key] || {}).fontWeight);
      setSelectedFontWeight(shouldEnable ? "700" : "");
    }

    function toggleSelectedItalic() {
      if (!selectedCells.size) return;

      const shouldEnable = Array.from(selectedCells).some((key) => (cellFormats[key] || {}).fontStyle !== "italic");
      applyCellFormatToSelection((format) => ({
        ...format,
        fontStyle: shouldEnable ? "italic" : "",
      }));
    }

    function setSelectedHorizontalAlign(alignment) {
      applyCellFormatToSelection((format) => ({
        ...format,
        horizontalAlign: alignment || "",
      }));
    }

    function setSelectedVerticalAlign(alignment) {
      applyCellFormatToSelection((format) => ({
        ...format,
        verticalAlign: alignment || "",
      }));
    }

    function autoFitSelectedCells() {
      autoFitSheetSize(selectedCells);
      persistSheetState();
      renderGrid();
      focusFirstSelectedCell();
    }

    function writeResultToCurrentCell(result) {
      if (!currentCell || !currentCell.input) {
        focusFirstSelectedCell();
      }

      const target = currentCell || (() => {
        const firstKey = selectedCells.values().next().value;
        if (!firstKey) return null;
        const cell = parseCellKey(firstKey);
        return { input: null, rowIndex: cell.rowIndex, columnIndex: cell.columnIndex };
      })();

      if (!target) return;

      const key = cellKey(target.rowIndex, target.columnIndex);
      grid[target.rowIndex][target.columnIndex] = result;
      delete fieldBindings[key];
      if (target.input) target.input.value = result;
      persistSheetState();
      renderGrid();
      focusFirstSelectedCell();
    }

    function downloadExcelFile() {
      const html = buildExcelHtml(grid, cellFormats, { columnWidths });
      const blob = new Blob(["\ufeff", html], { type: "application/vnd.ms-excel;charset=utf-8" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = getExportFileName(dealId);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }

    addRowButton.addEventListener("click", () => {
      insertRowsAfter(grid.length - 1, 1);
    });

    addColumnButton.addEventListener("click", () => {
      insertColumnsAfter((grid[0] ? grid[0].length : DEFAULT_COLUMNS) - 1, 1);
    });

    reloadFieldsButton.addEventListener("click", () => loadDealFields({ compactAfterLoad: true }));
    if (addDealSheetButton) addDealSheetButton.addEventListener("click", () => addSheetToGroup(SHEET_TYPE_DEAL));
    if (addFunnelSheetButton) addFunnelSheetButton.addEventListener("click", () => addSheetToGroup(SHEET_TYPE_FUNNEL));
    window.addEventListener("pagehide", pruneEmptySheetsBeforeExit);
    window.addEventListener("beforeunload", pruneEmptySheetsBeforeExit);
    if (selectFilledButton) selectFilledButton.addEventListener("click", selectAllCells);
    if (copyCellsButton) copyCellsButton.addEventListener("click", copySelectedCells);
    if (pasteCellsButton) pasteCellsButton.addEventListener("click", pasteCopiedCells);
    if (clearSelectionButton) clearSelectionButton.addEventListener("click", requestClearSelectedCells);
    if (undoButton) undoButton.addEventListener("click", undoSheetState);
    if (redoButton) redoButton.addEventListener("click", redoSheetState);
    if (wrapTextButton) wrapTextButton.addEventListener("click", toggleWrapSelectedCells);
    if (autoFitButton) autoFitButton.addEventListener("click", autoFitSelectedCells);
    if (fillColorSelect) fillColorSelect.addEventListener("change", () => setSelectedFillColor(fillColorSelect.value));
    if (fontSizeSelect) fontSizeSelect.addEventListener("change", () => setSelectedFontSize(fontSizeSelect.value));
    if (boldButton) boldButton.addEventListener("click", toggleSelectedBold);
    if (italicButton) italicButton.addEventListener("click", toggleSelectedItalic);
    if (alignTopButton) alignTopButton.addEventListener("click", () => setSelectedVerticalAlign("top"));
    if (alignMiddleButton) alignMiddleButton.addEventListener("click", () => setSelectedVerticalAlign("middle"));
    if (alignBottomButton) alignBottomButton.addEventListener("click", () => setSelectedVerticalAlign("bottom"));
    if (alignLeftButton) alignLeftButton.addEventListener("click", () => setSelectedHorizontalAlign("left"));
    if (alignCenterButton) alignCenterButton.addEventListener("click", () => setSelectedHorizontalAlign("center"));
    if (alignRightButton) alignRightButton.addEventListener("click", () => setSelectedHorizontalAlign("right"));
    if (exportExcelButton) exportExcelButton.addEventListener("click", downloadExcelFile);
    if (formulaLibraryButton) formulaLibraryButton.addEventListener("click", openFormulaModal);
    if (formulaModalClose) formulaModalClose.addEventListener("click", closeFormulaModal);
    if (cancelFormulaButton) cancelFormulaButton.addEventListener("click", closeFormulaModal);
    if (saveFormulaButton) saveFormulaButton.addEventListener("click", saveFormulaFromModal);
    if (applyFormulaButton) applyFormulaButton.addEventListener("click", applySelectedFormulaToCell);
    if (formulaInput) formulaInput.addEventListener("input", handleFormulaInput);
    if (formulaModal) {
      formulaModal.addEventListener("click", (event) => {
        if (event.target === formulaModal) closeFormulaModal();
      });
    }
    if (deleteConfirmClose) deleteConfirmClose.addEventListener("click", closeDeleteConfirmModal);
    if (cancelDeleteButton) cancelDeleteButton.addEventListener("click", closeDeleteConfirmModal);
    if (confirmDeleteButton) confirmDeleteButton.addEventListener("click", clearSelectedCells);
    if (deleteConfirmModal) {
      deleteConfirmModal.addEventListener("click", (event) => {
        if (event.target === deleteConfirmModal) closeDeleteConfirmModal();
      });
    }
    if (helpButton) helpButton.addEventListener("click", openHelpModal);
    if (helpModalClose) helpModalClose.addEventListener("click", closeHelpModal);
    if (helpModal) {
      helpModal.addEventListener("click", (event) => {
        if (event.target === helpModal) closeHelpModal();
      });
    }
    if (fieldPopoverClose) fieldPopoverClose.addEventListener("click", closeFieldPopover);
    fieldSearch.addEventListener("input", () => renderFieldList(fieldSearch.value));
    document.addEventListener("pointermove", (event) => {
      updateDragSelectionFromPointer(event);
      updateGridResize(event);
      updateHeaderDrag(event);
    });
    document.addEventListener("pointerup", (event) => {
      endDragSelection(event);
      endGridResize(event);
      endHeaderDrag(event);
    });
    document.addEventListener("pointercancel", (event) => {
      endDragSelection(event);
      endGridResize(event);
      endHeaderDrag(event);
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closeFieldPopover();
        closeGridContextMenu();
        closeFormulaModal();
        closeDeleteConfirmModal();
        closeFormulaSuggestions();
      }
      if (isCopyShortcut(event) && selectedCells.size && shouldUseCellClipboardShortcut(event)) {
        copySelectedCells();
      }
      if (isPasteShortcut(event) && cellClipboard && shouldUseCellClipboardShortcut(event)) {
        event.preventDefault();
        pasteCopiedCells();
      }
    });
    document.addEventListener("copy", (event) => {
      if (!selectedCells.size || !shouldUseCellClipboardShortcut(event)) return;
      const clipboard = copySelectedCells();
      if (!clipboard || !event.clipboardData) return;
      event.clipboardData.setData("text/plain", getClipboardText(clipboard));
      event.preventDefault();
    });
    document.addEventListener("paste", (event) => {
      if (!cellClipboard || !shouldUseCellClipboardShortcut(event)) return;
      event.preventDefault();
      pasteCopiedCells();
    });
    document.addEventListener("click", (event) => {
      if (popover.hidden) return;
      if (popover.contains(event.target) || event.target.classList.contains("field-picker-button")) return;
      closeFieldPopover();
    });
    document.addEventListener("click", (event) => {
      if (gridContextMenu && !gridContextMenu.contains(event.target)) closeGridContextMenu();
    });
    document.addEventListener("click", (event) => {
      if (!formulaSuggestions || formulaSuggestions.hidden) return;
      if (formulaSuggestions.contains(event.target) || event.target === formulaSuggestionInput) return;
      closeFormulaSuggestions();
    });

    loadSheetLists();
    bindHistoryToStorageKey();
    renderGrid();
    updateSheetModeControls();

    if (!window.BX24 || typeof window.BX24.init !== "function") {
      if (dealContext) dealContext.textContent = `${DISPLAY_VERSION}. Локальный режим без Bitrix24 SDK.`;
      return;
    }

    window.BX24.init(async () => {
      await initializeSharedStorage();
      await resolveDealContext();
      await loadDealFields();
    });
  }

  if (typeof window !== "undefined" && typeof document !== "undefined") {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", bootBrowserApp);
    } else {
      bootBrowserApp();
    }
  }

  return {
    DEFAULT_COLUMNS,
    DEFAULT_COLUMN_WIDTH,
    DEFAULT_ROW_HEIGHT,
    DEFAULT_ROWS,
    MAX_SHEETS_PER_GROUP,
    MAX_GRID_COLUMNS,
    MAX_GRID_ROWS,
    addColumn,
    addRow,
    addRecentFormula,
    addSavedFormula,
    appendFormulaReference,
    applyFormulaToGridCell,
    applyFieldBindings,
    areSheetSnapshotsEqual,
    buildExcelHtml,
    calculateSelectedCells,
    cellKey,
    clampColumnWidth,
    clampRowHeight,
    clearCellSelectionState,
    columnName,
    columnIndexFromName,
    cloneSheetSnapshot,
    countBoundFieldsOnSheet,
    createCellClipboard,
    createGrid,
    deleteColumnInSheetState,
    deleteRowInSheetState,
    evaluateFormula,
    escapeHtml,
    extractDealId,
    findCategoryName,
    findStatusName,
    formatCompany,
    formatContact,
    formatDealFieldValue,
    formatFormulaInput,
    formatUser,
    getCellDisplayValue,
    getCellRangeBounds,
    getClipboardText,
    getClipboardTargetCellKeys,
    getDealStageEntityId,
    getExportCellContent,
    getExportCellType,
    getExportFileName,
    getExportGrid,
    getTrimmedSheetState,
    getAutoFitColumnWidths,
    getFilledCellKeys,
    getFunnelStorageKey,
    getDeleteIndexMap,
    getInsertIndexMap,
    getMoveIndexMap,
    getSortedCellKeys,
    getGridStorageKey,
    getNormalizedSheetState,
    getRangeCellKeys,
    getSelectedColumns,
    getSelectedRows,
    getSelectionEdgeClassNames,
    getSheetStorageKey,
    getSheetListStorageKey,
    isSheetStateEmpty,
    normalizeSheetList,
    pasteCellClipboard,
    pruneEmptySheetList,
    getUsedGridBounds,
    isCopyShortcut,
    isPasteShortcut,
    inferFieldBindingsFromGrid,
    insertColumnsInSheetState,
    insertRowsInSheetState,
    loadRecentFormulas,
    loadSavedFormulas,
    loadSheetState,
    parseSheetListData,
    parseSheetStateData,
    measureRowHeight,
    measureColumnWidth,
    moveArrayItem,
    moveColumnInSheetState,
    moveRowInSheetState,
    getAutoFitRowHeights,
    normalizeCategoryId,
    normalizeCellFormat,
    normalizeFields,
    normalizeIdList,
    normalizeColumnWidths,
    normalizeRecentFormulas,
    normalizeRowHeights,
    parseCellKey,
    parseCellNumber,
    parseExportNumber,
    parseFormulaReference,
    normalizeSavedFormula,
    normalizeSavedFormulas,
    removeSavedFormula,
    remapFormulaReferences,
    saveRecentFormulas,
    saveSavedFormulas,
    saveSheetState,
    serializeSheetList,
    serializeSheetState,
    sanitizeFormulaInput,
    shiftFormulaReferences,
  };
});
