(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.ExcelTabB24 = factory();
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const DEFAULT_ROWS = 12;
  const DEFAULT_COLUMNS = 8;
  const STORAGE_KEY = "excel-tab-b24-grid-v1";
  const DEAL_STORAGE_KEY_PREFIX = "excel-tab-b24-grid-deal-v1";
  const DISPLAY_VERSION = "Excel Tab B24 v.5 Marketplace B24";
  const DEFAULT_COLUMN_WIDTH = 132;
  const MAX_COLUMN_WIDTH = 420;
  const MIN_COLUMN_WIDTH = 90;
  const FILL_COLORS = ["", "#fff2cc", "#d9ead3", "#cfe2f3", "#f4cccc", "#eadcf8"];
  const FONT_WEIGHTS = ["", "600", "700", "800"];

  function createGrid(rows = DEFAULT_ROWS, columns = DEFAULT_COLUMNS) {
    return Array.from({ length: rows }, () => Array.from({ length: columns }, () => ""));
  }

  function addRow(grid) {
    const width = Math.max(1, grid[0] ? grid[0].length : DEFAULT_COLUMNS);
    return [...grid, Array.from({ length: width }, () => "")];
  }

  function addColumn(grid) {
    const source = grid.length ? grid : createGrid(DEFAULT_ROWS, 0);
    return source.map((row) => [...row, ""]);
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

  function measureColumnWidth(grid, columnIndex) {
    const maxLength = grid.reduce((max, row) => {
      const textLength = String(row[columnIndex] || "").length;
      return Math.max(max, textLength);
    }, columnName(columnIndex).length);
    return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.ceil(maxLength * 8.5) + 42));
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

  function calculateSelectedCells(grid, selectedCells, operation) {
    const values = getSortedCellKeys(selectedCells)
      .map((key) => {
        const { rowIndex, columnIndex } = parseCellKey(key);
        return parseCellNumber(grid[rowIndex] && grid[rowIndex][columnIndex]);
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

  function buildExcelHtml(grid, cellFormats = {}) {
    const rows = getExportGrid(grid);
    const body = rows
      .map((row, rowIndex) => {
        const cells = row
          .map((value, columnIndex) => {
            const style = getExportCellStyle(cellFormats[cellKey(rowIndex, columnIndex)]);
            return `<td${style ? ` style="${style}"` : ""}>${escapeHtml(value)}</td>`;
          })
          .join("");
        return `<tr>${cells}</tr>`;
      })
      .join("");

    return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    table { border-collapse: collapse; }
    td { border: 1px solid #d7dde8; mso-number-format:"\\@"; padding: 4px 8px; white-space: pre-wrap; }
  </style>
</head>
<body>
  <table>${body}</table>
</body>
</html>`;
  }

  function getExportCellStyle(format = {}) {
    const normalizedFormat = normalizeCellFormat(format);
    const styles = [];
    if (normalizedFormat.fillColor) styles.push(`background-color:${normalizedFormat.fillColor}`);
    if (normalizedFormat.fontWeight) styles.push(`font-weight:${normalizedFormat.fontWeight}`);
    return styles.join(";");
  }

  function getExportFileName(dealId) {
    const suffix = dealId ? ` deal ${dealId}` : "";
    return `Excel Tab B24${suffix}.xls`;
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
    const candidates = [
      ...idKeys.map((key) => input[key]),
      ...collectNestedValues(input, idKeys),
    ];

    for (const candidate of candidates) {
      const value = Number.parseInt(candidate, 10);
      if (Number.isInteger(value) && value > 0) return value;
    }

    const textSources = [input.url, input.URI, input.PLACEMENT_OPTIONS, input.placement_options]
      .map((value) => (typeof value === "string" ? value : ""))
      .filter(Boolean);

    for (const text of textSources) {
      const match = text.match(/\/crm\/deal\/(?:details|show)\/(\d+)\//i);
      if (match) return Number.parseInt(match[1], 10);
    }

    return null;
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
    return compactName([user.LAST_NAME, user.NAME, user.SECOND_NAME]) || user.EMAIL || String(user.ID || "");
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
    if (!Array.isArray(categories)) return "";
    const category = categories.find((item) => String(item.ID) === String(categoryId));
    return category ? category.NAME || category.TITLE || String(category.ID || "") : "";
  }

  function applyDisplayValues(fields, displayValues = {}) {
    return fields.map((field) => ({
      ...field,
      value: Object.prototype.hasOwnProperty.call(displayValues, field.id) ? displayValues[field.id] : field.value,
    }));
  }

  function applyFieldBindings(grid, fieldBindings = {}, fields = []) {
    const valuesByFieldId = fields.reduce((values, field) => {
      values[field.id] = formatDealFieldValue(field.value);
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

  function formatDealFieldValue(value) {
    if (value === null || typeof value === "undefined") return "";
    if (Array.isArray(value)) return value.map(formatDealFieldValue).filter(Boolean).join(", ");
    if (typeof value === "object") {
      if ("VALUE" in value) return formatDealFieldValue(value.VALUE);
      if ("value" in value) return formatDealFieldValue(value.value);
      return JSON.stringify(value);
    }
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
        const users = await callOptionalMethod("user.get", { FILTER: { ID: id } }, []);
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

    return STORAGE_KEY;
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

  function loadSheetState(storageKey = STORAGE_KEY) {
    try {
      const saved = window.localStorage.getItem(storageKey);
      const parsed = saved ? JSON.parse(saved) : null;
      if (Array.isArray(parsed) && parsed.length && Array.isArray(parsed[0])) {
        return { cellFormats: {}, columnWidths: [], fieldBindings: {}, grid: parsed, wrappedCells: new Set() };
      }
      if (parsed && Array.isArray(parsed.grid) && parsed.grid.length && Array.isArray(parsed.grid[0])) {
        return {
          cellFormats: parsed.cellFormats && typeof parsed.cellFormats === "object" ? parsed.cellFormats : {},
          columnWidths: Array.isArray(parsed.columnWidths) ? parsed.columnWidths : [],
          fieldBindings: parsed.fieldBindings && typeof parsed.fieldBindings === "object" ? parsed.fieldBindings : {},
          grid: parsed.grid,
          wrappedCells: new Set(Array.isArray(parsed.wrappedCells) ? parsed.wrappedCells : []),
        };
      }
    } catch (error) {
      window.localStorage.removeItem(storageKey);
    }

    return { cellFormats: {}, columnWidths: [], fieldBindings: {}, grid: createGrid(), wrappedCells: new Set() };
  }

  function saveSheetState(state, storageKey = STORAGE_KEY) {
    window.localStorage.setItem(
      storageKey,
      JSON.stringify({
        cellFormats: state.cellFormats || {},
        columnWidths: state.columnWidths || [],
        fieldBindings: state.fieldBindings || {},
        grid: state.grid,
        wrappedCells: Array.from(state.wrappedCells || []),
      })
    );
  }

  function bootBrowserApp() {
    const table = document.getElementById("sheet");
    const cellTemplate = document.getElementById("cellTemplate");
    const popover = document.getElementById("fieldPopover");
    const fieldPopoverClose = document.getElementById("fieldPopoverClose");
    const fieldList = document.getElementById("fieldList");
    const fieldSearch = document.getElementById("fieldSearch");
    const fieldStatus = document.getElementById("fieldStatus");
    const gridStatus = document.getElementById("gridStatus");
    const dealContext = document.getElementById("dealContext");
    const addRowButton = document.getElementById("addRowButton");
    const addColumnButton = document.getElementById("addColumnButton");
    const reloadFieldsButton = document.getElementById("reloadFieldsButton");
    const selectFilledButton = document.getElementById("selectFilledButton");
    const selectionActions = document.getElementById("selectionActions");
    const wrapTextButton = document.getElementById("wrapTextButton");
    const autoFitButton = document.getElementById("autoFitButton");
    const fillColorSelect = document.getElementById("fillColorSelect");
    const fontWeightSelect = document.getElementById("fontWeightSelect");
    const boldButton = document.getElementById("boldButton");
    const calcAddButton = document.getElementById("calcAddButton");
    const calcSubtractButton = document.getElementById("calcSubtractButton");
    const calcMultiplyButton = document.getElementById("calcMultiplyButton");
    const calcDivideButton = document.getElementById("calcDivideButton");
    const calcStatus = document.getElementById("calcStatus");
    const exportExcelButton = document.getElementById("exportExcelButton");

    let storageKey = getGridStorageKey(null);
    let sheetState = loadSheetState(storageKey);
    let grid = sheetState.grid;
    let cellFormats = sheetState.cellFormats;
    let wrappedCells = sheetState.wrappedCells;
    let columnWidths = sheetState.columnWidths;
    let fieldBindings = sheetState.fieldBindings;
    let currentCell = null;
    let dealFields = [];
    let dealId = null;
    let popoverAnchor = null;
    let selectedCells = new Set();
    let selectionAnchor = null;

    function persistSheetState() {
      saveSheetState({ cellFormats, columnWidths, fieldBindings, grid, wrappedCells }, storageKey);
    }

    function updateSelectionActions() {
      if (!selectionActions) return;
      selectionActions.hidden = selectedCells.size === 0;
    }

    function paintSelection() {
      table.querySelectorAll("td[data-cell-key]").forEach((cell) => {
        cell.classList.toggle("is-selected", selectedCells.has(cell.dataset.cellKey));
      });
      updateSelectionActions();
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
        if (next.has(key)) next.delete(key);
        else next.add(key);
        selectedCells = next;
        selectionAnchor = { rowIndex, columnIndex };
        paintSelection();
      } else {
        selectionAnchor = { rowIndex, columnIndex };
        setSelectedCells([key]);
      }
    }

    function setCurrentCell(input, rowIndex, columnIndex) {
      currentCell = { input, rowIndex, columnIndex };
    }

    function updateGridStatus() {
      const rows = grid.length;
      const columns = grid[0] ? grid[0].length : 0;
      gridStatus.textContent = `${rows} строк, ${columns} столбцов`;
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
        col.style.width = `${columnWidths[column] || DEFAULT_COLUMN_WIDTH}px`;
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
        th.textContent = columnName(column);
        th.title = "Выделить столбец";
        th.addEventListener("click", (event) => {
          const columnCells = grid.map((row, rowIndex) => cellKey(rowIndex, column));
          if (event.ctrlKey || event.metaKey) {
            const next = new Set(selectedCells);
            columnCells.forEach((key) => next.add(key));
            selectedCells = next;
            paintSelection();
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
        const heading = document.createElement("th");
        heading.className = "row-heading";
        heading.textContent = String(rowIndex + 1);
        heading.title = "Выделить строку";
        heading.addEventListener("click", (event) => {
          const rowCells = row.map((value, columnIndex) => cellKey(rowIndex, columnIndex));
          if (event.ctrlKey || event.metaKey) {
            const next = new Set(selectedCells);
            rowCells.forEach((key) => next.add(key));
            selectedCells = next;
            paintSelection();
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
          td.classList.toggle("is-selected", selectedCells.has(key));
          td.classList.toggle("is-wrapped", wrappedCells.has(key));
          const format = normalizeCellFormat(cellFormats[key] || {});
          if (format.fillColor) td.style.backgroundColor = format.fillColor;
          const fragment = cellTemplate.content.cloneNode(true);
          const input = fragment.querySelector(".cell-input");
          const picker = fragment.querySelector(".field-picker-button");
          input.value = value;
          if (format.fillColor) input.style.backgroundColor = format.fillColor;
          if (format.fontWeight) input.style.fontWeight = format.fontWeight;
          input.dataset.row = String(rowIndex);
          input.dataset.column = String(columnIndex);
          picker.dataset.row = String(rowIndex);
          picker.dataset.column = String(columnIndex);
          td.addEventListener("click", (event) => {
            setCurrentCell(input, rowIndex, columnIndex);
            selectCell(rowIndex, columnIndex, {
              shiftKey: event.shiftKey,
              toggleKey: event.ctrlKey || event.metaKey,
            });
          });
          input.addEventListener("mousedown", (event) => {
            setCurrentCell(input, rowIndex, columnIndex);
            selectCell(rowIndex, columnIndex, {
              shiftKey: event.shiftKey,
              toggleKey: event.ctrlKey || event.metaKey,
            });
          });
          input.addEventListener("click", (event) => {
            event.stopPropagation();
          });
          input.addEventListener("focus", () => {
            setCurrentCell(input, rowIndex, columnIndex);
            if (!selectedCells.has(key)) selectCell(rowIndex, columnIndex);
          });
          input.addEventListener("input", () => {
            grid[rowIndex][columnIndex] = input.value;
            delete fieldBindings[key];
            persistSheetState();
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
      updateGridStatus();
      updateSelectionActions();
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
          const formatted = formatDealFieldValue(field.value);
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
        url: document.referrer || window.location.href,
      });

      dealContext.textContent = dealId
        ? `Сделка #${dealId}`
        : "Сделка не определена. Обновите вкладку после полного открытия карточки.";

      storageKey = getGridStorageKey(dealId);
      sheetState = loadSheetState(storageKey);
      grid = sheetState.grid;
      cellFormats = sheetState.cellFormats;
      wrappedCells = sheetState.wrappedCells;
      columnWidths = sheetState.columnWidths;
      fieldBindings = sheetState.fieldBindings;
      currentCell = null;
      selectedCells = new Set();
      selectionAnchor = null;
      renderGrid();
    }

    async function loadDealFields() {
      if (!dealId) {
        fieldStatus.textContent = "Поля сделки: карточка не определена";
        return;
      }

      fieldStatus.textContent = "Поля сделки: загрузка...";
      try {
        const fields = await callMethod("crm.deal.fields");
        const deal = await callMethod("crm.deal.get", { id: dealId });
        const normalizedFields = normalizeFields(fields, deal);
        const displayValues = await loadDisplayValues(normalizedFields, deal);
        dealFields = applyDisplayValues(normalizedFields, displayValues);
        const updatedBindings = applyFieldBindings(grid, fieldBindings, dealFields);
        if (updatedBindings.changed) {
          grid = updatedBindings.grid;
          persistSheetState();
          renderGrid();
        }
        fieldStatus.textContent = `Поля сделки: ${dealFields.length}`;
      } catch (error) {
        fieldStatus.textContent = `Поля сделки: ошибка (${error.message || error})`;
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

    function selectFilledCells() {
      const keys = getFilledCellKeys(grid);
      if (!keys.length) {
        setSelectedCells([]);
        return;
      }

      const first = parseCellKey(keys[0]);
      selectionAnchor = first;
      setSelectedCells(keys);
      focusFirstSelectedCell();
    }

    function toggleWrapSelectedCells() {
      if (!selectedCells.size) return;

      const keys = Array.from(selectedCells);
      const shouldWrap = keys.some((key) => !wrappedCells.has(key));
      keys.forEach((key) => {
        if (shouldWrap) wrappedCells.add(key);
        else wrappedCells.delete(key);
      });

      persistSheetState();
      renderGrid();
      focusFirstSelectedCell();
    }

    function applyCellFormatToSelection(updateFormat) {
      if (!selectedCells.size) return;

      selectedCells.forEach((key) => {
        const nextFormat = normalizeCellFormat(updateFormat({ ...(cellFormats[key] || {}) }));
        const hasFormat = nextFormat.fillColor || nextFormat.fontWeight;
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

    function toggleSelectedBold() {
      if (!selectedCells.size) return;

      const shouldEnable = Array.from(selectedCells).some((key) => !(cellFormats[key] || {}).fontWeight);
      setSelectedFontWeight(shouldEnable ? "700" : "");
    }

    function autoFitSelectedColumns() {
      const columns = selectedCells.size
        ? getSelectedColumns(selectedCells)
        : Array.from({ length: grid[0] ? grid[0].length : DEFAULT_COLUMNS }, (item, index) => index);

      columns.forEach((columnIndex) => {
        columnWidths[columnIndex] = measureColumnWidth(grid, columnIndex);
      });

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

    function runCalculation(operation) {
      if (!selectedCells.size) return;

      const result = calculateSelectedCells(grid, selectedCells, operation);
      if (result.error) {
        if (calcStatus) calcStatus.textContent = `Расчет: ${result.error}`;
        return;
      }

      writeResultToCurrentCell(result.value);
      if (calcStatus) calcStatus.textContent = `Расчет: ${result.value}`;
    }

    function downloadExcelFile() {
      const html = buildExcelHtml(grid, cellFormats);
      const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = getExportFileName(dealId);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }

    addRowButton.addEventListener("click", () => {
      grid = addRow(grid);
      persistSheetState();
      renderGrid();
    });

    addColumnButton.addEventListener("click", () => {
      grid = addColumn(grid);
      columnWidths[grid[0].length - 1] = DEFAULT_COLUMN_WIDTH;
      persistSheetState();
      renderGrid();
    });

    reloadFieldsButton.addEventListener("click", loadDealFields);
    if (selectFilledButton) selectFilledButton.addEventListener("click", selectFilledCells);
    if (wrapTextButton) wrapTextButton.addEventListener("click", toggleWrapSelectedCells);
    if (autoFitButton) autoFitButton.addEventListener("click", autoFitSelectedColumns);
    if (fillColorSelect) fillColorSelect.addEventListener("change", () => setSelectedFillColor(fillColorSelect.value));
    if (fontWeightSelect) fontWeightSelect.addEventListener("change", () => setSelectedFontWeight(fontWeightSelect.value));
    if (boldButton) boldButton.addEventListener("click", toggleSelectedBold);
    if (calcAddButton) calcAddButton.addEventListener("click", () => runCalculation("add"));
    if (calcSubtractButton) calcSubtractButton.addEventListener("click", () => runCalculation("subtract"));
    if (calcMultiplyButton) calcMultiplyButton.addEventListener("click", () => runCalculation("multiply"));
    if (calcDivideButton) calcDivideButton.addEventListener("click", () => runCalculation("divide"));
    if (exportExcelButton) exportExcelButton.addEventListener("click", downloadExcelFile);
    if (fieldPopoverClose) fieldPopoverClose.addEventListener("click", closeFieldPopover);
    fieldSearch.addEventListener("input", () => renderFieldList(fieldSearch.value));
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeFieldPopover();
    });
    document.addEventListener("click", (event) => {
      if (popover.hidden) return;
      if (popover.contains(event.target) || event.target.classList.contains("field-picker-button")) return;
      closeFieldPopover();
    });

    renderGrid();

    if (!window.BX24 || typeof window.BX24.init !== "function") {
      dealContext.textContent = `${DISPLAY_VERSION}. Локальный режим без Bitrix24 SDK.`;
      fieldStatus.textContent = "Поля сделки: локальный режим";
      return;
    }

    window.BX24.init(async () => {
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
    DEFAULT_ROWS,
    addColumn,
    addRow,
    applyFieldBindings,
    buildExcelHtml,
    calculateSelectedCells,
    cellKey,
    columnName,
    createGrid,
    escapeHtml,
    extractDealId,
    findCategoryName,
    findStatusName,
    formatCompany,
    formatContact,
    formatDealFieldValue,
    formatUser,
    getDealStageEntityId,
    getExportFileName,
    getExportGrid,
    getFilledCellKeys,
    getSortedCellKeys,
    getGridStorageKey,
    getRangeCellKeys,
    getSelectedColumns,
    getUsedGridBounds,
    parseCellNumber,
    loadSheetState,
    measureColumnWidth,
    normalizeCellFormat,
    normalizeFields,
    normalizeIdList,
    parseCellKey,
    saveSheetState,
  };
});
