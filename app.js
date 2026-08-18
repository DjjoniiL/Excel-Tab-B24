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
  const DISPLAY_VERSION = "Excel Tab B24 v.0.1 Marketplace B24";

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
        value: deal[id],
      }))
      .sort((left, right) => left.title.localeCompare(right.title, "ru"));
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

  function loadGrid() {
    try {
      const saved = window.localStorage.getItem(STORAGE_KEY);
      const parsed = saved ? JSON.parse(saved) : null;
      if (Array.isArray(parsed) && parsed.length && Array.isArray(parsed[0])) return parsed;
    } catch (error) {
      window.localStorage.removeItem(STORAGE_KEY);
    }
    return createGrid();
  }

  function saveGrid(grid) {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(grid));
  }

  function bootBrowserApp() {
    const table = document.getElementById("sheet");
    const cellTemplate = document.getElementById("cellTemplate");
    const popover = document.getElementById("fieldPopover");
    const fieldList = document.getElementById("fieldList");
    const fieldSearch = document.getElementById("fieldSearch");
    const fieldStatus = document.getElementById("fieldStatus");
    const gridStatus = document.getElementById("gridStatus");
    const dealContext = document.getElementById("dealContext");
    const addRowButton = document.getElementById("addRowButton");
    const addColumnButton = document.getElementById("addColumnButton");
    const reloadFieldsButton = document.getElementById("reloadFieldsButton");

    let grid = loadGrid();
    let currentCell = null;
    let dealFields = [];
    let dealId = null;

    function updateGridStatus() {
      const rows = grid.length;
      const columns = grid[0] ? grid[0].length : 0;
      gridStatus.textContent = `${rows} строк, ${columns} столбцов`;
    }

    function renderGrid() {
      table.innerHTML = "";
      const columnCount = grid[0] ? grid[0].length : DEFAULT_COLUMNS;
      const thead = document.createElement("thead");
      const headRow = document.createElement("tr");
      const corner = document.createElement("th");
      corner.className = "corner-heading";
      headRow.appendChild(corner);

      for (let column = 0; column < columnCount; column += 1) {
        const th = document.createElement("th");
        th.textContent = columnName(column);
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
        tr.appendChild(heading);

        row.forEach((value, columnIndex) => {
          const td = document.createElement("td");
          const fragment = cellTemplate.content.cloneNode(true);
          const input = fragment.querySelector(".cell-input");
          const picker = fragment.querySelector(".field-picker-button");
          input.value = value;
          input.dataset.row = String(rowIndex);
          input.dataset.column = String(columnIndex);
          picker.dataset.row = String(rowIndex);
          picker.dataset.column = String(columnIndex);
          input.addEventListener("input", () => {
            grid[rowIndex][columnIndex] = input.value;
            saveGrid(grid);
          });
          picker.addEventListener("click", (event) => {
            currentCell = { input, rowIndex, columnIndex };
            openFieldPopover(event.currentTarget);
          });
          td.appendChild(fragment);
          tr.appendChild(td);
        });

        tbody.appendChild(tr);
      });

      table.appendChild(tbody);
      updateGridStatus();
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
          if (!currentCell) return;
          const formatted = formatDealFieldValue(field.value);
          currentCell.input.value = formatted;
          grid[currentCell.rowIndex][currentCell.columnIndex] = formatted;
          saveGrid(grid);
          closeFieldPopover();
          currentCell.input.focus();
        });
        fieldList.appendChild(button);
      });
    }

    function openFieldPopover(anchor) {
      renderFieldList(fieldSearch.value);
      const rect = anchor.getBoundingClientRect();
      popover.hidden = false;
      popover.style.left = `${Math.min(rect.left, window.innerWidth - 300)}px`;
      popover.style.top = `${Math.min(rect.bottom + 4, window.innerHeight - 310)}px`;
      fieldSearch.focus();
    }

    function closeFieldPopover() {
      popover.hidden = true;
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
        dealFields = normalizeFields(fields, deal);
        fieldStatus.textContent = `Поля сделки: ${dealFields.length}`;
      } catch (error) {
        fieldStatus.textContent = `Поля сделки: ошибка (${error.message || error})`;
      }
    }

    addRowButton.addEventListener("click", () => {
      grid = addRow(grid);
      saveGrid(grid);
      renderGrid();
    });

    addColumnButton.addEventListener("click", () => {
      grid = addColumn(grid);
      saveGrid(grid);
      renderGrid();
    });

    reloadFieldsButton.addEventListener("click", loadDealFields);
    fieldSearch.addEventListener("input", () => renderFieldList(fieldSearch.value));
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
    columnName,
    createGrid,
    extractDealId,
    formatDealFieldValue,
    normalizeFields,
  };
});
