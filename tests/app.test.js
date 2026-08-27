"use strict";

const assert = require("node:assert/strict");
const app = require("../app.js");

function withMockLocalStorage(callback) {
  const originalWindow = global.window;
  const storage = new Map();
  global.window = {
    localStorage: {
      getItem: (key) => (storage.has(key) ? storage.get(key) : null),
      removeItem: (key) => storage.delete(key),
      setItem: (key, value) => storage.set(key, String(value)),
    },
  };

  try {
    callback(storage);
  } finally {
    global.window = originalWindow;
  }
}

function testCreateGridDefaults() {
  const grid = app.createGrid();
  assert.equal(grid.length, 9);
  assert.equal(grid[0].length, 7);
  assert.equal(grid[8][6], "");
}

function testAddRowAndColumn() {
  const grid = [["a", "b"]];
  const withRow = app.addRow(grid);
  assert.deepEqual(withRow, [["a", "b"], ["", ""]]);

  const withColumn = app.addColumn(withRow);
  assert.deepEqual(withColumn, [["a", "b", ""], ["", "", ""]]);
}

function testColumnName() {
  assert.equal(app.columnName(0), "A");
  assert.equal(app.columnName(25), "Z");
  assert.equal(app.columnName(26), "AA");
  assert.equal(app.columnIndexFromName("A"), 0);
  assert.equal(app.columnIndexFromName("AA"), 26);
}

function testExtractDealId() {
  assert.equal(app.extractDealId({ DEAL_ID: "42" }), 42);
  assert.equal(app.extractDealId({ URI: "/crm/deal/details/351/" }), 351);
  assert.equal(app.extractDealId({ url: "https://portal/crm/deal/show/19/" }), 19);
  assert.equal(app.extractDealId({ options: { ENTITY_VALUE_ID: "77" } }), 77);
  assert.equal(app.extractDealId({ PLACEMENT_OPTIONS: "{\"ID\":\"88\"}" }), 88);
  assert.equal(app.extractDealId({}), null);
}

function testGridStorageKey() {
  assert.equal(app.getGridStorageKey(null), "excel-tab-b24-grid-v1");
  assert.equal(app.getGridStorageKey("42"), "excel-tab-b24-grid-deal-v1-42");
  assert.equal(app.getGridStorageKey(351), "excel-tab-b24-grid-deal-v1-351");
  assert.equal(app.normalizeCategoryId("0"), 0);
  assert.equal(app.normalizeCategoryId("7"), 7);
  assert.equal(app.normalizeCategoryId("bad"), null);
  assert.equal(app.getFunnelStorageKey("0"), "excel-tab-b24-grid-funnel-v1-0");
  assert.equal(app.getFunnelStorageKey(7), "excel-tab-b24-grid-funnel-v1-7");
  assert.equal(app.getFunnelStorageKey(null), "excel-tab-b24-grid-funnel-v1-local");
  assert.equal(app.getSheetStorageKey("deal", 42, 7), "excel-tab-b24-grid-deal-v1-42");
  assert.equal(app.getSheetStorageKey("funnel", 42, 7), "excel-tab-b24-grid-funnel-v1-7");
}

function testReferenceFormatting() {
  assert.deepEqual(app.normalizeIdList([{ VALUE: "12" }, "0", 14]), [12, 14]);
  assert.equal(app.formatUser({ LAST_NAME: "Иванов", NAME: "Иван", SECOND_NAME: "Иванович" }), "Иванов Иван Иванович");
  assert.equal(app.formatUser({ ID: 5, EMAIL: "user@example.com" }), "user@example.com");
  assert.equal(app.formatContact({ NAME: "Анна", LAST_NAME: "Петрова" }), "Петрова Анна");
  assert.equal(app.formatContact({ ID: 8 }), "8");
  assert.equal(app.formatCompany({ TITLE: "Ромашка" }), "Ромашка");
  assert.equal(app.formatCompany({ ID: 9 }), "9");
  assert.equal(app.getDealStageEntityId(0), "DEAL_STAGE");
  assert.equal(app.getDealStageEntityId(3), "DEAL_STAGE_3");
  assert.equal(app.findCategoryName([], 0), "Общая воронка");
  assert.equal(app.findCategoryName([{ id: "3", name: "Service" }], 3), "Service");
  assert.equal(app.findStatusName([{ STATUS_ID: "NEW", NAME: "Новая" }], "NEW"), "Новая");
  assert.equal(app.findCategoryName([{ ID: "2", NAME: "Продажи" }], 2), "Продажи");
}

function testSelectionHelpers() {
  assert.equal(app.cellKey(2, 3), "2:3");
  assert.deepEqual(app.parseCellKey("2:3"), { rowIndex: 2, columnIndex: 3 });
  assert.deepEqual(app.getRangeCellKeys({ rowIndex: 1, columnIndex: 1 }, { rowIndex: 2, columnIndex: 2 }), [
    "1:1",
    "1:2",
    "2:1",
    "2:2",
  ]);
  assert.deepEqual(app.getFilledCellKeys([["", "x"], ["  ", "y"]]), ["0:1", "1:1"]);
  assert.deepEqual(app.getSelectedColumns(new Set(["2:3", "0:1", "1:3"])), [1, 3]);
  assert.ok(app.measureColumnWidth([["short"], ["long long text value"]], 0) > 132);
  assert.equal(app.measureColumnWidth([["x"]], 0), 90);
  assert.equal(app.measureColumnWidth([[Array.from({ length: 200 }, () => "x").join("")]], 0), 420);

  const autoFitGrid = [["manual value with enough length", ""], ["", ""]];
  const boundGrid = app.applyFieldBindings(autoFitGrid, { "1:1": "TITLE" }, [
    { id: "TITLE", value: "CRM field value with enough length" },
  ]).grid;
  const widths = app.getAutoFitColumnWidths(boundGrid, new Set(["0:0", "1:1"]));
  assert.ok(widths[0] > 132);
  assert.ok(widths[1] > 132);
  assert.equal(typeof widths[2], "undefined");
}

function testFieldBindingsAndExport() {
  const grid = [["old", ""], ["manual", ""]];
  const fields = [{ id: "TITLE", value: "fresh" }];
  const updated = app.applyFieldBindings(grid, { "0:0": "TITLE" }, fields);
  assert.equal(updated.changed, true);
  assert.deepEqual(updated.grid, [["fresh", ""], ["manual", ""]]);
  assert.equal(app.applyFieldBindings(grid, { "1:0": "MISSING" }, fields).changed, false);

  const exportGrid = app.getExportGrid([["a", ""], ["", "b"], ["", ""]]);
  assert.deepEqual(exportGrid, [["a", ""], ["", "b"]]);
  assert.deepEqual(app.getUsedGridBounds([[""], [""]]), { lastColumn: 0, lastRow: 0 });
  assert.equal(app.escapeHtml("<x & \"y\">"), "&lt;x &amp; &quot;y&quot;&gt;");
  assert.match(app.buildExcelHtml([["a", "b"], ["", ""]]), /<table><tr><td>a<\/td><td>b<\/td><\/tr><\/table>/);
  assert.equal(app.getExportFileName(7), "Excel Tab B24 deal 7.xls");
}

function testFormulaCells() {
  const grid = [
    ["2", "3", "=A1 + B1"],
    ["4", "5", "=A2 * B2"],
    ["", "", "=C1 + C2"],
  ];

  assert.deepEqual(app.parseFormulaReference("C3"), { columnIndex: 2, rowIndex: 2 });
  assert.equal(app.appendFormulaReference("=", "B1"), "=B1");
  assert.equal(app.appendFormulaReference("=B1", "C1"), "=B1 + C1");
  assert.equal(app.appendFormulaReference("=B1 * ", "C1"), "=B1 * C1");
  assert.equal(app.getCellDisplayValue(grid, 0, 2), "5");
  assert.equal(app.getCellDisplayValue(grid, 1, 2), "20");
  assert.equal(app.getCellDisplayValue(grid, 2, 2), "25");
  assert.equal(app.getCellDisplayValue([["=A1"]], 0, 0), "#ОШИБКА");
  assert.equal(app.shiftFormulaReferences("= E4 + B4", 1, 0), "= E5 + B5");
  assert.equal(app.shiftFormulaReferences("= E4 + B4", 0, 1), "= F4 + C4");
  assert.deepEqual(app.calculateSelectedCells(grid, new Set(["0:2", "1:2"]), "add"), { error: "", value: "25" });
  assert.match(app.buildExcelHtml([["2", "3", "=A1+B1"]]), /<td>5<\/td>/);
}

function testSavedFormulas() {
  assert.equal(app.sanitizeFormulaInput("=A1+С3+тест"), "=A1+3+");
  assert.equal(app.normalizeSavedFormula(" A1+B1 "), "=A1+B1");
  assert.equal(app.normalizeSavedFormula("=A1+B1"), "=A1+B1");
  assert.deepEqual(app.normalizeSavedFormulas([" A1+B1 ", "=A1+B1", "", " C1*2 "]), ["=A1+B1", "=C1*2"]);

  const added = app.addSavedFormula(["=A1+B1"], "C1*2");
  assert.deepEqual(added, { error: "", formula: "=C1*2", formulas: ["=A1+B1", "=C1*2"] });
  assert.deepEqual(app.removeSavedFormula(["=A1+B1", "=C1*2"], "A1+B1"), ["=C1*2"]);
  assert.equal(app.addSavedFormula([], " ").error, "Введите формулу");

  const applied = app.applyFormulaToGridCell([["", "2"]], 0, 0, "B1*2");
  assert.equal(applied.error, "");
  assert.equal(applied.changed, true);
  assert.deepEqual(applied.grid, [["=B1*2", "2"]]);
  assert.equal(app.applyFormulaToGridCell([[""]], 5, 0, "=A1").error, "Ячейка не найдена");

  withMockLocalStorage(() => {
    app.saveSavedFormulas(["A1+B1", "=A1+B1", "C1*2"], "formulas");
    assert.deepEqual(app.loadSavedFormulas("formulas"), ["=A1+B1", "=C1*2"]);
  });
}

function testClearCellSelectionState() {
  const cleared = app.clearCellSelectionState(
    {
      cellFormats: {
        "0:0": { fillColor: "#fff2cc", fontWeight: "700" },
        "1:1": { fillColor: "#d9ead3" },
      },
      fieldBindings: { "0:0": "TITLE", "1:1": "OPPORTUNITY" },
      grid: [
        ["=B1*2", "2"],
        ["keep", "remove"],
      ],
      wrappedCells: new Set(["0:0", "1:1"]),
    },
    new Set(["0:0", "1:1"])
  );

  assert.equal(cleared.changed, true);
  assert.deepEqual(cleared.grid, [["", "2"], ["keep", ""]]);
  assert.deepEqual(cleared.cellFormats, {});
  assert.deepEqual(cleared.fieldBindings, {});
  assert.deepEqual(Array.from(cleared.wrappedCells), []);
}

function testSheetStateStorage() {
  withMockLocalStorage((storage) => {
    storage.set("legacy", JSON.stringify([["a"]]));
    const legacy = app.loadSheetState("legacy");
    assert.deepEqual(legacy.grid, [["a"]]);
    assert.deepEqual(legacy.cellFormats, {});
    assert.deepEqual(legacy.columnWidths, []);
    assert.deepEqual(Array.from(legacy.wrappedCells), []);
    assert.deepEqual(legacy.fieldBindings, {});

    const wrappedCells = new Set(["0:0"]);
    app.saveSheetState(
      {
        cellFormats: { "0:0": { fillColor: "#fff2cc", fontWeight: "700" } },
        columnWidths: [180],
        fieldBindings: { "0:0": "TITLE" },
        grid: [["fresh"]],
        wrappedCells,
      },
      "state"
    );
    const saved = app.loadSheetState("state");
    assert.deepEqual(saved.grid, [["fresh"]]);
    assert.deepEqual(saved.cellFormats, { "0:0": { fillColor: "#fff2cc", fontWeight: "700" } });
    assert.deepEqual(saved.columnWidths, [180]);
    assert.deepEqual(saved.fieldBindings, { "0:0": "TITLE" });
    assert.deepEqual(Array.from(saved.wrappedCells), ["0:0"]);
  });
}

function testNormalizeAndFormatFields() {
  const fields = {
    TITLE: { title: "Название" },
    OPPORTUNITY: { formLabel: "Сумма" },
  };
  const deal = {
    TITLE: "Новая сделка",
    OPPORTUNITY: 1500,
  };

  const normalized = app.normalizeFields(fields, deal);
  assert.equal(normalized.length, 2);
  assert.equal(app.formatDealFieldValue(["a", "b"]), "a, b");
  assert.equal(app.formatDealFieldValue({ VALUE: "x" }), "x");
  assert.equal(app.formatDealFieldValue(null), "");
}

function testCalculationsAndCellStyles() {
  const grid = [["10", "2,5", "0"], [" 3 000 ", "text", ""]];
  const selected = new Set(["0:0", "0:1", "1:0"]);

  assert.equal(app.parseCellNumber("1 200,50 руб."), 1200.5);
  assert.equal(app.parseCellNumber("text"), null);
  assert.deepEqual(app.getSortedCellKeys(new Set(["2:1", "0:2", "0:1"])), ["0:1", "0:2", "2:1"]);
  assert.deepEqual(app.calculateSelectedCells(grid, selected, "add"), { error: "", value: "3012,5" });
  assert.deepEqual(app.calculateSelectedCells(grid, selected, "subtract"), { error: "", value: "-2992,5" });
  assert.deepEqual(app.calculateSelectedCells([["2"], ["3"], ["4"]], new Set(["0:0", "1:0", "2:0"]), "multiply"), {
    error: "",
    value: "24",
  });
  assert.deepEqual(app.calculateSelectedCells([["12"], ["3"]], new Set(["0:0", "1:0"]), "divide"), {
    error: "",
    value: "4",
  });
  assert.equal(app.calculateSelectedCells([["12"], ["0"]], new Set(["0:0", "1:0"]), "divide").error, "Деление на ноль");
  assert.deepEqual(app.normalizeCellFormat({ fillColor: "bad;color:red", fontWeight: "999" }), {
    fillColor: "",
    fontWeight: "",
  });
  assert.match(
    app.buildExcelHtml([["styled"]], { "0:0": { fillColor: "#fff2cc", fontWeight: "700" } }),
    /<td style="background-color:#fff2cc;font-weight:700">styled<\/td>/
  );
}

testCreateGridDefaults();
testAddRowAndColumn();
testColumnName();
testExtractDealId();
testGridStorageKey();
testReferenceFormatting();
testSelectionHelpers();
testFieldBindingsAndExport();
testFormulaCells();
testSavedFormulas();
testClearCellSelectionState();
testSheetStateStorage();
testNormalizeAndFormatFields();
testCalculationsAndCellStyles();

console.log("app tests passed");
