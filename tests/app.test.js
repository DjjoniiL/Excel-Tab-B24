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
  assert.equal(grid.length, 12);
  assert.equal(grid[0].length, 8);
  assert.equal(grid[11][7], "");
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

function testSheetStateStorage() {
  withMockLocalStorage((storage) => {
    storage.set("legacy", JSON.stringify([["a"]]));
    const legacy = app.loadSheetState("legacy");
    assert.deepEqual(legacy.grid, [["a"]]);
    assert.deepEqual(legacy.columnWidths, []);
    assert.deepEqual(Array.from(legacy.wrappedCells), []);
    assert.deepEqual(legacy.fieldBindings, {});

    const wrappedCells = new Set(["0:0"]);
    app.saveSheetState(
      {
        columnWidths: [180],
        fieldBindings: { "0:0": "TITLE" },
        grid: [["fresh"]],
        wrappedCells,
      },
      "state"
    );
    const saved = app.loadSheetState("state");
    assert.deepEqual(saved.grid, [["fresh"]]);
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

testCreateGridDefaults();
testAddRowAndColumn();
testColumnName();
testExtractDealId();
testGridStorageKey();
testReferenceFormatting();
testSelectionHelpers();
testFieldBindingsAndExport();
testSheetStateStorage();
testNormalizeAndFormatFields();

console.log("app tests passed");
