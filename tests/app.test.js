"use strict";

const assert = require("node:assert/strict");
const app = require("../app.js");

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
  assert.equal(app.extractDealId({}), null);
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
testNormalizeAndFormatFields();

console.log("app tests passed");
