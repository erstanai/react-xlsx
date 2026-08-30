import assert from "node:assert/strict";
import test from "node:test";
import { offsetFormulaReferences } from "./formula-references.ts";

test("shifts relative references by the delta", () => {
  assert.equal(offsetFormulaReferences("B11", 1, 0), "B12");
  assert.equal(offsetFormulaReferences("B11", 0, 2), "D11");
  assert.equal(offsetFormulaReferences("B11*2+C3", 3, 1), "C14*2+D6");
});

test("keeps $-anchored components fixed", () => {
  assert.equal(offsetFormulaReferences("$B11", 2, 2), "$B13");
  assert.equal(offsetFormulaReferences("B$11", 2, 2), "D$11");
  assert.equal(offsetFormulaReferences("$B$11", 5, 5), "$B$11");
});

test("adjusts both ends of a range independently", () => {
  assert.equal(offsetFormulaReferences("SUM(B2:B10)", 1, 1), "SUM(C3:C11)");
  assert.equal(offsetFormulaReferences("SUM($B$2:B10)", 1, 1), "SUM($B$2:C11)");
});

test("returns #REF! when a reference would leave the sheet", () => {
  assert.equal(offsetFormulaReferences("A1", -1, 0), "#REF!");
  assert.equal(offsetFormulaReferences("A1", 0, -1), "#REF!");
  assert.equal(offsetFormulaReferences("A1+B2", -1, 0), "#REF!+B1");
  assert.equal(offsetFormulaReferences("XFD1", 0, 1), "#REF!");
  assert.equal(offsetFormulaReferences("A1048576", 1, 0), "#REF!");
});

test("never touches string literals", () => {
  assert.equal(offsetFormulaReferences('CONCAT("B11",B11)', 1, 0), 'CONCAT("B11",B12)');
  assert.equal(offsetFormulaReferences('IF(A1="x""B2""",B2,C3)', 1, 0), 'IF(A2="x""B2""",B3,C4)');
});

test("preserves quoted sheet prefixes and adjusts the reference after them", () => {
  assert.equal(offsetFormulaReferences("'Q1 Data'!B2", 1, 0), "'Q1 Data'!B3");
  assert.equal(offsetFormulaReferences("'X5'!A1+A1", 0, 1), "'X5'!B1+B1");
});

test("does not mistake functions or names for references", () => {
  assert.equal(offsetFormulaReferences("LOG10(A1)", 1, 1), "LOG10(B2)");
  assert.equal(offsetFormulaReferences("ATAN2(A1,B2)", 1, 0), "ATAN2(A2,B3)");
  assert.equal(offsetFormulaReferences("MY.NAME2+A1", 1, 0), "MY.NAME2+A2");
  assert.equal(offsetFormulaReferences("_B2+B2", 1, 0), "_B2+B3");
});

test("leaves unquoted sheet names that resemble references untouched", () => {
  assert.equal(offsetFormulaReferences("X5!A1", 1, 0), "X5!A2");
});

test("leaves whole-column and whole-row references unchanged", () => {
  assert.equal(offsetFormulaReferences("SUM(B:B)", 5, 5), "SUM(B:B)");
  assert.equal(offsetFormulaReferences("SUM(2:2)", 5, 5), "SUM(2:2)");
});

test("is the identity for a zero delta", () => {
  const formula = 'IF($A$1>0,SUM(B2:B10),"none")';
  assert.equal(offsetFormulaReferences(formula, 0, 0), formula);
});

test("preserves lowercase reference spelling", () => {
  assert.equal(offsetFormulaReferences("b11+B11", 1, 1), "c12+C12");
});

test("handles column letter rollover", () => {
  assert.equal(offsetFormulaReferences("Z1", 0, 1), "AA1");
  assert.equal(offsetFormulaReferences("AA1", 0, -1), "Z1");
});
