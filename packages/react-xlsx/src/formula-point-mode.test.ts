import assert from "node:assert/strict";
import test from "node:test";
import { applyFormulaPointInsert } from "./formula-point-mode.ts";

test("appends a reference when the formula ends with an operator", () => {
  assert.deepEqual(applyFormulaPointInsert("=", null, "B11"), {
    next: "=B11",
    span: { start: 1, end: 4 },
    caret: 4,
  });
  assert.deepEqual(applyFormulaPointInsert("=A1+", null, "B2"), {
    next: "=A1+B2",
    span: { start: 4, end: 6 },
    caret: 6,
  });
  assert.deepEqual(applyFormulaPointInsert("=SUM(", null, "C3"), {
    next: "=SUM(C3",
    span: { start: 5, end: 7 },
    caret: 7,
  });
  assert.deepEqual(applyFormulaPointInsert("=SUM(A1,", null, "D4"), {
    next: "=SUM(A1,D4",
    span: { start: 8, end: 10 },
    caret: 10,
  });
  assert.deepEqual(applyFormulaPointInsert("=A1:", null, "A9"), {
    next: "=A1:A9",
    span: { start: 4, end: 6 },
    caret: 6,
  });
});

test("a second click replaces the pointed reference", () => {
  const first = applyFormulaPointInsert("=A1+", null, "B2");
  assert.ok(first);
  const second = applyFormulaPointInsert(first.next, first.span, "C7");
  assert.deepEqual(second, {
    next: "=A1+C7",
    span: { start: 4, end: 6 },
    caret: 6,
  });
});

test("declines when the formula cannot accept a reference", () => {
  assert.equal(applyFormulaPointInsert("=A1", null, "B2"), null);
  assert.equal(applyFormulaPointInsert("=SUM(A1)", null, "B2"), null);
  assert.equal(applyFormulaPointInsert('="text"', null, "B2"), null);
});

test("declines for non-formula input and empty references", () => {
  assert.equal(applyFormulaPointInsert("hello", null, "B2"), null);
  assert.equal(applyFormulaPointInsert("", null, "B2"), null);
  assert.equal(applyFormulaPointInsert("=A1+", null, ""), null);
});

test("ignores a stale span that no longer fits the value", () => {
  // The span survives only while the value is untouched; a shrunken value
  // (programmatic edit) invalidates it and the tail rule decides instead.
  assert.equal(applyFormulaPointInsert("=A1", { start: 2, end: 9 }, "B2"), null);
  assert.deepEqual(applyFormulaPointInsert("=A1+", { start: 2, end: 9 }, "B2"), {
    next: "=A1+B2",
    span: { start: 4, end: 6 },
    caret: 6,
  });
});

test("replaces a pointed reference that sits before a suffix", () => {
  // Point, then the user typed an operator after the ref, then pointed again:
  // the span was cleared by typing, so the new ref appends after the operator.
  const first = applyFormulaPointInsert("=", null, "B2");
  assert.ok(first);
  const afterTyping = `${first.next}*2`;
  assert.equal(applyFormulaPointInsert(afterTyping, null, "C3"), null);
  const withOperator = `${afterTyping}+`;
  assert.deepEqual(applyFormulaPointInsert(withOperator, null, "C3"), {
    next: "=B2*2+C3",
    span: { start: 6, end: 8 },
    caret: 8,
  });
});
