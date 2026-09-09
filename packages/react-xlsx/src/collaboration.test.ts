import test from "node:test";
import assert from "node:assert/strict";
import { captureJournalDelta, mergeCellState, normalizeCollaborationStyle, transformCellAddress, transformFormulaReferences } from "./collaboration.ts";

const state = (value: unknown, style: unknown = {}) => ({ value, formula: null, style });

test("delta capture collapses local edits with original before state and excludes formula caches", () => {
  const entries = [
    { revision: 1, worksheet: "Data", cell: { row: 0, col: 0 }, before: state(12), after: state(20) },
    { revision: 2, worksheet: "Data", cell: { row: 0, col: 0 }, before: state(20), after: { ...state(42), formula: "=SUM(B1:B2)" } }
  ];
  assert.deepEqual(captureJournalDelta(entries, 0, 2, 0, 0, 0), {
    kind: "cell_delta", baseRevision: 0, revision: 2, structureRevision: 0,
    changes: [{ worksheet: "Data", address: "A1", before: { kind: "value", value: 12 }, after: { kind: "formula", formula: "SUM(B1:B2)" } }]
  });
  assert.equal(captureJournalDelta(entries, 0, 2, 0, 1, 0), null);
  assert.equal(captureJournalDelta([{ ...entries[0], after: state(20, { font: { bold: true } }) }], 0, 1, 0, 0, 0), null);
});

test("cell rebase retains disjoint content and style properties, conflicts only on competing properties", () => {
  const before = state(12, { font: { bold: false, color: "red" } });
  const local = state(12, { font: { bold: true, color: "red" } });
  const remote = state(30, { font: { bold: false, color: "blue" } });
  assert.deepEqual(mergeCellState(before, local, remote), { state: state(30, { font: { bold: true, color: "blue" } }), conflicts: [] });
  assert.deepEqual(mergeCellState(state(12), state(20), state(30)).conflicts, ["content"]);
  assert.deepEqual(mergeCellState(state(12), state(20), state(20)).conflicts, []);
});

test("implicit default styles merge by property and an undone style does not force a snapshot", () => {
  const defaults = normalizeCollaborationStyle(null) as { font: { bold: boolean; italic: boolean } };
  const bold = { ...defaults, font: { ...defaults.font, bold: true } };
  const italic = { ...defaults, font: { ...defaults.font, italic: true } };
  const merged = mergeCellState(state(12, null), state(12, bold), state(12, italic));
  assert.deepEqual(merged.conflicts, []);
  assert.deepEqual((merged.state.style as typeof defaults).font, { ...defaults.font, bold: true, italic: true });
  assert.deepEqual(captureJournalDelta([{ revision: 1, worksheet: "Data", cell: { row: 0, col: 0 }, before: state(12, null), after: state(12, defaults) }], 0, 1, 0, 0, 0)?.changes, []);
});

test("known structure transforms shift addresses and refuse deleted targets", () => {
  const insert = { kind: "insert_rows", worksheet: "Data", index: 1, count: 2, operationId: "op1" } as const;
  assert.deepEqual(transformCellAddress({ row: 3, col: 2 }, "Data", [insert]), { row: 5, col: 2 });
  assert.deepEqual(transformCellAddress({ row: 3, col: 2 }, "Other", [insert]), { row: 3, col: 2 });
  assert.deepEqual(transformCellAddress({ row: 3, col: 2 }, "data", [insert]), { row: 5, col: 2 });
  assert.equal(transformCellAddress({ row: 1, col: 2 }, "Data", [{ ...insert, kind: "delete_rows" }]), null);
  assert.throws(() => transformCellAddress({ row: 0, col: 0 }, "Data", [{ ...insert, kind: "unknown" } as never]), /unsupported_structure_transform/);
});

test("structure transforms rewrite explicit references, preserve quoted text, and reject unsupported coordinates", () => {
  const transform = { kind: "insert_rows", worksheet: "Data", index: 1, count: 2, operationId: "insert" } as const;
  assert.equal(transformFormulaReferences('SUM($A$1:A4)+"A2"+Other!B2+\'Data\'!C2', "Data", [transform]), 'SUM($A$1:A6)+"A2"+Other!B2+\'Data\'!C4');
  assert.equal(transformFormulaReferences('SUM(A1:A4)+B2', "Data", [{ ...transform, kind: "delete_rows" }]), 'SUM(A1:A2)+#REF!');
  assert.throws(() => transformFormulaReferences('SUM(A:A)', "Data", [transform]), /unsupported_formula_transform/);
});

test("formula transforms retain function names, case-insensitive sheet refs, and errors through sequential operations", () => {
  const first = { kind: "delete_rows", worksheet: "Data", index: 1, count: 1, operationId: "one" } as const;
  const second = { kind: "insert_columns", worksheet: "Data", index: 1, count: 1, operationId: "two" } as const;
  assert.equal(transformFormulaReferences('LOG10 (100)+data!B2+#REF!+SUM(B1:C3)', 'Data', [first, second]), 'LOG10 (100)+#REF!+#REF!+SUM(C1:D2)');
  assert.throws(() => transformFormulaReferences('B1#', 'Data', [first]), /unsupported_formula_transform/);
  assert.throws(() => transformFormulaReferences('Sheet1:Sheet2!A1', 'Data', [first]), /unsupported_formula_transform/);
  for (const formula of ['INDIRECT("B2")', 'OFFSET(B2,1,1)', 'SUM(MyNamedRange)', 'SUM(B3:A1)', 'SUM("broken)']) {
    assert.throws(() => transformFormulaReferences(formula, 'Data', [first]), /unsupported_formula_transform/);
  }
});
