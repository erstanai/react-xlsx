import assert from "node:assert/strict";
import test from "node:test";
import {
  moveArrayEntry,
  movedArrayIndex,
  moveWorkbookIndexedEntries,
  moveWorkbookIndexedGroups
} from "./worksheet-order.ts";

test("moveArrayEntry and movedArrayIndex preserve identity in both directions", () => {
  assert.deepEqual(moveArrayEntry(["A", "B", "C"], 0, 2), ["B", "C", "A"]);
  assert.deepEqual([0, 1, 2].map((index) => movedArrayIndex(index, 0, 2)), [2, 0, 1]);
  assert.deepEqual(moveArrayEntry(["A", "B", "C"], 2, 0), ["C", "A", "B"]);
  assert.deepEqual([0, 1, 2].map((index) => movedArrayIndex(index, 2, 0)), [1, 2, 0]);
});

test("worksheet-indexed asset groups follow their worksheet and receive new indexes", () => {
  const moved = moveWorkbookIndexedGroups([
    [{ id: "image-a", workbookSheetIndex: 0 }],
    [{ id: "image-b", workbookSheetIndex: 1 }],
    [{ id: "image-c", workbookSheetIndex: 2 }]
  ], 0, 2);

  assert.deepEqual(moved, [
    [{ id: "image-b", workbookSheetIndex: 0 }],
    [{ id: "image-c", workbookSheetIndex: 1 }],
    [{ id: "image-a", workbookSheetIndex: 2 }]
  ]);
});

test("nullable worksheet origins follow the same reordered association", () => {
  const moved = moveWorkbookIndexedEntries([
    { workbookSheetIndex: 0, path: "sheet-a.xml" },
    null,
    { workbookSheetIndex: 2, path: "sheet-c.xml" }
  ], 2, 0);

  assert.deepEqual(moved, [
    { workbookSheetIndex: 0, path: "sheet-c.xml" },
    { workbookSheetIndex: 1, path: "sheet-a.xml" },
    null
  ]);
});

test("invalid moves fail without returning a partially reordered array", () => {
  assert.throws(() => moveArrayEntry(["A"], -1, 0), RangeError);
  assert.throws(() => moveArrayEntry(["A"], 0, 1), RangeError);
});

test("workbook indexes remain authoritative when visible tabs are non-contiguous", () => {
  const visibleTabs = [
    { name: "Visible A", workbookSheetIndex: 0 },
    { name: "Visible C", workbookSheetIndex: 2 }
  ];
  const from = visibleTabs[1]?.workbookSheetIndex as number;
  const to = visibleTabs[0]?.workbookSheetIndex as number;

  assert.deepEqual(
    moveArrayEntry(["visible-a", "hidden-b", "visible-c"], from, to),
    ["visible-c", "visible-a", "hidden-b"]
  );
});
