import assert from "node:assert/strict";
import test from "node:test";
import { computePasteStampOffsets, parseClipboardText } from "./clipboard.ts";

test("parses an Excel-style multi-cell TSV payload", () => {
  assert.deepEqual(
    parseClipboardText("North\t10\ttrue\r\nSouth\t20\tfalse\r\n"),
    [
      ["North", "10", "true"],
      ["South", "20", "false"],
    ],
  );
});

test("preserves empty cells at the middle and edge of a pasted range", () => {
  assert.deepEqual(
    parseClipboardText("A\t\tC\n\tB\t"),
    [
      ["A", "", "C"],
      ["", "B", ""],
    ],
  );
});

test("fills the whole selection when it is an exact multiple of the clipboard block", () => {
  assert.deepEqual(
    computePasteStampOffsets(3, 1, 1, 1),
    [
      { colOffset: 0, rowOffset: 0 },
      { colOffset: 0, rowOffset: 1 },
      { colOffset: 0, rowOffset: 2 },
    ],
  );
  assert.deepEqual(
    computePasteStampOffsets(4, 2, 2, 1),
    [
      { colOffset: 0, rowOffset: 0 },
      { colOffset: 1, rowOffset: 0 },
      { colOffset: 0, rowOffset: 2 },
      { colOffset: 1, rowOffset: 2 },
    ],
  );
});

test("pastes a single stamp for 1x1 selections and non-multiple shapes", () => {
  assert.deepEqual(computePasteStampOffsets(1, 1, 1, 1), [{ colOffset: 0, rowOffset: 0 }]);
  assert.deepEqual(computePasteStampOffsets(5, 1, 2, 1), [{ colOffset: 0, rowOffset: 0 }]);
  assert.deepEqual(computePasteStampOffsets(2, 3, 2, 2), [{ colOffset: 0, rowOffset: 0 }]);
  assert.deepEqual(computePasteStampOffsets(0, 0, 1, 1), [{ colOffset: 0, rowOffset: 0 }]);
});
