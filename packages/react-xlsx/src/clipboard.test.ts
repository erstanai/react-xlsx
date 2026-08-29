import assert from "node:assert/strict";
import test from "node:test";
import { parseClipboardText } from "./clipboard.ts";

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
