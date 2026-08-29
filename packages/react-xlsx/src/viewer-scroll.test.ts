import assert from "node:assert/strict";
import test from "node:test";
import { buildInitialViewportKey } from "./viewer-scroll.ts";

test("used-range changes do not replay initial viewport positioning", () => {
  const before = buildInitialViewportKey({
    displayFileName: "forecast.xlsx",
    activeSheetIndex: 0,
    activeSheet: { workbookSheetIndex: 0, name: "Forecast", minUsedRow: -1, minUsedCol: -1 },
    isWorkerBacked: false,
  });
  const afterTyping = buildInitialViewportKey({
    displayFileName: "forecast.xlsx",
    activeSheetIndex: 0,
    activeSheet: { workbookSheetIndex: 0, name: "Forecast", minUsedRow: 24, minUsedCol: 18 },
    isWorkerBacked: false,
  });

  assert.equal(afterTyping, before);
});

test("workbook and tab identity changes still initialize the viewport", () => {
  const firstSheet = buildInitialViewportKey({
    displayFileName: "forecast.xlsx",
    activeSheetIndex: 0,
    activeSheet: { workbookSheetIndex: 0, name: "Forecast", minUsedRow: 0, minUsedCol: 0 },
    isWorkerBacked: false,
  });
  const nextSheet = buildInitialViewportKey({
    displayFileName: "forecast.xlsx",
    activeSheetIndex: 1,
    activeSheet: { workbookSheetIndex: 1, name: "Assumptions", minUsedRow: 0, minUsedCol: 0 },
    isWorkerBacked: false,
  });
  const nextWorkbook = buildInitialViewportKey({
    displayFileName: "actuals.xlsx",
    activeSheetIndex: 0,
    activeSheet: { workbookSheetIndex: 0, name: "Forecast", minUsedRow: 0, minUsedCol: 0 },
    isWorkerBacked: false,
  });

  assert.notEqual(nextSheet, firstSheet);
  assert.notEqual(nextWorkbook, firstSheet);
});
