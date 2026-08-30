import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";

const installRoot = resolve(process.argv[2] ?? ".");
const moduleRoot = join(installRoot, "node_modules", "@dukelib", "sheets-wasm");
const moduleUrl = pathToFileURL(join(moduleRoot, "duke_sheets_wasm.js")).href;
const { default: init, Workbook } = await import(moduleUrl);

await init({
  module_or_path: readFileSync(join(moduleRoot, "duke_sheets_wasm_bg.wasm"))
});

const workbook = new Workbook();
workbook.renameSheet(0, "Data");
workbook.addSheet("Summary");
workbook.getSheet(0).setCell("A1", 41);
workbook.getSheet(1).setFormula("A1", "=Data!A1+1");

workbook.renameSheet(0, "Source");
assert.equal(workbook.getSheet(1).getFormulaAt(0, 0), "='Source'!A1+1");

workbook.moveSheet(0, 1);
const reopened = Workbook.fromBytes(workbook.saveXlsxBytes());
assert.deepEqual(reopened.sheetNames, ["Summary", "Source"]);
assert.equal(reopened.getSheet(0).getFormulaAt(0, 0), "='Source'!A1+1");
reopened.calculate();
assert.equal(reopened.getSheet(0).getCalculatedValue("A1").asNumber(), 42);
assert.equal(reopened.getSheet(1).getCell("A1").asNumber(), 41);

console.log(JSON.stringify({
  formula: reopened.getSheet(0).getFormulaAt(0, 0),
  sheets: reopened.sheetNames,
  value: reopened.getSheet(0).getCalculatedValue("A1").asNumber()
}));
