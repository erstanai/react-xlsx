/**
 * Parses the plain-text clipboard shape emitted by Excel and compatible grids.
 * Rows are newline-delimited and cells are tab-delimited. A single terminal
 * newline is ignored because spreadsheet copy operations commonly include it.
 */
export function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const rows = normalized.split("\n");

  if (rows.length > 1 && rows[rows.length - 1] === "") {
    rows.pop();
  }

  return rows.map((row) => row.split("\t"));
}
