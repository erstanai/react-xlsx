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

/**
 * Excel paste-tiling rule: when the selected range spans multiple cells and
 * its dimensions are exact multiples of the clipboard block, the block is
 * stamped repeatedly to fill the selection (copying one cell into a
 * highlighted column fills the whole column). Any other shape pastes a
 * single stamp at the anchor.
 */
export function computePasteStampOffsets(
  selectionRows: number,
  selectionCols: number,
  clipRows: number,
  clipCols: number
): Array<{ colOffset: number; rowOffset: number }> {
  const single = [{ colOffset: 0, rowOffset: 0 }];
  if (clipRows < 1 || clipCols < 1) {
    return single;
  }
  if (selectionRows * selectionCols <= 1) {
    return single;
  }
  if (selectionRows % clipRows !== 0 || selectionCols % clipCols !== 0) {
    return single;
  }

  const stamps: Array<{ colOffset: number; rowOffset: number }> = [];
  for (let row = 0; row < selectionRows / clipRows; row += 1) {
    for (let col = 0; col < selectionCols / clipCols; col += 1) {
      stamps.push({ colOffset: col * clipCols, rowOffset: row * clipRows });
    }
  }
  return stamps;
}
