export type XlsxCellContent = { kind: "value"; value: string | number | boolean | null } | { kind: "formula"; formula: string };
export type XlsxCellChange = { worksheet: string; address: string; before: XlsxCellContent; after: XlsxCellContent };
export type XlsxCellDeltaCapture = { kind: "cell_delta"; baseRevision: number; revision: number; structureRevision: number; changes: XlsxCellChange[] };
export type XlsxStructureTransform = { kind: "insert_rows" | "delete_rows" | "insert_columns" | "delete_columns"; worksheet: string; index: number; count: number; operationId: string };
export type XlsxCellConflict = { worksheet: string; address: string; property: string; reason: string };
export type XlsxReconcileResult = { applied: boolean; revision: number; cleanRevision: number; structureRevision: number; dirty: boolean; conflicts?: XlsxCellConflict[] };
export type XlsxReconcileOptions = { baseRevision: number; acknowledgedRevision?: number; expectedRevision?: number; canApply?: () => boolean; structureUnchanged?: boolean };
export type XlsxJournalCellState = { value: unknown; formula: string | null; style: unknown };
export type XlsxJournalEntry = { revision: number; worksheet: string; cell: { row: number; col: number }; before: XlsxJournalCellState; after: XlsxJournalCellState };

// Duke 0.1.23 returns null for an unstyled cell, but a style patch materializes
// these defaults. Treat both representations identically when comparing
// properties; restoring null must also clear a previously materialized style.
const DEFAULT_CELL_STYLE = {
  font: { name: "Calibri", size: 11, bold: false, italic: false, underline: "none", strikethrough: false, color: { colorType: "auto" }, verticalAlign: "baseline" },
  fill: { fillType: "none" }, border: { diagonalDirection: "none" },
  alignment: { horizontal: "general", vertical: "bottom", wrapText: false, shrinkToFit: false, indent: 0, rotation: 0, readingOrder: "contextDependent" },
  numberFormat: { formatType: "general", formatString: "General", isDateFormat: false }, protection: { locked: true, hidden: false }
};
export function normalizeCollaborationStyle(style: unknown): unknown { return style ?? cloneCollaborationValue(DEFAULT_CELL_STYLE); }
export function collaborationStyleEqual(left: unknown, right: unknown): boolean { return collaborationEqual(normalizeCollaborationStyle(left), normalizeCollaborationStyle(right)); }

export function cloneCollaborationValue<T>(value: T): T { return structuredClone(value); }
export function collaborationEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (!left || !right || typeof left !== "object" || typeof right !== "object") return false;
  if (Array.isArray(left) !== Array.isArray(right)) return false;
  const a = Object.keys(left).sort(), b = Object.keys(right).sort();
  return a.length === b.length && a.every((key, index) => key === b[index] && collaborationEqual((left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key]));
}
export function cellStateContent(state: XlsxJournalCellState): XlsxCellContent | null {
  if (state.formula) return { kind: "formula", formula: state.formula.replace(/^=/, "") };
  const value = state.value ?? null;
  if (value === null || typeof value === "string" || typeof value === "boolean" || (typeof value === "number" && Number.isFinite(value))) return { kind: "value", value };
  return null;
}
export function isCellContent(content: XlsxCellContent): boolean {
  if (!content || typeof content !== "object") return false;
  if (content.kind === "formula") return typeof content.formula === "string" && content.formula.length > 0 && !content.formula.startsWith("=");
  if (content.kind !== "value") return false;
  return content.value === null || typeof content.value === "string" || typeof content.value === "boolean" || (typeof content.value === "number" && Number.isFinite(content.value));
}
export function stateWithContent(state: XlsxJournalCellState, content: XlsxCellContent): XlsxJournalCellState {
  return { ...cloneCollaborationValue(state), formula: content.kind === "formula" ? content.formula : null, value: content.kind === "value" ? content.value : null };
}
export function addressFromCell(cell: { row: number; col: number }): string {
  let column = cell.col + 1, name = "";
  while (column > 0) { name = String.fromCharCode(65 + (column - 1) % 26) + name; column = Math.floor((column - 1) / 26); }
  return `${name}${cell.row + 1}`;
}
export function cellFromAddress(address: string): { row: number; col: number } {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address);
  if (!match) throw new Error("invalid_cell_delta_address");
  const col = [...match[1]].reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  if (row >= 1_048_576 || col >= 16_384) throw new Error("invalid_cell_delta_address");
  return { row, col };
}
export function journalCellKey(worksheet: string, cell: { row: number; col: number }): string { return JSON.stringify([worksheet, cell.row, cell.col]); }
export function aggregateJournal(entries: XlsxJournalEntry[]): XlsxJournalEntry[] {
  const cells = new Map<string, XlsxJournalEntry>();
  for (const entry of entries) {
    const key = journalCellKey(entry.worksheet, entry.cell);
    const previous = cells.get(key);
    cells.set(key, previous ? { ...previous, revision: entry.revision, after: cloneCollaborationValue(entry.after) } : cloneCollaborationValue(entry));
  }
  return [...cells.values()].filter((entry) => !collaborationEqual(cellStateContent(entry.before), cellStateContent(entry.after)) || !collaborationStyleEqual(entry.before.style, entry.after.style));
}
export function captureJournalDelta(entries: XlsxJournalEntry[], baseRevision: number, revision: number, baselineStructure: number, structureRevision: number, unsupportedRevision: number): XlsxCellDeltaCapture | null {
  if (baselineStructure !== structureRevision || unsupportedRevision > baseRevision || !Number.isSafeInteger(baseRevision) || baseRevision > revision) return null;
  const changes: XlsxCellChange[] = [];
  for (const entry of aggregateJournal(entries.filter((entry) => entry.revision > baseRevision))) {
    const before = cellStateContent(entry.before), after = cellStateContent(entry.after);
    if (!before || !after || !collaborationStyleEqual(entry.before.style, entry.after.style)) return null;
    if (!collaborationEqual(before, after)) changes.push({ worksheet: entry.worksheet, address: addressFromCell(entry.cell), before, after });
    if (changes.length > 1000) return null;
  }
  return { kind: "cell_delta", baseRevision, revision, structureRevision, changes };
}

/** Apply local properties onto a new server base without overwriting competing edits. */
export function mergeCellState(before: XlsxJournalCellState, local: XlsxJournalCellState, remote: XlsxJournalCellState): { state: XlsxJournalCellState; conflicts: string[] } {
  const conflicts: string[] = [];
  const merge = (base: unknown, ours: unknown, theirs: unknown, path: string): unknown => {
    if (collaborationEqual(base, ours)) return cloneCollaborationValue(theirs);
    if (collaborationEqual(base, theirs) || collaborationEqual(ours, theirs)) return cloneCollaborationValue(ours);
    if ([base, ours, theirs].every((value) => value && typeof value === "object" && !Array.isArray(value))) {
      const result: Record<string, unknown> = {};
      for (const key of new Set([...Object.keys(base as object), ...Object.keys(ours as object), ...Object.keys(theirs as object)])) {
        const value = merge((base as Record<string, unknown>)[key], (ours as Record<string, unknown>)[key], (theirs as Record<string, unknown>)[key], `${path}.${key}`);
        if (value !== undefined) result[key] = value;
      }
      return result;
    }
    conflicts.push(path);
    return cloneCollaborationValue(theirs);
  };
  const baseContent = cellStateContent(before), localContent = cellStateContent(local), remoteContent = cellStateContent(remote);
  if (!baseContent || !localContent || !remoteContent) conflicts.push("content");
  const content = collaborationEqual(baseContent, localContent) ? remoteContent
    : collaborationEqual(baseContent, remoteContent) || collaborationEqual(localContent, remoteContent) ? localContent
    : (conflicts.push("content"), remoteContent);
  const style = merge(normalizeCollaborationStyle(before.style), normalizeCollaborationStyle(local.style), normalizeCollaborationStyle(remote.style), "style");
  return { state: { ...stateWithContent(remote, content ?? { kind: "value", value: null }), style }, conflicts: [...new Set(conflicts)] };
}

export function transformCellAddress(cell: { row: number; col: number }, worksheet: string, transforms: XlsxStructureTransform[]): { row: number; col: number } | null {
  if (!Array.isArray(transforms) || transforms.length > 256) throw new Error("unsupported_structure_transform");
  const result = { ...cell };
  const ids = new Set<string>();
  for (const transform of transforms) {
    if (!["insert_rows", "delete_rows", "insert_columns", "delete_columns"].includes(transform.kind)
      || typeof transform.worksheet !== "string" || !transform.worksheet.trim() || transform.worksheet.length > 31 || /[\x00-\x1f\[\]:*?/\\]/.test(transform.worksheet)
      || typeof transform.operationId !== "string" || !transform.operationId.trim() || transform.operationId.length > 200 || ids.has(transform.operationId)
      || !Number.isSafeInteger(transform.index) || transform.index < 0 || !Number.isSafeInteger(transform.count) || transform.count < 1) throw new Error("unsupported_structure_transform");
    ids.add(transform.operationId);
    const axis = transform.kind.endsWith("rows") ? "row" : "col";
    const limit = axis === "row" ? 1_048_576 : 16_384;
    if (transform.index + transform.count > limit) throw new Error("unsupported_structure_transform");
    if (transform.worksheet.toLowerCase() !== worksheet.toLowerCase()) continue;
    if (transform.kind.startsWith("delete") && result[axis] >= transform.index && result[axis] < transform.index + transform.count) return null;
    if (result[axis] >= transform.index) result[axis] += transform.kind.startsWith("insert") ? transform.count : -transform.count;
    if (result[axis] >= limit) return null;
  }
  return result;
}

// Keep the supported formula subset aligned with the Erstan structure endpoint.
const FUNCTIONS = new Set(('SUM SUMIF SUMIFS SUMPRODUCT AVERAGE AVERAGEIF AVERAGEIFS MIN MAX COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS IF IFS IFERROR IFNA AND OR NOT XOR TRUE FALSE ABS SIGN ROUND ROUNDUP ROUNDDOWN MROUND CEILING FLOOR INT TRUNC MOD POWER SQRT EXP LN LOG LOG10 SIN COS TAN ASIN ACOS ATAN ATAN2 PI DEGREES RADIANS CONCAT CONCATENATE TEXTJOIN LEFT RIGHT MID LEN LOWER UPPER PROPER TRIM CLEAN SUBSTITUTE REPLACE FIND SEARCH EXACT VALUE TEXT DATE TIME DAY MONTH YEAR HOUR MINUTE SECOND WEEKDAY DAYS EDATE EOMONTH TODAY NOW VLOOKUP HLOOKUP XLOOKUP LOOKUP MATCH XMATCH INDEX ISBLANK ISNUMBER ISTEXT ISERROR ISERR ISNA N T ROW ROWS COLUMN COLUMNS CHOOSE NA').split(' '));

/** Only explicit A1 coordinates are transformable; dynamic/whole-axis forms fail closed. */
export function transformFormulaReferences(formula: string, worksheet: string, transforms: XlsxStructureTransform[]): string {
  if (!formula || formula.startsWith("=") || formula.length > 32768 || formula.includes("\u0000")) throw new Error("unsupported_formula_transform");
  let output = formula;
  for (const transform of transforms) {
    transformCellAddress({ row: 0, col: 0 }, "", [transform]);
    const strings: string[] = [];
    const masked = output.replace(/"(?:[^"]|"")*"|#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!)/gi, (value) => `\u0000${strings.push(value) - 1}\u0000`);
    if (/[\[\]#"]/.test(masked) || /\b\$?[A-Z]{1,3}:\$?[A-Z]{1,3}\b|\b\$?\d+:\$?\d+\b/i.test(masked)) throw new Error("unsupported_formula_transform");
    const reference = /(?<![A-Za-z0-9_.:])(?:(('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?)(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?(?![A-Za-z0-9_]|\s*\()/g;
    const rewritten = masked.replace(reference, (whole, prefix: string | undefined, rawSheet: string | undefined, first: string, last: string | undefined) => {
      const name = rawSheet ? rawSheet.startsWith("'") ? rawSheet.slice(1, -1).replace(/''/g, "'") : rawSheet : worksheet;
      const parse = (value: string) => cellFromAddress(value.replace(/\$/g, "").toUpperCase());
      const a = parse(first), b = last ? parse(last) : null;
      if (b && (b.row < a.row || b.col < a.col)) throw new Error("unsupported_formula_transform");
      if (name.toLowerCase() !== transform.worksheet.toLowerCase()) return whole;
      const axis = transform.kind.endsWith("rows") ? "row" : "col";
      const end = transform.index + transform.count;
      const shift = (value: number, boundary?: "start" | "end") => {
        if (value < transform.index) return value;
        if (transform.kind.startsWith("insert")) return value + transform.count;
        if (value >= end) return value - transform.count;
        return boundary === "start" ? transform.index : boundary === "end" ? transform.index - 1 : null;
      };
      if (b && transform.kind.startsWith("delete") && a[axis] >= transform.index && b[axis] < end) return "#REF!";
      const nextA = shift(a[axis], b ? "start" : undefined), nextB = b ? shift(b[axis], "end") : null;
      if (nextA === null || (b && (nextB === null || nextB < nextA))) return "#REF!";
      a[axis] = nextA;
      if (b && nextB !== null) b[axis] = nextB;
      const format = (cell: { row: number; col: number }, original: string) => {
        if (cell.row >= 1_048_576 || cell.col >= 16_384) throw new Error("unsupported_formula_transform");
        const address = addressFromCell(cell), parts = /^([A-Z]+)(\d+)$/.exec(address)!;
        return `${original.startsWith("$") ? "$" : ""}${parts[1]}${/\$\d/.test(original) ? "$" : ""}${parts[2]}`;
      };
      return `${prefix ?? ""}${format(a, first)}${b && last ? `:${format(b, last)}` : ""}`;
    });
    const unchecked = masked.replace(reference, '0').replace(/\u0000\d+\u0000/g, '0').replace(/\b\d+(?:\.\d*)?(?:E[+-]?\d+)?\b/gi, '0');
    if (/[!'$:@{}\\]/.test(unchecked)) throw new Error("unsupported_formula_transform");
    for (const match of unchecked.matchAll(/[A-Za-z_][A-Za-z0-9_.]*/g)) {
      const token = match[0].toUpperCase();
      const call = /^\s*\(/.test(unchecked.slice(match.index! + match[0].length));
      if (!(call && FUNCTIONS.has(token)) && !(!call && ['TRUE', 'FALSE'].includes(token))) throw new Error("unsupported_formula_transform");
    }
    output = rewritten.replace(/\u0000(\d+)\u0000/g, (_whole, index: string) => strings[Number(index)]);
  }
  return output;
}
