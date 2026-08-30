/**
 * Relative A1 reference adjustment for copy/paste and fill operations.
 *
 * Excel semantics: when a formula moves by (rowDelta, colDelta), every
 * relative component of every cell reference shifts by the same delta while
 * `$`-anchored components stay fixed. References that would land outside the
 * sheet become `#REF!`.
 *
 * The tokenizer is deliberately conservative:
 * - String literals (`"..."`, with `""` escapes) are never touched.
 * - Quoted sheet prefixes (`'Q1 Data'!B2`) are preserved verbatim; the
 *   reference that follows them still adjusts.
 * - A candidate followed by `(` or another identifier character is a
 *   function or name (`LOG10(`, `TAX_2024`), not a reference.
 * - Whole-column (`B:B`) and whole-row (`2:2`) references are left unchanged
 *   rather than half-adjusted.
 *
 * Excel forbids defined names that collide with valid A1 addresses, so a
 * matched `letters+digits` token inside the valid column/row space is always
 * a real reference.
 */

const MAX_COL = 16_384; // XFD
const MAX_ROW = 1_048_576;

function columnLettersToNumber(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + (letter.charCodeAt(0) - 64);
  }
  return value;
}

function columnNumberToLetters(column: number): string {
  let value = column;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

// The `!` in the lookahead guards unquoted sheet names that would parse as
// references (`X5!A1`); a real reference is never followed by `!`.
const REFERENCE_PATTERN = /(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})(?![A-Za-z0-9_.(!])/g;

function adjustSegment(segment: string, rowDelta: number, colDelta: number): string {
  REFERENCE_PATTERN.lastIndex = 0;
  return segment.replace(
    REFERENCE_PATTERN,
    (match, colAnchor: string, colLetters: string, rowAnchor: string, rowDigits: string, offset: number) => {
      // A candidate preceded by an identifier character is the tail of a
      // longer name (`_B2`), except `$`/`:`/`!` which legitimately precede
      // references and are not part of identifiers.
      const preceding = segment[offset - 1];
      if (preceding !== undefined && /[A-Za-z0-9_.]/.test(preceding)) {
        return match;
      }

      const column = columnLettersToNumber(colLetters);
      const row = Number(rowDigits);
      if (column > MAX_COL || row > MAX_ROW) {
        return match;
      }

      const nextColumn = colAnchor === "$" ? column : column + colDelta;
      const nextRow = rowAnchor === "$" ? row : row + rowDelta;
      if (nextColumn < 1 || nextColumn > MAX_COL || nextRow < 1 || nextRow > MAX_ROW) {
        return "#REF!";
      }

      const letters = colLetters === colLetters.toLowerCase() && colLetters !== colLetters.toUpperCase()
        ? columnNumberToLetters(nextColumn).toLowerCase()
        : columnNumberToLetters(nextColumn);
      return `${colAnchor}${letters}${rowAnchor}${nextRow}`;
    }
  );
}

export function offsetFormulaReferences(formula: string, rowDelta: number, colDelta: number): string {
  if ((rowDelta === 0 && colDelta === 0) || !formula) {
    return formula;
  }

  let result = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];

    if (char === '"') {
      // String literal: copy through the closing quote, honoring "" escapes.
      let end = index + 1;
      while (end < formula.length) {
        if (formula[end] === '"') {
          if (formula[end + 1] === '"') {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += formula.slice(index, end);
      index = end;
      continue;
    }

    if (char === "'") {
      // Quoted sheet name: copy verbatim through the closing quote,
      // honoring '' escapes. The reference after the `!` still adjusts.
      let end = index + 1;
      while (end < formula.length) {
        if (formula[end] === "'") {
          if (formula[end + 1] === "'") {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += formula.slice(index, end);
      index = end;
      continue;
    }

    // Plain segment: run to the next literal boundary and adjust it whole.
    let end = index;
    while (end < formula.length && formula[end] !== '"' && formula[end] !== "'") {
      end += 1;
    }
    result += adjustSegment(formula.slice(index, end), rowDelta, colDelta);
    index = end;
  }

  return result;
}
