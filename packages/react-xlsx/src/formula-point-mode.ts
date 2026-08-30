/**
 * Excel-style formula point mode.
 *
 * While the in-cell editor holds a formula, clicking another cell inserts
 * that cell's reference into the formula instead of committing the edit and
 * moving the selection. A second click replaces the reference inserted by
 * the first (the "point span"), typing anything ends the replace window.
 *
 * The decision logic is pure so it can be tested exhaustively; the viewer
 * owns the React glue (state, focus, caret).
 */

export type FormulaPointSpan = { start: number; end: number };

export type FormulaPointInsertion = {
  next: string;
  span: FormulaPointSpan;
  caret: number;
};

/**
 * Characters after which a formula can accept a reference: the leading `=`,
 * arithmetic/text/comparison operators, an opening parenthesis or brace,
 * argument and union separators, and the range/intersection forms.
 */
const REFERENCE_ACCEPTING_TAIL = /[=+\-*/(,&^<>:%;{ ]$/;

export function applyFormulaPointInsert(
  value: string,
  span: FormulaPointSpan | null,
  reference: string
): FormulaPointInsertion | null {
  if (!value.startsWith("=") || !reference) {
    return null;
  }

  const activeSpan =
    span && span.start >= 0 && span.start <= span.end && span.end <= value.length
      ? span
      : null;

  const base = activeSpan ? value.slice(0, activeSpan.start) : value;
  const suffix = activeSpan ? value.slice(activeSpan.end) : "";

  if (!activeSpan && !REFERENCE_ACCEPTING_TAIL.test(value)) {
    // A complete token ends the formula ("=A1" typed out): clicking another
    // cell falls through to the normal commit-and-move behavior.
    return null;
  }

  const caret = base.length + reference.length;
  return {
    next: `${base}${reference}${suffix}`,
    span: { start: base.length, end: caret },
    caret,
  };
}
