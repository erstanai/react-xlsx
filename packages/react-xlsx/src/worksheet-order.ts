export function movedArrayIndex(index: number, from: number, to: number) {
  if (index === from) {
    return to;
  }
  if (from < to && index > from && index <= to) {
    return index - 1;
  }
  if (from > to && index >= to && index < from) {
    return index + 1;
  }
  return index;
}

export function moveArrayEntry<T>(values: readonly T[], from: number, to: number): T[] {
  if (
    !Number.isInteger(from)
    || !Number.isInteger(to)
    || from < 0
    || to < 0
    || from >= values.length
    || to >= values.length
  ) {
    throw new RangeError(`Cannot move array entry from ${from} to ${to}.`);
  }

  const next = values.slice();
  const [entry] = next.splice(from, 1);
  next.splice(to, 0, entry as T);
  return next;
}

export function moveWorkbookIndexedGroups<T extends { workbookSheetIndex: number }>(
  groups: readonly (readonly T[])[],
  from: number,
  to: number
): T[][] {
  return moveArrayEntry(groups, from, to).map((group, workbookSheetIndex) => (
    group.map((entry) => ({
      ...entry,
      workbookSheetIndex
    }))
  ));
}

export function moveWorkbookIndexedEntries<T extends { workbookSheetIndex: number }>(
  entries: readonly (T | null)[],
  from: number,
  to: number
): Array<T | null> {
  return moveArrayEntry(entries, from, to).map((entry, workbookSheetIndex) => (
    entry
      ? {
          ...entry,
          workbookSheetIndex
        }
      : null
  ));
}
