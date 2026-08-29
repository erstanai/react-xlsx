export type InitialViewportSheetIdentity = {
  workbookSheetIndex: number;
  name: string;
  // These values are intentionally accepted but excluded from the identity.
  // Editing can expand the used range without changing the active worksheet.
  minUsedRow: number;
  minUsedCol: number;
};

export function buildInitialViewportKey(input: {
  displayFileName: string;
  activeSheetIndex: number;
  activeSheet: InitialViewportSheetIdentity | null | undefined;
  isWorkerBacked: boolean;
}): string | null {
  if (!input.activeSheet) return null;
  return [
    input.displayFileName,
    input.activeSheetIndex,
    input.activeSheet.workbookSheetIndex,
    input.activeSheet.name,
    input.isWorkerBacked ? "worker" : "main",
  ].join("|");
}
