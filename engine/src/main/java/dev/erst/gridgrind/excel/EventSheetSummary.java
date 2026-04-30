package dev.erst.gridgrind.excel;

/** Minimal factual EVENT_READ sheet-summary data gathered from one worksheet stream. */
record EventSheetSummary(
    boolean selected,
    WorkbookSheetResult.SheetProtection protection,
    int physicalRowCount,
    int lastRowIndex,
    int lastColumnIndex) {}
