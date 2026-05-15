package dev.erst.gridgrind.excel.event;

import dev.erst.gridgrind.excel.WorkbookSheetResult;

/** Minimal factual EVENT_READ sheet-summary data gathered from one worksheet stream. */
public record EventSheetSummary(
    boolean selected,
    WorkbookSheetResult.SheetProtection protection,
    int physicalRowCount,
    int lastRowIndex,
    int lastColumnIndex) {}
