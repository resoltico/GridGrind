package dev.erst.gridgrind.engine.backend.eventread;

import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;

/** Minimal factual EVENT_READ sheet-summary data gathered from one worksheet stream. */
public record EventSheetSummary(
    boolean selected,
    SheetProtectionReport protection,
    int physicalRowCount,
    int lastRowIndex,
    int lastColumnIndex) {}
