package dev.erst.gridgrind.engine.backend.eventread;

import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;

/** One workbook sheet entry discovered from workbook.xml during EVENT_READ metadata parsing. */
public record EventSheetReference(
    String name, String relationshipId, ExcelSheetVisibility visibility) {}
