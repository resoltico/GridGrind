package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;

/** One workbook sheet entry discovered from workbook.xml during EVENT_READ metadata parsing. */
record EventSheetReference(String name, String relationshipId, ExcelSheetVisibility visibility) {}
