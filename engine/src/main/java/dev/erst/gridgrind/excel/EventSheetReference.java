package dev.erst.gridgrind.excel;

/** One workbook sheet entry discovered from workbook.xml during EVENT_READ metadata parsing. */
record EventSheetReference(String name, String relationshipId, ExcelSheetVisibility visibility) {}
