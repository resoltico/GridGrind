package dev.erst.gridgrind.excel.foundation;

/**
 * GridGrind-owned sheet-visibility states for workbook and sheet summaries and mutations.
 *
 * <p>When surfaced through the contract module, these constant names are part of the published wire
 * vocabulary and must be treated as breaking-change-sensitive.
 */
public enum ExcelSheetVisibility {
  VISIBLE,
  HIDDEN,
  VERY_HIDDEN
}
