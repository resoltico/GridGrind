package dev.erst.gridgrind.excel;

/** Controls how the workbook formula runtime handles missing external workbook references. */
public enum ExcelFormulaMissingWorkbookPolicy {
  /** Reject evaluation when any external workbook reference cannot be resolved. */
  ERROR,

  /** Use the formula cell's cached value when an external workbook reference is unavailable. */
  USE_CACHED_VALUE
}
