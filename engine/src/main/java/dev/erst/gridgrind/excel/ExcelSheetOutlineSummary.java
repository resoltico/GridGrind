package dev.erst.gridgrind.excel;

/** Outline-summary placement defaults that affect future grouped rows and columns. */
public record ExcelSheetOutlineSummary(boolean rowSumsBelow, boolean rowSumsRight) {
  /** Returns the effective Excel defaults for outline summary placement. */
  public static ExcelSheetOutlineSummary defaults() {
    return new ExcelSheetOutlineSummary(true, true);
  }
}
