package dev.erst.gridgrind.excel;

/** Screen-facing sheet display flags captured and authored as one cohesive view state. */
public record ExcelSheetDisplay(
    boolean displayGridlines,
    boolean displayZeros,
    boolean displayRowColHeadings,
    boolean displayFormulas,
    boolean rightToLeft) {
  /** Returns the effective Excel defaults for one unconfigured worksheet view. */
  public static ExcelSheetDisplay defaults() {
    return new ExcelSheetDisplay(true, true, true, false, false);
  }
}
