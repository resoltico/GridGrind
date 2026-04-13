package dev.erst.gridgrind.excel;

/** Default row and column sizing applied when individual rows or columns do not override size. */
public record ExcelSheetDefaults(int defaultColumnWidth, double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static ExcelSheetDefaults defaults() {
    return new ExcelSheetDefaults(8, 15.0d);
  }

  public ExcelSheetDefaults {
    if (defaultColumnWidth <= 0) {
      throw new IllegalArgumentException("defaultColumnWidth must be greater than 0");
    }
    if (!Double.isFinite(defaultRowHeightPoints) || defaultRowHeightPoints <= 0.0d) {
      throw new IllegalArgumentException(
          "defaultRowHeightPoints must be finite and greater than 0");
    }
  }
}
