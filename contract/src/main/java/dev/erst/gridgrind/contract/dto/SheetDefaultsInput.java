package dev.erst.gridgrind.contract.dto;

/** Default row and column sizing authored as part of sheet-presentation state. */
public record SheetDefaultsInput(Integer defaultColumnWidth, Double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static SheetDefaultsInput defaults() {
    return new SheetDefaultsInput(8, 15.0d);
  }

  public SheetDefaultsInput {
    defaultColumnWidth =
        defaultColumnWidth == null ? defaults().defaultColumnWidth() : defaultColumnWidth;
    defaultRowHeightPoints =
        defaultRowHeightPoints == null
            ? defaults().defaultRowHeightPoints()
            : defaultRowHeightPoints;
    dev.erst.gridgrind.excel.ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaultColumnWidth, "defaultColumnWidth");
    dev.erst.gridgrind.excel.ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaultRowHeightPoints, "defaultRowHeightPoints");
  }
}
