package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;

/** Default row and column sizing authored as part of sheet-presentation state. */
public record SheetDefaultsInput(int defaultColumnWidth, double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static SheetDefaultsInput defaults() {
    return new SheetDefaultsInput(8, 15.0d);
  }

  public SheetDefaultsInput {
    dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaultColumnWidth, "defaultColumnWidth");
    dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaultRowHeightPoints, "defaultRowHeightPoints");
  }

  /** Creates sheet defaults while applying the documented row-height and column-width defaults. */
  @JsonCreator
  public SheetDefaultsInput(Integer defaultColumnWidth, Double defaultRowHeightPoints) {
    this(
        java.util.Objects.requireNonNull(defaultColumnWidth, "defaultColumnWidth must not be null")
            .intValue(),
        java.util.Objects.requireNonNull(
                defaultRowHeightPoints, "defaultRowHeightPoints must not be null")
            .doubleValue());
  }
}
