package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** Default row and column sizing authored as part of sheet-presentation state. */
public record SheetDefaultsInput(Integer defaultColumnWidth, Double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static SheetDefaultsInput defaults() {
    return new SheetDefaultsInput(8, 15.0d);
  }

  public SheetDefaultsInput {
    Objects.requireNonNull(defaultColumnWidth, "defaultColumnWidth must not be null");
    Objects.requireNonNull(defaultRowHeightPoints, "defaultRowHeightPoints must not be null");
    dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaultColumnWidth, "defaultColumnWidth");
    dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaultRowHeightPoints, "defaultRowHeightPoints");
  }

  @JsonCreator
  static SheetDefaultsInput create(
      @JsonProperty("defaultColumnWidth") Integer defaultColumnWidth,
      @JsonProperty("defaultRowHeightPoints") Double defaultRowHeightPoints) {
    SheetDefaultsInput defaults = defaults();
    return new SheetDefaultsInput(
        defaultColumnWidth == null ? defaults.defaultColumnWidth() : defaultColumnWidth,
        defaultRowHeightPoints == null
            ? defaults.defaultRowHeightPoints()
            : defaultRowHeightPoints);
  }
}
