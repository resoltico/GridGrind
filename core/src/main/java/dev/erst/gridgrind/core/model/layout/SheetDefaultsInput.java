package dev.erst.gridgrind.core.model.layout;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;

/** Default row and column sizing authored as part of sheet-presentation state. */
@CatalogType(
    id = "SheetDefaultsInput",
    summary =
        "Default row and column sizing authored as part of sheet-presentation state. Supply explicit defaultColumnWidth and defaultRowHeightPoints values. defaultColumnWidth must be > 0 and <= 255; defaultRowHeightPoints must be > 0 and <= 409.0.",
    displayOrder = 66)
public record SheetDefaultsInput(int defaultColumnWidth, double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static SheetDefaultsInput defaults() {
    return new SheetDefaultsInput(8, 15.0d);
  }

  public SheetDefaultsInput {
    dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaultColumnWidth, "defaultColumnWidth");
    dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaultRowHeightPoints, "defaultRowHeightPoints");
  }

  /** Creates sheet defaults while applying the documented row-height and column-width defaults. */
  @JsonCreator
  public SheetDefaultsInput(
      @JsonProperty("defaultColumnWidth") Integer defaultColumnWidth,
      @JsonProperty("defaultRowHeightPoints") Double defaultRowHeightPoints) {
    this(
        java.util.Objects.requireNonNull(defaultColumnWidth, "defaultColumnWidth must not be null")
            .intValue(),
        java.util.Objects.requireNonNull(
                defaultRowHeightPoints, "defaultRowHeightPoints must not be null")
            .doubleValue());
  }
}
