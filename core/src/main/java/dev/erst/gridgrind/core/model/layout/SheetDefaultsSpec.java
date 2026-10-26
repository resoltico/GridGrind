package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Default row and column sizing authored as part of sheet-presentation state. */
@CatalogType(
    id = "SheetDefaultsSpec",
    summary =
        "Default row and column sizing authored as part of sheet-presentation state. Supply explicit defaultColumnWidth and defaultRowHeightPoints values. defaultColumnWidth must be > 0 and <= 255; defaultRowHeightPoints must be > 0 and <= 409.0.",
    displayOrder = 66)
public record SheetDefaultsSpec(int defaultColumnWidth, double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static SheetDefaultsSpec defaults() {
    return new SheetDefaultsSpec(8, 15.0d);
  }

  public SheetDefaultsSpec {
    dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaultColumnWidth, "defaultColumnWidth");
    dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaultRowHeightPoints, "defaultRowHeightPoints");
  }
}
