package dev.erst.gridgrind.protocol.dto;

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
    if (defaultColumnWidth <= 0) {
      throw new IllegalArgumentException("defaultColumnWidth must be greater than 0");
    }
    if (!Double.isFinite(defaultRowHeightPoints) || defaultRowHeightPoints <= 0.0d) {
      throw new IllegalArgumentException(
          "defaultRowHeightPoints must be finite and greater than 0");
    }
  }
}
