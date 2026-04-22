package dev.erst.gridgrind.contract.dto;

/**
 * Factual default row and column sizing reported for one worksheet.
 *
 * <p>Readback stays truthful on reopen, so malformed positive persisted default row heights are
 * preserved instead of being clamped to normal authoring ceilings.
 */
public record SheetDefaultsReport(int defaultColumnWidth, double defaultRowHeightPoints) {
  /** Returns the effective Excel defaults for one new worksheet. */
  public static SheetDefaultsReport defaults() {
    return new SheetDefaultsReport(8, 15.0d);
  }

  public SheetDefaultsReport {
    if (defaultColumnWidth <= 0) {
      throw new IllegalArgumentException("defaultColumnWidth must be greater than 0");
    }
    if (!Double.isFinite(defaultRowHeightPoints) || defaultRowHeightPoints <= 0.0d) {
      throw new IllegalArgumentException(
          "defaultRowHeightPoints must be finite and greater than 0");
    }
  }
}
