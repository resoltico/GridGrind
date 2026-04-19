package dev.erst.gridgrind.contract.dto;

/** Factual outline-summary placement reported for one worksheet. */
public record SheetOutlineSummaryReport(boolean rowSumsBelow, boolean rowSumsRight) {
  /** Returns the effective Excel defaults for outline summary placement. */
  public static SheetOutlineSummaryReport defaults() {
    return new SheetOutlineSummaryReport(true, true);
  }
}
