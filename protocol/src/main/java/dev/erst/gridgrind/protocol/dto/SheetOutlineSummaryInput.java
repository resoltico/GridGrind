package dev.erst.gridgrind.protocol.dto;

/** Outline-summary placement authored as part of sheet-presentation state. */
public record SheetOutlineSummaryInput(Boolean rowSumsBelow, Boolean rowSumsRight) {
  /** Returns the effective Excel defaults for outline summary placement. */
  public static SheetOutlineSummaryInput defaults() {
    return new SheetOutlineSummaryInput(true, true);
  }

  public SheetOutlineSummaryInput {
    rowSumsBelow = rowSumsBelow == null ? defaults().rowSumsBelow() : rowSumsBelow;
    rowSumsRight = rowSumsRight == null ? defaults().rowSumsRight() : rowSumsRight;
  }
}
