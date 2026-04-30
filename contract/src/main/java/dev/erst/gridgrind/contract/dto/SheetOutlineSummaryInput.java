package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;

/** Outline-summary placement authored as part of sheet-presentation state. */
public record SheetOutlineSummaryInput(boolean rowSumsBelow, boolean rowSumsRight) {
  /** Returns the effective Excel defaults for outline summary placement. */
  public static SheetOutlineSummaryInput defaults() {
    return new SheetOutlineSummaryInput(true, true);
  }

  public SheetOutlineSummaryInput {}

  /** Creates outline-summary settings while applying the documented summary-placement defaults. */
  @JsonCreator
  public SheetOutlineSummaryInput(Boolean rowSumsBelow, Boolean rowSumsRight) {
    this(
        java.util.Objects.requireNonNull(rowSumsBelow, "rowSumsBelow must not be null")
            .booleanValue(),
        java.util.Objects.requireNonNull(rowSumsRight, "rowSumsRight must not be null")
            .booleanValue());
  }
}
