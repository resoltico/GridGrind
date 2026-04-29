package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** Outline-summary placement authored as part of sheet-presentation state. */
public record SheetOutlineSummaryInput(Boolean rowSumsBelow, Boolean rowSumsRight) {
  /** Returns the effective Excel defaults for outline summary placement. */
  public static SheetOutlineSummaryInput defaults() {
    return new SheetOutlineSummaryInput(true, true);
  }

  public SheetOutlineSummaryInput {
    Objects.requireNonNull(rowSumsBelow, "rowSumsBelow must not be null");
    Objects.requireNonNull(rowSumsRight, "rowSumsRight must not be null");
  }

  @JsonCreator
  static SheetOutlineSummaryInput create(
      @JsonProperty("rowSumsBelow") Boolean rowSumsBelow,
      @JsonProperty("rowSumsRight") Boolean rowSumsRight) {
    SheetOutlineSummaryInput defaults = defaults();
    return new SheetOutlineSummaryInput(
        rowSumsBelow == null ? defaults.rowSumsBelow() : rowSumsBelow,
        rowSumsRight == null ? defaults.rowSumsRight() : rowSumsRight);
  }
}
