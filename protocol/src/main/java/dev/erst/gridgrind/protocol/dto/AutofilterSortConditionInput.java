package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** One authored sort condition nested inside an autofilter sort-state payload. */
public record AutofilterSortConditionInput(
    String range, boolean descending, String sortBy, ColorInput color, Integer iconId) {
  public AutofilterSortConditionInput {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    sortBy = sortBy == null ? "" : sortBy;
    if (iconId != null && iconId < 0) {
      throw new IllegalArgumentException("iconId must not be negative");
    }
  }
}
