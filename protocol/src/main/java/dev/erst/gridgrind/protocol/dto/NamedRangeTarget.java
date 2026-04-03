package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Protocol-facing explicit cell or rectangular range target for named-range authoring. */
public record NamedRangeTarget(String sheetName, String range) {
  public NamedRangeTarget {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(range, "range must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    if (sheetName.length() > 31) {
      throw new IllegalArgumentException("sheetName must not exceed 31 characters: " + sheetName);
    }
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
  }
}
