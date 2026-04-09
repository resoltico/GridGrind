package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelSheetNames;
import java.util.Objects;

/** Protocol-facing explicit cell or rectangular range target for named-range authoring. */
public record NamedRangeTarget(String sheetName, String range) {
  public NamedRangeTarget {
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
  }
}
