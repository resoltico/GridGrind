package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One authored sort condition nested under a mutable-workbook autofilter sort state. */
public record ExcelAutofilterSortCondition(
    String range, boolean descending, String sortBy, ExcelColor color, Integer iconId) {
  public ExcelAutofilterSortCondition {
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
