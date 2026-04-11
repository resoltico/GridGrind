package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One factual sort condition stored inside an autofilter sort-state payload. */
public record ExcelAutofilterSortConditionSnapshot(
    String range, boolean descending, String sortBy, ExcelColorSnapshot color, Integer iconId) {
  public ExcelAutofilterSortConditionSnapshot {
    Objects.requireNonNull(range, "range must not be null");
    sortBy = sortBy == null ? "" : sortBy;
    if (iconId != null && iconId < 0) {
      throw new IllegalArgumentException("iconId must not be negative");
    }
  }
}
