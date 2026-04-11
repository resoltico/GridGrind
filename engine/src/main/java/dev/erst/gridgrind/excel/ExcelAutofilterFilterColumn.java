package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One authored autofilter filter-column payload for a mutable workbook. */
public record ExcelAutofilterFilterColumn(
    long columnId, boolean showButton, ExcelAutofilterFilterCriterion criterion) {
  public ExcelAutofilterFilterColumn {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    Objects.requireNonNull(criterion, "criterion must not be null");
  }
}
