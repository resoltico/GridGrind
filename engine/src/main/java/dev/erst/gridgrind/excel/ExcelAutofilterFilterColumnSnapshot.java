package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Factual metadata for one persisted autofilter filter column. */
public record ExcelAutofilterFilterColumnSnapshot(
    long columnId, boolean showButton, ExcelAutofilterFilterCriterionSnapshot criterion) {
  public ExcelAutofilterFilterColumnSnapshot {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    Objects.requireNonNull(criterion, "criterion must not be null");
  }
}
