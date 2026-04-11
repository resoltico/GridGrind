package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Factual sort-state metadata stored on one autofilter definition. */
public record ExcelAutofilterSortStateSnapshot(
    String range,
    boolean caseSensitive,
    boolean columnSort,
    String sortMethod,
    List<ExcelAutofilterSortConditionSnapshot> conditions) {
  public ExcelAutofilterSortStateSnapshot {
    Objects.requireNonNull(range, "range must not be null");
    sortMethod = sortMethod == null ? "" : sortMethod;
    conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
    for (ExcelAutofilterSortConditionSnapshot condition : conditions) {
      Objects.requireNonNull(condition, "conditions must not contain null values");
    }
  }
}
