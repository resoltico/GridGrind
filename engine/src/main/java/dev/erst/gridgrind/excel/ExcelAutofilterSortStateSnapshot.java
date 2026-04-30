package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual sort-state metadata stored on one autofilter definition. */
public record ExcelAutofilterSortStateSnapshot(
    String range,
    boolean caseSensitive,
    boolean columnSort,
    Optional<ExcelAutofilterSortMethod> sortMethod,
    List<ExcelAutofilterSortConditionSnapshot> conditions) {
  public ExcelAutofilterSortStateSnapshot {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Optional<ExcelAutofilterSortMethod> normalizedSortMethod =
        Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    sortMethod = normalizedSortMethod;
    conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
    for (ExcelAutofilterSortConditionSnapshot condition : conditions) {
      Objects.requireNonNull(condition, "conditions must not contain null values");
    }
  }
}
