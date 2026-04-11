package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Authored sort-state payload nested under mutable-workbook autofilter authoring. */
public record ExcelAutofilterSortState(
    String range,
    boolean caseSensitive,
    boolean columnSort,
    String sortMethod,
    List<ExcelAutofilterSortCondition> conditions) {
  public ExcelAutofilterSortState {
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
    if (conditions.isEmpty()) {
      throw new IllegalArgumentException("conditions must not be empty");
    }
    for (ExcelAutofilterSortCondition condition : conditions) {
      Objects.requireNonNull(condition, "conditions must not contain null values");
    }
  }
}
