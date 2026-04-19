package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Authored sort-state payload nested under sheet autofilter authoring. */
public record AutofilterSortStateInput(
    String range,
    Boolean caseSensitive,
    Boolean columnSort,
    String sortMethod,
    List<AutofilterSortConditionInput> conditions) {
  public AutofilterSortStateInput {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    caseSensitive = Boolean.TRUE.equals(caseSensitive);
    columnSort = Boolean.TRUE.equals(columnSort);
    sortMethod = sortMethod == null ? "" : sortMethod;
    conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
    if (conditions.isEmpty()) {
      throw new IllegalArgumentException("conditions must not be empty");
    }
    for (AutofilterSortConditionInput condition : conditions) {
      Objects.requireNonNull(condition, "conditions must not contain null values");
    }
  }
}
