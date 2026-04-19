package dev.erst.gridgrind.contract.dto;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual sort-state metadata stored on one autofilter definition. */
public record AutofilterSortStateReport(
    String range,
    boolean caseSensitive,
    boolean columnSort,
    String sortMethod,
    List<AutofilterSortConditionReport> conditions) {
  public AutofilterSortStateReport {
    Objects.requireNonNull(range, "range must not be null");
    sortMethod = sortMethod == null ? "" : sortMethod;
    conditions = copyValues(conditions, "conditions");
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copy);
  }
}
