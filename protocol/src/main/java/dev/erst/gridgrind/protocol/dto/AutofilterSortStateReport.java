package dev.erst.gridgrind.protocol.dto;

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
    conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
    for (AutofilterSortConditionReport condition : conditions) {
      Objects.requireNonNull(condition, "conditions must not contain null values");
    }
  }
}
