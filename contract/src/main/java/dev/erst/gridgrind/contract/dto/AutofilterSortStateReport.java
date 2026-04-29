package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
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
  /** Creates a report sort-state with the implicit default sort method. */
  public AutofilterSortStateReport(
      String range,
      boolean caseSensitive,
      boolean columnSort,
      List<AutofilterSortConditionReport> conditions) {
    this(range, caseSensitive, columnSort, "", conditions);
  }

  public AutofilterSortStateReport {
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    conditions = copyValues(conditions, "conditions");
  }

  @JsonCreator
  static AutofilterSortStateReport create(
      @JsonProperty("range") String range,
      @JsonProperty("caseSensitive") boolean caseSensitive,
      @JsonProperty("columnSort") boolean columnSort,
      @JsonProperty("sortMethod") String sortMethod,
      @JsonProperty("conditions") List<AutofilterSortConditionReport> conditions) {
    return new AutofilterSortStateReport(
        range, caseSensitive, columnSort, sortMethod == null ? "" : sortMethod, conditions);
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
