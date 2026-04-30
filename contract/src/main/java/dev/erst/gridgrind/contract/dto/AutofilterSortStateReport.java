package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual sort-state metadata stored on one autofilter definition. */
public record AutofilterSortStateReport(
    String range,
    boolean caseSensitive,
    boolean columnSort,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelAutofilterSortMethod> sortMethod,
    List<AutofilterSortConditionReport> conditions) {
  /** Creates a report sort-state with no explicit sort-method override. */
  public static AutofilterSortStateReport withoutSortMethod(
      String range,
      boolean caseSensitive,
      boolean columnSort,
      List<AutofilterSortConditionReport> conditions) {
    return new AutofilterSortStateReport(
        range, caseSensitive, columnSort, Optional.empty(), conditions);
  }

  public AutofilterSortStateReport {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Optional<ExcelAutofilterSortMethod> normalizedSortMethod =
        Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    sortMethod = normalizedSortMethod;
    conditions = copyConditions(conditions);
  }

  @JsonCreator
  static AutofilterSortStateReport create(
      @JsonProperty("range") String range,
      @JsonProperty("caseSensitive") boolean caseSensitive,
      @JsonProperty("columnSort") boolean columnSort,
      @JsonProperty("sortMethod") Optional<ExcelAutofilterSortMethod> sortMethod,
      @JsonProperty("conditions") List<AutofilterSortConditionReport> conditions) {
    return new AutofilterSortStateReport(range, caseSensitive, columnSort, sortMethod, conditions);
  }

  private static List<AutofilterSortConditionReport> copyConditions(
      List<AutofilterSortConditionReport> conditions) {
    Objects.requireNonNull(conditions, "conditions must not be null");
    List<AutofilterSortConditionReport> copiedConditions = new ArrayList<>(conditions.size());
    for (AutofilterSortConditionReport condition : conditions) {
      copiedConditions.add(
          Objects.requireNonNull(condition, "conditions must not contain null values"));
    }
    return List.copyOf(copiedConditions);
  }
}
