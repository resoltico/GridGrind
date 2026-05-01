package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Authored sort-state payload nested under sheet autofilter authoring. */
public record AutofilterSortStateInput(
    String range,
    boolean caseSensitive,
    boolean columnSort,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelAutofilterSortMethod> sortMethod,
    List<AutofilterSortConditionInput> conditions) {
  /** Creates a sort-state payload with no explicit sort-method override. */
  public static AutofilterSortStateInput withoutSortMethod(
      String range,
      boolean caseSensitive,
      boolean columnSort,
      List<AutofilterSortConditionInput> conditions) {
    return new AutofilterSortStateInput(
        range, caseSensitive, columnSort, Optional.empty(), conditions);
  }

  public AutofilterSortStateInput {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Optional<ExcelAutofilterSortMethod> normalizedSortMethod =
        Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    sortMethod = normalizedSortMethod;
    conditions = copyConditions(conditions);
    if (conditions.isEmpty()) {
      throw new IllegalArgumentException("conditions must not be empty");
    }
  }

  /** Reads one sort-state payload with explicit booleans and one typed optional sort method. */
  @JsonCreator
  public AutofilterSortStateInput(
      @JsonProperty("range") String range,
      @JsonProperty("caseSensitive") Boolean caseSensitive,
      @JsonProperty("columnSort") Boolean columnSort,
      @JsonProperty("sortMethod") Optional<ExcelAutofilterSortMethod> sortMethod,
      @JsonProperty("conditions") List<AutofilterSortConditionInput> conditions) {
    this(
        range,
        Objects.requireNonNull(caseSensitive, "caseSensitive must not be null").booleanValue(),
        Objects.requireNonNull(columnSort, "columnSort must not be null").booleanValue(),
        Objects.requireNonNull(sortMethod, "sortMethod must not be null"),
        conditions);
  }

  private static List<AutofilterSortConditionInput> copyConditions(
      List<AutofilterSortConditionInput> conditions) {
    Objects.requireNonNull(conditions, "conditions must not be null");
    List<AutofilterSortConditionInput> copiedConditions = new ArrayList<>(conditions.size());
    for (AutofilterSortConditionInput condition : conditions) {
      copiedConditions.add(
          Objects.requireNonNull(condition, "conditions must not contain null values"));
    }
    return List.copyOf(copiedConditions);
  }
}
