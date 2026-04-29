package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.List;
import java.util.Objects;

/** Authored sort-state payload nested under sheet autofilter authoring. */
public record AutofilterSortStateInput(
    String range,
    Boolean caseSensitive,
    Boolean columnSort,
    String sortMethod,
    List<AutofilterSortConditionInput> conditions) {
  /** Creates a sort-state payload with the implicit default sort method. */
  public AutofilterSortStateInput(
      String range,
      Boolean caseSensitive,
      Boolean columnSort,
      List<AutofilterSortConditionInput> conditions) {
    this(range, caseSensitive, columnSort, "", conditions);
  }

  public AutofilterSortStateInput {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(caseSensitive, "caseSensitive must not be null");
    Objects.requireNonNull(columnSort, "columnSort must not be null");
    Objects.requireNonNull(sortMethod, "sortMethod must not be null");
    conditions = List.copyOf(Objects.requireNonNull(conditions, "conditions must not be null"));
    if (conditions.isEmpty()) {
      throw new IllegalArgumentException("conditions must not be empty");
    }
    for (AutofilterSortConditionInput condition : conditions) {
      Objects.requireNonNull(condition, "conditions must not contain null values");
    }
  }

  @JsonCreator
  static AutofilterSortStateInput create(
      @JsonProperty("range") String range,
      @JsonProperty("caseSensitive") Boolean caseSensitive,
      @JsonProperty("columnSort") Boolean columnSort,
      @JsonProperty("sortMethod") String sortMethod,
      @JsonProperty("conditions") List<AutofilterSortConditionInput> conditions) {
    return new AutofilterSortStateInput(
        range,
        Boolean.TRUE.equals(caseSensitive),
        Boolean.TRUE.equals(columnSort),
        sortMethod == null ? "" : sortMethod,
        conditions);
  }
}
