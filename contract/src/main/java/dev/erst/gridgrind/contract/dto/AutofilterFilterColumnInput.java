package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** One authored autofilter filter-column payload. */
public record AutofilterFilterColumnInput(
    long columnId, boolean showButton, AutofilterFilterCriterionInput criterion) {
  /** Creates a filter-column payload with an explicit visible button. */
  public static AutofilterFilterColumnInput visibleButton(
      long columnId, AutofilterFilterCriterionInput criterion) {
    return new AutofilterFilterColumnInput(columnId, true, criterion);
  }

  public AutofilterFilterColumnInput {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    Objects.requireNonNull(criterion, "criterion must not be null");
  }

  /** Creates a filter-column payload from the authored wire shape. */
  @JsonCreator
  public AutofilterFilterColumnInput(
      @JsonProperty("columnId") long columnId,
      @JsonProperty("showButton") Boolean showButton,
      @JsonProperty("criterion") AutofilterFilterCriterionInput criterion) {
    this(
        columnId,
        Objects.requireNonNull(showButton, "showButton must not be null").booleanValue(),
        criterion);
  }
}
