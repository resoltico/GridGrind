package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** One authored autofilter filter-column payload. */
public record AutofilterFilterColumnInput(
    long columnId, Boolean showButton, AutofilterFilterCriterionInput criterion) {
  /** Creates a filter-column payload with Excel's default visible show-button state. */
  public AutofilterFilterColumnInput(long columnId, AutofilterFilterCriterionInput criterion) {
    this(columnId, true, criterion);
  }

  public AutofilterFilterColumnInput {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    Objects.requireNonNull(showButton, "showButton must not be null");
    Objects.requireNonNull(criterion, "criterion must not be null");
  }

  @JsonCreator
  static AutofilterFilterColumnInput create(
      @JsonProperty("columnId") long columnId,
      @JsonProperty("showButton") Boolean showButton,
      @JsonProperty("criterion") AutofilterFilterCriterionInput criterion) {
    return new AutofilterFilterColumnInput(
        columnId, showButton == null ? Boolean.TRUE : showButton, criterion);
  }
}
