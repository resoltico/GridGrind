package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** One authored autofilter filter-column payload. */
public record AutofilterFilterColumnInput(
    long columnId, Boolean showButton, AutofilterFilterCriterionInput criterion) {
  public AutofilterFilterColumnInput {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    showButton = showButton == null ? Boolean.TRUE : showButton;
    Objects.requireNonNull(criterion, "criterion must not be null");
  }
}
