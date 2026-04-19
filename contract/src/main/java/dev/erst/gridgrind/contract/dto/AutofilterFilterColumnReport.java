package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual metadata for one persisted autofilter filter column. */
public record AutofilterFilterColumnReport(
    long columnId, boolean showButton, AutofilterFilterCriterionReport criterion) {
  public AutofilterFilterColumnReport {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    Objects.requireNonNull(criterion, "criterion must not be null");
  }
}
