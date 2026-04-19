package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** One factual sort condition stored inside an autofilter sort-state payload. */
public record AutofilterSortConditionReport(
    String range, boolean descending, String sortBy, CellColorReport color, Integer iconId) {
  public AutofilterSortConditionReport {
    Objects.requireNonNull(range, "range must not be null");
    sortBy = sortBy == null ? "" : sortBy;
    if (iconId != null && iconId < 0) {
      throw new IllegalArgumentException("iconId must not be negative");
    }
  }
}
