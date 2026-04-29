package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** One factual sort condition stored inside an autofilter sort-state payload. */
public record AutofilterSortConditionReport(
    String range, boolean descending, String sortBy, CellColorReport color, Integer iconId) {
  /** Creates a report condition with the implicit default sort-by mode. */
  public AutofilterSortConditionReport(
      String range, boolean descending, CellColorReport color, Integer iconId) {
    this(range, descending, "", color, iconId);
  }

  public AutofilterSortConditionReport {
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(sortBy, "sortBy must not be null");
    if (iconId != null && iconId < 0) {
      throw new IllegalArgumentException("iconId must not be negative");
    }
  }

  @JsonCreator
  static AutofilterSortConditionReport create(
      @JsonProperty("range") String range,
      @JsonProperty("descending") boolean descending,
      @JsonProperty("sortBy") String sortBy,
      @JsonProperty("color") CellColorReport color,
      @JsonProperty("iconId") Integer iconId) {
    return new AutofilterSortConditionReport(
        range, descending, sortBy == null ? "" : sortBy, color, iconId);
  }
}
