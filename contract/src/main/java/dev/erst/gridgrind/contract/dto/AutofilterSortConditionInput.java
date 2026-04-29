package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** One authored sort condition nested inside an autofilter sort-state payload. */
public record AutofilterSortConditionInput(
    String range, boolean descending, String sortBy, ColorInput color, Integer iconId) {
  /** Creates a sort condition with the implicit workbook-default sort-by mode. */
  public AutofilterSortConditionInput(
      String range, boolean descending, ColorInput color, Integer iconId) {
    this(range, descending, "", color, iconId);
  }

  public AutofilterSortConditionInput {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(sortBy, "sortBy must not be null");
    if (iconId != null && iconId < 0) {
      throw new IllegalArgumentException("iconId must not be negative");
    }
  }

  @JsonCreator
  static AutofilterSortConditionInput create(
      @JsonProperty("range") String range,
      @JsonProperty("descending") boolean descending,
      @JsonProperty("sortBy") String sortBy,
      @JsonProperty("color") ColorInput color,
      @JsonProperty("iconId") Integer iconId) {
    return new AutofilterSortConditionInput(
        range, descending, sortBy == null ? "" : sortBy, color, iconId);
  }
}
