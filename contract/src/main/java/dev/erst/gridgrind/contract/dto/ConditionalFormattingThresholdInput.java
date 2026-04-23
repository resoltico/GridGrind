package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import java.util.Objects;

/** Threshold payload shared by authored advanced conditional-formatting rules. */
public record ConditionalFormattingThresholdInput(
    ExcelConditionalFormattingThresholdType type, String formula, Double value) {
  public ConditionalFormattingThresholdInput {
    Objects.requireNonNull(type, "type must not be null");
    if (formula != null && formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException("value must be finite");
    }
  }
}
