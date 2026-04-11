package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authored threshold used by advanced conditional-formatting rules. */
public record ExcelConditionalFormattingThreshold(
    ExcelConditionalFormattingThresholdType type, String formula, Double value) {
  public ExcelConditionalFormattingThreshold {
    Objects.requireNonNull(type, "type must not be null");
    if (formula != null && formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException("value must be finite");
    }
  }
}
