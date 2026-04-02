package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Factual threshold metadata loaded from an advanced conditional-formatting rule. */
public record ExcelConditionalFormattingThresholdSnapshot(
    ExcelConditionalFormattingThresholdType type, String formula, Double value) {
  public ExcelConditionalFormattingThresholdSnapshot {
    Objects.requireNonNull(type, "type must not be null");
    if (formula != null && formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
  }
}
