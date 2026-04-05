package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import java.util.Objects;

/** Protocol-facing factual report for one advanced conditional-formatting threshold. */
public record ConditionalFormattingThresholdReport(
    ExcelConditionalFormattingThresholdType type, String formula, Double value) {
  public ConditionalFormattingThresholdReport {
    Objects.requireNonNull(type, "type must not be null");
    if (formula != null && formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
  }
}
