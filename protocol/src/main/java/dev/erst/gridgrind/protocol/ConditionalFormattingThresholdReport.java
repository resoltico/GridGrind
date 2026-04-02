package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
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

  /** Converts one engine threshold snapshot into the protocol report shape. */
  public static ConditionalFormattingThresholdReport fromExcel(
      ExcelConditionalFormattingThresholdSnapshot threshold) {
    Objects.requireNonNull(threshold, "threshold must not be null");
    return new ConditionalFormattingThresholdReport(
        threshold.type(), threshold.formula(), threshold.value());
  }
}
