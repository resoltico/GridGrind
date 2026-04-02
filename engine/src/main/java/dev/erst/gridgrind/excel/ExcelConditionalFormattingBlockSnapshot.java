package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Factual conditional-formatting block metadata loaded from one sheet. */
public record ExcelConditionalFormattingBlockSnapshot(
    List<String> ranges, List<ExcelConditionalFormattingRuleSnapshot> rules) {
  public ExcelConditionalFormattingBlockSnapshot {
    Objects.requireNonNull(ranges, "ranges must not be null");
    ranges = List.copyOf(ranges);
    for (String range : ranges) {
      if (range.isBlank()) {
        throw new IllegalArgumentException("ranges must not contain null or blank values");
      }
    }
    Objects.requireNonNull(rules, "rules must not be null");
    rules = List.copyOf(rules);
    for (ExcelConditionalFormattingRuleSnapshot rule : rules) {
      Objects.requireNonNull(rule, "rules must not contain null values");
    }
  }
}
