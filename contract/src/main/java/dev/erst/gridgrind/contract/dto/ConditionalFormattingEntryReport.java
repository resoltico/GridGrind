package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Protocol-facing factual report for one conditional-formatting block loaded from a sheet. */
public record ConditionalFormattingEntryReport(
    List<String> ranges, List<ConditionalFormattingRuleReport> rules) {
  public ConditionalFormattingEntryReport {
    Objects.requireNonNull(ranges, "ranges must not be null");
    ranges = List.copyOf(ranges);
    for (String range : ranges) {
      Objects.requireNonNull(range, "ranges must not contain null values");
      if (range.isBlank()) {
        throw new IllegalArgumentException("ranges must not contain blank values");
      }
    }
    Objects.requireNonNull(rules, "rules must not be null");
    rules = List.copyOf(rules);
    for (ConditionalFormattingRuleReport rule : rules) {
      Objects.requireNonNull(rule, "rules must not contain null values");
    }
  }
}
