package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
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

  /** Converts one engine conditional-formatting block snapshot into the protocol report shape. */
  public static ConditionalFormattingEntryReport fromExcel(
      ExcelConditionalFormattingBlockSnapshot block) {
    Objects.requireNonNull(block, "block must not be null");
    return new ConditionalFormattingEntryReport(
        block.ranges(),
        block.rules().stream().map(ConditionalFormattingRuleReport::fromExcel).toList());
  }
}
