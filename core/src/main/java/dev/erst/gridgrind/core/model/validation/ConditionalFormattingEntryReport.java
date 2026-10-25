package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.List;
import java.util.Objects;

/** Protocol-facing factual report for one conditional-formatting block loaded from a sheet. */
public record ConditionalFormattingEntryReport(
    List<RangeLabel> ranges, List<ConditionalFormattingRuleReport> rules) {
  public ConditionalFormattingEntryReport {
    Objects.requireNonNull(ranges, "ranges must not be null");
    ranges = List.copyOf(ranges);
    Objects.requireNonNull(rules, "rules must not be null");
    rules = List.copyOf(rules);
    for (ConditionalFormattingRuleReport rule : rules) {
      Objects.requireNonNull(rule, "rules must not contain null values");
    }
  }
}
