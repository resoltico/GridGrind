package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Protocol-facing authored conditional-formatting block with ordered rules and target ranges. */
public record ConditionalFormattingBlockInput(
    List<String> ranges, List<ConditionalFormattingRuleInput> rules) {
  public ConditionalFormattingBlockInput {
    ranges = copyRanges(ranges);
    rules = copyRules(rules);
    if (rules.isEmpty()) {
      throw new IllegalArgumentException("rules must not be empty");
    }
  }

  /** Converts this transport model into the engine block model. */
  public ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlockDefinition() {
    return new ExcelConditionalFormattingBlockDefinition(
        ranges,
        rules.stream()
            .map(ConditionalFormattingRuleInput::toExcelConditionalFormattingRule)
            .toList());
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String range : copy) {
      Objects.requireNonNull(range, "ranges must not contain null values");
      if (range.isBlank()) {
        throw new IllegalArgumentException("ranges must not contain blank values");
      }
      if (!unique.add(range)) {
        throw new IllegalArgumentException("ranges must not contain duplicates");
      }
    }
    return copy;
  }

  private static List<ConditionalFormattingRuleInput> copyRules(
      List<ConditionalFormattingRuleInput> rules) {
    Objects.requireNonNull(rules, "rules must not be null");
    List<ConditionalFormattingRuleInput> copy = List.copyOf(rules);
    for (ConditionalFormattingRuleInput rule : copy) {
      Objects.requireNonNull(rule, "rules must not contain null values");
    }
    return copy;
  }
}
