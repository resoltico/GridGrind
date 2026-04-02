package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** One authored conditional-formatting block with target ranges and ordered rule definitions. */
public record ExcelConditionalFormattingBlockDefinition(
    List<String> ranges, List<ExcelConditionalFormattingRule> rules) {
  public ExcelConditionalFormattingBlockDefinition {
    ranges = copyRanges(ranges);
    rules = copyRules(rules);
    if (rules.isEmpty()) {
      throw new IllegalArgumentException("rules must not be empty");
    }
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String range : copy) {
      String normalized = ExcelComparisonFormulaSupport.requireNonBlank(range, "ranges");
      if (!unique.add(normalized)) {
        throw new IllegalArgumentException("ranges must not contain duplicates");
      }
    }
    return copy;
  }

  private static List<ExcelConditionalFormattingRule> copyRules(
      List<ExcelConditionalFormattingRule> rules) {
    Objects.requireNonNull(rules, "rules must not be null");
    List<ExcelConditionalFormattingRule> copy = List.copyOf(rules);
    for (ExcelConditionalFormattingRule rule : copy) {
      Objects.requireNonNull(rule, "rules must not contain null values");
    }
    return copy;
  }
}
