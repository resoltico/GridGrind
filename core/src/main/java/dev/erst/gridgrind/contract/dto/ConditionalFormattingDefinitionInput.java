package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Protocol-facing authored conditional-formatting definition with an ordered rule list. */
public record ConditionalFormattingDefinitionInput(List<ConditionalFormattingRuleInput> rules) {
  public ConditionalFormattingDefinitionInput {
    rules = copyRules(rules);
    if (rules.isEmpty()) {
      throw new IllegalArgumentException("rules must not be empty");
    }
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
