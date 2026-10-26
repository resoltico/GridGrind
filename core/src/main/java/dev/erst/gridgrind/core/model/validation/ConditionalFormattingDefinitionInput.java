package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.List;
import java.util.Objects;

/** Protocol-facing authored conditional-formatting definition with an ordered rule list. */
@CatalogType(
    id = "ConditionalFormattingDefinitionInput",
    summary =
        "One authored conditional-formatting definition with an ordered rule list. Target ranges are owned by the selector; rules must not be empty.",
    displayOrder = 58)
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
