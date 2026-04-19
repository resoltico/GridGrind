package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Complete protocol-facing definition of one supported data-validation rule. */
public record DataValidationInput(
    DataValidationRuleInput rule,
    Boolean allowBlank,
    Boolean suppressDropDownArrow,
    DataValidationPromptInput prompt,
    DataValidationErrorAlertInput errorAlert) {
  public DataValidationInput {
    Objects.requireNonNull(rule, "rule must not be null");
    allowBlank = Boolean.TRUE.equals(allowBlank);
    suppressDropDownArrow = Boolean.TRUE.equals(suppressDropDownArrow);
  }
}
