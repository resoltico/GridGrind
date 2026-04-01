package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
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

  /** Converts this transport model into the engine definition. */
  public ExcelDataValidationDefinition toExcelDataValidationDefinition() {
    return new ExcelDataValidationDefinition(
        rule.toExcelDataValidationRule(),
        allowBlank,
        suppressDropDownArrow,
        prompt == null ? null : prompt.toExcelDataValidationPrompt(),
        errorAlert == null ? null : errorAlert.toExcelDataValidationErrorAlert());
  }
}
