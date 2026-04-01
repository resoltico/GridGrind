package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Complete GridGrind-owned definition of one supported data-validation rule. */
public record ExcelDataValidationDefinition(
    ExcelDataValidationRule rule,
    boolean allowBlank,
    boolean suppressDropDownArrow,
    ExcelDataValidationPrompt prompt,
    ExcelDataValidationErrorAlert errorAlert) {
  public ExcelDataValidationDefinition {
    Objects.requireNonNull(rule, "rule must not be null");
  }
}
