package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Complete GridGrind-owned definition of one supported data-validation rule. */
public record ExcelDataValidationDefinition(
    ExcelDataValidationRule rule,
    boolean allowBlank,
    boolean suppressDropDownArrow,
    Optional<ExcelDataValidationPrompt> prompt,
    Optional<ExcelDataValidationErrorAlert> errorAlert) {
  public ExcelDataValidationDefinition {
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(prompt, "prompt must not be null");
    Objects.requireNonNull(errorAlert, "errorAlert must not be null");
  }
}
