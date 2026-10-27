package dev.erst.gridgrind.core.model.validation;

import java.util.Objects;
import java.util.Optional;

/** Data-validation settings with materialized prompt and error text. */
public record ResolvedDataValidation(
    DataValidationRuleSpec rule,
    boolean allowBlank,
    boolean suppressDropDownArrow,
    Optional<ResolvedValidationPrompt> prompt,
    Optional<ResolvedValidationErrorAlert> errorAlert) {
  public ResolvedDataValidation {
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(prompt, "prompt must not be null");
    Objects.requireNonNull(errorAlert, "errorAlert must not be null");
  }
}
