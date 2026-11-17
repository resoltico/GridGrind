package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** One primitive text invariant shared by validated values and non-identifying text fields. */
public final class TextValidation {
  private TextValidation() {}

  /** Returns the exact text, rejecting null or Unicode-blank values with field-specific wording. */
  public static String requireNonBlank(@Nullable String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      var problem =
          fieldName != null && !fieldName.isBlank()
              ? dev.erst.gridgrind.core.problem.FieldValidationProblem.atField(
                  fieldName, dev.erst.gridgrind.core.problem.FieldValidationBasicRule.NON_BLANK)
              : new dev.erst.gridgrind.core.problem.MessageInvariant(
                  "Value must not be blank.", java.util.Optional.empty());
      throw new dev.erst.gridgrind.core.problem.ValueInvariantException(problem);
    }
    return value;
  }
}
