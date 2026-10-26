package dev.erst.gridgrind.core.model.validation;

import java.util.Objects;

/** Materialized prompt text attached to a data-validation rule. */
public record ResolvedValidationPrompt(String title, String text, boolean showPromptBox) {
  public ResolvedValidationPrompt {
    title = requireNonBlank(title, "title");
    text = requireNonBlank(text, "text");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
