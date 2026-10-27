package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Materialized prompt text attached to a data-validation rule. */
public record ResolvedValidationPrompt(String title, String text, boolean showPromptBox) {
  public ResolvedValidationPrompt {
    title = TextValidation.requireNonBlank(title, "title");
    text = TextValidation.requireNonBlank(text, "text");
  }
}
