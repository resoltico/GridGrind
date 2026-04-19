package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Prompt-box configuration attached to one data-validation rule. */
public record DataValidationPromptInput(
    TextSourceInput title, TextSourceInput text, Boolean showPromptBox) {
  public DataValidationPromptInput {
    title = requireNonBlankSource(title, "title");
    text = requireNonBlankSource(text, "text");
    showPromptBox = showPromptBox == null || showPromptBox;
  }

  private static TextSourceInput requireNonBlankSource(TextSourceInput value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
