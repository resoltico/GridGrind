package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Prompt-box configuration attached to one data-validation rule. */
public record DataValidationPromptInput(String title, String text, Boolean showPromptBox) {
  public DataValidationPromptInput {
    title = requireNonBlank(title, "title");
    text = requireNonBlank(text, "text");
    showPromptBox = showPromptBox == null || showPromptBox;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
