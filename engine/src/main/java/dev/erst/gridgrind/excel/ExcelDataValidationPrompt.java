package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Prompt-box configuration attached to one data-validation rule. */
public record ExcelDataValidationPrompt(String title, String text, boolean showPromptBox) {
  public ExcelDataValidationPrompt {
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
