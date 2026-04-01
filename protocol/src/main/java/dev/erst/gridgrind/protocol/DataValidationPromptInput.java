package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import java.util.Objects;

/** Prompt-box configuration attached to one data-validation rule. */
public record DataValidationPromptInput(String title, String text, Boolean showPromptBox) {
  public DataValidationPromptInput {
    title = requireNonBlank(title, "title");
    text = requireNonBlank(text, "text");
    showPromptBox = showPromptBox == null || showPromptBox;
  }

  /** Converts this transport model into the engine model. */
  public ExcelDataValidationPrompt toExcelDataValidationPrompt() {
    return new ExcelDataValidationPrompt(title, text, showPromptBox);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
