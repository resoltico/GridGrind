package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Error-box configuration attached to one data-validation rule. */
public record DataValidationErrorAlertInput(
    ExcelDataValidationErrorStyle style, String title, String text, Boolean showErrorBox) {
  public DataValidationErrorAlertInput {
    Objects.requireNonNull(style, "style must not be null");
    title = requireNonBlank(title, "title");
    text = requireNonBlank(text, "text");
    showErrorBox = showErrorBox == null || showErrorBox;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
