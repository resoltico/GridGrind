package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Error-box configuration attached to one data-validation rule. */
public record ExcelDataValidationErrorAlert(
    ExcelDataValidationErrorStyle style, String title, String text, boolean showErrorBox) {
  public ExcelDataValidationErrorAlert {
    Objects.requireNonNull(style, "style must not be null");
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
