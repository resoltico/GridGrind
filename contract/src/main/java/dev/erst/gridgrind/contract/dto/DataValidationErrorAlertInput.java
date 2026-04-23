package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Error-box configuration attached to one data-validation rule. */
public record DataValidationErrorAlertInput(
    ExcelDataValidationErrorStyle style,
    TextSourceInput title,
    TextSourceInput text,
    Boolean showErrorBox) {
  public DataValidationErrorAlertInput {
    Objects.requireNonNull(style, "style must not be null");
    title = requireNonBlankSource(title, "title");
    text = requireNonBlankSource(text, "text");
    showErrorBox = showErrorBox == null || showErrorBox;
  }

  private static TextSourceInput requireNonBlankSource(TextSourceInput value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
