package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Materialized error-box text attached to a data-validation rule. */
public record ResolvedValidationErrorAlert(
    ExcelDataValidationErrorStyle style, String title, String text, boolean showErrorBox) {
  public ResolvedValidationErrorAlert {
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
