package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Materialized error-box text attached to a data-validation rule. */
public record ResolvedValidationErrorAlert(
    ExcelDataValidationErrorStyle style, String title, String text, boolean showErrorBox) {
  public ResolvedValidationErrorAlert {
    Objects.requireNonNull(style, "style must not be null");
    title = TextValidation.requireNonBlank(title, "title");
    text = TextValidation.requireNonBlank(text, "text");
  }
}
