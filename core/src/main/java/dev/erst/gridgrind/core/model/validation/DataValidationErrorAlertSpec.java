package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Error-box configuration attached to one data-validation rule. */
@CatalogType(
    id = "DataValidationErrorAlertSpec",
    summary = "Optional error-box configuration shown when invalid data is entered.",
    displayOrder = 54)
public record DataValidationErrorAlertSpec(
    ExcelDataValidationErrorStyle style,
    TextSourceSpec title,
    TextSourceSpec text,
    @WireDefault("true") boolean showErrorBox) {
  public DataValidationErrorAlertSpec {
    Objects.requireNonNull(style, "style must not be null");
    title = requireNonBlankSource(title, "title");
    text = requireNonBlankSource(text, "text");
  }

  private static TextSourceSpec requireNonBlankSource(TextSourceSpec value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value instanceof TextSourceSpec.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
