package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Observed error-box text, style and display state, without authored source references. */
@CatalogType(
    id = "DataValidationErrorAlertReport",
    summary = "Observed error-box text, style and display state.",
    displayOrder = 109)
public record DataValidationErrorAlertReport(
    ExcelDataValidationErrorStyle style, String title, String text, boolean showErrorBox) {
  public DataValidationErrorAlertReport {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(text, "text must not be null");
  }
}
