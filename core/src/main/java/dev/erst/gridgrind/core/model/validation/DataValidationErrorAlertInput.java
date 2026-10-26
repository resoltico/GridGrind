package dev.erst.gridgrind.core.model.validation;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Error-box configuration attached to one data-validation rule. */
@CatalogType(
    id = "DataValidationErrorAlertInput",
    summary = "Optional error-box configuration shown when invalid data is entered.",
    displayOrder = 54)
public record DataValidationErrorAlertInput(
    ExcelDataValidationErrorStyle style,
    TextSourceInput title,
    TextSourceInput text,
    boolean showErrorBox) {
  public DataValidationErrorAlertInput {
    Objects.requireNonNull(style, "style must not be null");
    title = requireNonBlankSource(title, "title");
    text = requireNonBlankSource(text, "text");
  }

  private static TextSourceInput requireNonBlankSource(TextSourceInput value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  /** Creates an error-box payload while defaulting error visibility to Excel's enabled state. */
  @JsonCreator
  public DataValidationErrorAlertInput(
      @JsonProperty("style") ExcelDataValidationErrorStyle style,
      @JsonProperty("title") TextSourceInput title,
      @JsonProperty("text") TextSourceInput text,
      @JsonProperty("showErrorBox") Boolean showErrorBox) {
    this(style, title, text, !Boolean.FALSE.equals(showErrorBox));
  }
}
