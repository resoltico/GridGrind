package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.Objects;

/** Error-box configuration attached to one data-validation rule. */
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
    this(
        style,
        title,
        text,
        Objects.requireNonNull(showErrorBox, "showErrorBox must not be null").booleanValue());
  }
}
