package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Prompt-box configuration attached to one data-validation rule. */
public record DataValidationPromptInput(
    TextSourceInput title, TextSourceInput text, boolean showPromptBox) {
  public DataValidationPromptInput {
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

  /** Creates a prompt-box payload while defaulting prompt visibility to Excel's enabled state. */
  @JsonCreator
  public DataValidationPromptInput(
      @JsonProperty("title") TextSourceInput title,
      @JsonProperty("text") TextSourceInput text,
      @JsonProperty("showPromptBox") Boolean showPromptBox) {
    this(
        title,
        text,
        Objects.requireNonNull(showPromptBox, "showPromptBox must not be null").booleanValue());
  }
}
