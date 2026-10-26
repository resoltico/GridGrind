package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.source.TextSourceInput;
import java.util.Objects;

/** Prompt-box configuration attached to one data-validation rule. */
@CatalogType(
    id = "DataValidationPromptInput",
    summary = "Optional prompt-box configuration shown when a validated cell is selected.",
    displayOrder = 53)
public record DataValidationPromptInput(
    TextSourceInput title, TextSourceInput text, @WireDefault("true") boolean showPromptBox) {
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
}
