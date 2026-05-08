package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;
import java.util.Optional;

/** One ordered rich-text run authored for a string cell. */
public record RichTextRunInput(
    TextSourceInput source,
    @com.fasterxml.jackson.annotation.JsonInclude(
            com.fasterxml.jackson.annotation.JsonInclude.Include.NON_ABSENT)
        Optional<CellFontInput> font) {
  public RichTextRunInput {
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(font, "font must not be null");
    if (source instanceof TextSourceInput.Inline inline && inline.text().isEmpty()) {
      throw new IllegalArgumentException("source must not be empty");
    }
  }
}
