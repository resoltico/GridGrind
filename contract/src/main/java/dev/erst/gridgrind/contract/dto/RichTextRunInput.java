package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** One ordered rich-text run authored for a string cell. */
public record RichTextRunInput(TextSourceInput source, CellFontInput font) {
  public RichTextRunInput {
    Objects.requireNonNull(source, "source must not be null");
    if (source instanceof TextSourceInput.Inline inline && inline.text().isEmpty()) {
      throw new IllegalArgumentException("source must not be empty");
    }
  }
}
