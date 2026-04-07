package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** One ordered rich-text run authored for a string cell. */
public record RichTextRunInput(String text, CellFontInput font) {
  public RichTextRunInput {
    Objects.requireNonNull(text, "text must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
