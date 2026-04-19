package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual report of one rich-text run stored in a string cell. */
public record RichTextRunReport(String text, CellFontReport font) {
  public RichTextRunReport {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(font, "font must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
