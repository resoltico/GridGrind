package dev.erst.gridgrind.core.model.style;

import java.util.Objects;
import java.util.Optional;

/** One materialized rich-text run with a canonical font patch. */
public record ResolvedRichTextRun(String text, Optional<CellFontPatch> font) {
  public ResolvedRichTextRun {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(font, "font must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
