package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;

/** Factual report of one rich-text run stored in a string cell. */
@CatalogType(
    id = "RichTextRunReport",
    summary = "One factual rich-text run reported from a text cell or comment.",
    displayOrder = 86)
public record RichTextRunReport(String text, CellFontReport font) {
  public RichTextRunReport {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(font, "font must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
