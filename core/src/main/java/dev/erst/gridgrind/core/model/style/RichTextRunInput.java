package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.OoxmlTextValidation;
import dev.erst.gridgrind.core.source.TextSourceInput;
import java.util.Objects;
import java.util.Optional;

/** One ordered rich-text run authored for a string cell. */
@CatalogType(
    id = "RichTextRunInput",
    summary =
        "One ordered rich-text run for a string cell. text must be non-empty; font is an optional override patch. The ordered run texts concatenate to the stored plain string value.",
    displayOrder = 46)
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
    if (source instanceof TextSourceInput.Inline inline) {
      source =
          new TextSourceInput.Inline(
              OoxmlTextValidation.requireXml10CharacterData(inline.text(), "source.text"));
    }
  }
}
