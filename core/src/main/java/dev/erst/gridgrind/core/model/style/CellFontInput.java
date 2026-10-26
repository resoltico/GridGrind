package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing font patch used by {@link CellStylePatchInput}. */
@CatalogType(
    id = "CellFontInput",
    summary =
        "Font patch for cell styling; at least one field must be set. Colors can use RGB, theme, indexed, and tint semantics.",
    displayOrder = 45)
public record CellFontInput(
    Optional<Boolean> bold,
    Optional<Boolean> italic,
    Optional<String> fontName,
    Optional<FontHeightInput> fontHeight,
    Optional<ColorInput> fontColor,
    Optional<Boolean> underline,
    Optional<Boolean> strikeout) {
  public CellFontInput {
    Objects.requireNonNull(bold, "bold must not be null");
    Objects.requireNonNull(italic, "italic must not be null");
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    Objects.requireNonNull(fontColor, "fontColor must not be null");
    Objects.requireNonNull(underline, "underline must not be null");
    Objects.requireNonNull(strikeout, "strikeout must not be null");
    fontName.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("fontName must not be blank");
          }
        });
    if (bold.isEmpty()
        && italic.isEmpty()
        && fontName.isEmpty()
        && fontHeight.isEmpty()
        && fontColor.isEmpty()
        && underline.isEmpty()
        && strikeout.isEmpty()) {
      throw new IllegalArgumentException("font must set at least one attribute");
    }
  }
}
