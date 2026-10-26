package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import java.util.Objects;
import java.util.Optional;

/** Font patch applied through {@link ExcelCellStyle}. */
public record ExcelCellFont(
    Optional<Boolean> bold,
    Optional<Boolean> italic,
    Optional<String> fontName,
    Optional<FontHeightReport> fontHeight,
    Optional<ColorSpec> fontColor,
    Optional<Boolean> underline,
    Optional<Boolean> strikeout) {
  public ExcelCellFont {
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
