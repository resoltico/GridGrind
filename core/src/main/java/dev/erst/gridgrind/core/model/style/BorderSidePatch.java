package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing definition for one border side shared by authored style patches. */
@CatalogType(
    id = "BorderSidePatch",
    summary =
        "One border side shared by cell and differential style patches, defined by its border style and optional color semantics.",
    displayOrder = 50)
public record BorderSidePatch(Optional<ExcelBorderStyle> style, Optional<ColorSpec> color) {
  /** Creates a protocol border side with the supplied style and no explicit RGB color. */
  public BorderSidePatch(ExcelBorderStyle style) {
    this(Optional.ofNullable(style), Optional.empty());
  }

  /** Creates a protocol border side with an explicit optional color payload. */
  public BorderSidePatch(ExcelBorderStyle style, ColorSpec color) {
    this(Optional.ofNullable(style), Optional.ofNullable(color));
  }

  public BorderSidePatch {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    if (style.isEmpty() && color.isEmpty()) {
      throw new IllegalArgumentException("border side must set style and/or color");
    }
    if (style.isPresent() && style.orElseThrow() == ExcelBorderStyle.NONE && color.isPresent()) {
      throw new IllegalArgumentException("border side color is not supported when style is NONE");
    }
  }
}
