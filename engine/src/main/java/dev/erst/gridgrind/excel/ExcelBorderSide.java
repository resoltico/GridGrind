package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import java.util.Optional;

/** One side of a border patch or snapshot, defined by style and optional RGB color. */
public record ExcelBorderSide(Optional<ExcelBorderStyle> style, Optional<ExcelColor> color) {
  /** Creates a border side with the supplied style and no explicit RGB color override. */
  public ExcelBorderSide(ExcelBorderStyle style) {
    this(Optional.ofNullable(style), Optional.empty());
  }

  public ExcelBorderSide {
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
