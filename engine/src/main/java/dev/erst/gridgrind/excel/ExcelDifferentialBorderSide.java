package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;

/** One differential-style border side used by conditional-formatting rules. */
public record ExcelDifferentialBorderSide(ExcelBorderStyle style, String color) {
  public ExcelDifferentialBorderSide {
    Objects.requireNonNull(style, "style must not be null");
    color = ExcelRgbColorSupport.normalizeRgbHex(color, "color").orElse(null);
  }
}
