package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import java.util.Objects;

/** Protocol-facing definition for one conditional-formatting differential border side. */
public record DifferentialBorderSideInput(ExcelBorderStyle style, String color) {
  public DifferentialBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
  }

  /** Converts this transport model into the engine border-side model. */
  public ExcelDifferentialBorderSide toExcelDifferentialBorderSide() {
    return new ExcelDifferentialBorderSide(style, color);
  }
}
