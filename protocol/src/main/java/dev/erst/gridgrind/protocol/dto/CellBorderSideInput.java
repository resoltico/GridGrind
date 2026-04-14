package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelBorderStyle;

/** Protocol-facing definition for one border side within a style patch. */
public record CellBorderSideInput(
    ExcelBorderStyle style,
    String color,
    Integer colorTheme,
    Integer colorIndexed,
    Double colorTint) {
  /** Creates a protocol border side with the supplied style and no explicit RGB color. */
  public CellBorderSideInput(ExcelBorderStyle style) {
    this(style, null, null, null, null);
  }

  /** Creates a protocol border side with the supplied style and explicit RGB color. */
  public CellBorderSideInput(ExcelBorderStyle style, String color) {
    this(style, color, null, null, null);
  }

  public CellBorderSideInput {
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
    if (colorTheme != null && colorTheme < 0) {
      throw new IllegalArgumentException("colorTheme must not be negative");
    }
    if (colorIndexed != null && colorIndexed < 0) {
      throw new IllegalArgumentException("colorIndexed must not be negative");
    }
    if (colorTint != null && !Double.isFinite(colorTint)) {
      throw new IllegalArgumentException("colorTint must be finite");
    }
    if (color == null && colorTheme == null && colorIndexed == null && colorTint != null) {
      throw new IllegalArgumentException("colorTint requires color, colorTheme, or colorIndexed");
    }
    if (style == null && color == null && colorTheme == null && colorIndexed == null) {
      throw new IllegalArgumentException("border side must set style and/or color");
    }
    if (style == ExcelBorderStyle.NONE
        && (color != null || colorTheme != null || colorIndexed != null)) {
      throw new IllegalArgumentException("border side color is not supported when style is NONE");
    }
  }
}
