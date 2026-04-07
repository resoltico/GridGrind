package dev.erst.gridgrind.excel;

/** One side of a border patch or snapshot, defined by style and optional RGB color. */
public record ExcelBorderSide(ExcelBorderStyle style, String color) {
  /** Creates a border side with the supplied style and no explicit RGB color override. */
  public ExcelBorderSide(ExcelBorderStyle style) {
    this(style, null);
  }

  public ExcelBorderSide {
    color = ExcelRgbColorSupport.normalizeRgbHex(color, "color");
    if (style == null && color == null) {
      throw new IllegalArgumentException("border side must set style and/or color");
    }
    if (style == ExcelBorderStyle.NONE && color != null) {
      throw new IllegalArgumentException("border side color is not supported when style is NONE");
    }
  }
}
