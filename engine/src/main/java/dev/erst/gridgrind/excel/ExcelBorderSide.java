package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;

/** One side of a border patch or snapshot, defined by style and optional RGB color. */
public record ExcelBorderSide(ExcelBorderStyle style, ExcelColor color) {
  /** Creates a border side with the supplied style and no explicit RGB color override. */
  public ExcelBorderSide(ExcelBorderStyle style) {
    this(style, (ExcelColor) null);
  }

  public ExcelBorderSide {
    if (style == null && color == null) {
      throw new IllegalArgumentException("border side must set style and/or color");
    }
    if (style == ExcelBorderStyle.NONE && color != null) {
      throw new IllegalArgumentException("border side color is not supported when style is NONE");
    }
  }
}
