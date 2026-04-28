package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;

/** Protocol-facing definition for one border side within a style patch. */
public record CellBorderSideInput(ExcelBorderStyle style, ColorInput color) {
  /** Creates a protocol border side with the supplied style and no explicit RGB color. */
  public CellBorderSideInput(ExcelBorderStyle style) {
    this(style, (ColorInput) null);
  }

  public CellBorderSideInput {
    if (style == null && color == null) {
      throw new IllegalArgumentException("border side must set style and/or color");
    }
    if (style == ExcelBorderStyle.NONE && color != null) {
      throw new IllegalArgumentException("border side color is not supported when style is NONE");
    }
  }
}
