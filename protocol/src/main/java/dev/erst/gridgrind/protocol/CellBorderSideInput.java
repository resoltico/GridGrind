package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import java.util.Objects;

/** Protocol-facing definition for one border side within a style patch. */
public record CellBorderSideInput(ExcelBorderStyle style) {
  public CellBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
  }

  /** Converts this transport model into the engine border-side model. */
  public ExcelBorderSide toExcelBorderSide() {
    return new ExcelBorderSide(style);
  }
}
