package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One side of a border patch, currently defined only by its line style. */
public record ExcelBorderSide(ExcelBorderStyle style) {
  public ExcelBorderSide {
    Objects.requireNonNull(style, "style must not be null");
  }
}
