package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Border patch used by conditional-formatting differential styles. */
public record ExcelDifferentialBorder(
    ExcelDifferentialBorderSide all,
    ExcelDifferentialBorderSide top,
    ExcelDifferentialBorderSide right,
    ExcelDifferentialBorderSide bottom,
    ExcelDifferentialBorderSide left) {
  public ExcelDifferentialBorder {
    if (java.util.stream.Stream.of(all, top, right, bottom, left).allMatch(Objects::isNull)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
