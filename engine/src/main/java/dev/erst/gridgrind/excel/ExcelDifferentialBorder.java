package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Border patch used by conditional-formatting differential styles. */
public record ExcelDifferentialBorder(
    @Nullable ExcelBorderSide all,
    @Nullable ExcelBorderSide top,
    @Nullable ExcelBorderSide right,
    @Nullable ExcelBorderSide bottom,
    @Nullable ExcelBorderSide left) {
  public ExcelDifferentialBorder {
    if (java.util.stream.Stream.of(all, top, right, bottom, left).allMatch(Objects::isNull)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
