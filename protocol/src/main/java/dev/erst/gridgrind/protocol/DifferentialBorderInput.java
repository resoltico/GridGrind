package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import java.util.Objects;
import java.util.stream.Stream;

/** Protocol-facing differential border patch used by conditional-formatting rule styles. */
public record DifferentialBorderInput(
    DifferentialBorderSideInput all,
    DifferentialBorderSideInput top,
    DifferentialBorderSideInput right,
    DifferentialBorderSideInput bottom,
    DifferentialBorderSideInput left) {
  public DifferentialBorderInput {
    if (hasNoSides(all, top, right, bottom, left)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  /** Converts this transport model into the engine border model. */
  public ExcelDifferentialBorder toExcelDifferentialBorder() {
    return new ExcelDifferentialBorder(
        toSide(all), toSide(top), toSide(right), toSide(bottom), toSide(left));
  }

  private static boolean hasNoSides(Object... sides) {
    return Stream.of(sides).allMatch(Objects::isNull);
  }

  private static ExcelDifferentialBorderSide toSide(DifferentialBorderSideInput side) {
    return side == null ? null : side.toExcelDifferentialBorderSide();
  }
}
