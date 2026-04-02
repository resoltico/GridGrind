package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import java.util.Objects;
import java.util.stream.Stream;

/** Protocol-facing factual report for one conditional-formatting differential border patch. */
public record DifferentialBorderReport(
    DifferentialBorderSideReport all,
    DifferentialBorderSideReport top,
    DifferentialBorderSideReport right,
    DifferentialBorderSideReport bottom,
    DifferentialBorderSideReport left) {
  public DifferentialBorderReport {
    if (hasNoSides(all, top, right, bottom, left)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  /** Converts one engine differential border into the protocol report shape. */
  public static DifferentialBorderReport fromExcel(ExcelDifferentialBorder border) {
    if (border == null) {
      return null;
    }
    return new DifferentialBorderReport(
        toSideReport(border.all()),
        toSideReport(border.top()),
        toSideReport(border.right()),
        toSideReport(border.bottom()),
        toSideReport(border.left()));
  }

  private static boolean hasNoSides(Object... sides) {
    return Stream.of(sides).allMatch(Objects::isNull);
  }

  private static DifferentialBorderSideReport toSideReport(ExcelDifferentialBorderSide side) {
    return side == null ? null : DifferentialBorderSideReport.fromExcel(side);
  }
}
