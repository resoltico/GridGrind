package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import java.util.Objects;

/** Protocol-facing factual report for one conditional-formatting differential border side. */
public record DifferentialBorderSideReport(ExcelBorderStyle style, String color) {
  public DifferentialBorderSideReport {
    Objects.requireNonNull(style, "style must not be null");
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
  }

  /** Converts one engine border-side snapshot into the protocol report shape. */
  public static DifferentialBorderSideReport fromExcel(ExcelDifferentialBorderSide side) {
    Objects.requireNonNull(side, "side must not be null");
    return new DifferentialBorderSideReport(side.style(), side.color());
  }
}
