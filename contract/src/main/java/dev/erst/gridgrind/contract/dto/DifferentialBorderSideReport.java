package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;

/** Protocol-facing factual report for one conditional-formatting differential border side. */
public record DifferentialBorderSideReport(ExcelBorderStyle style, String color) {
  public DifferentialBorderSideReport {
    Objects.requireNonNull(style, "style must not be null");
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
  }
}
