package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import java.util.Objects;

/** Protocol-facing definition for one conditional-formatting differential border side. */
public record DifferentialBorderSideInput(ExcelBorderStyle style, String color) {
  public DifferentialBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
  }
}
