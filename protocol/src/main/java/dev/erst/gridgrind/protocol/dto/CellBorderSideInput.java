package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import java.util.Objects;

/** Protocol-facing definition for one border side within a style patch. */
public record CellBorderSideInput(ExcelBorderStyle style) {
  public CellBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
  }
}
