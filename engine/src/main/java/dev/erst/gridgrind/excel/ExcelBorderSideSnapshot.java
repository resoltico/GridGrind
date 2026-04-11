package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One resolved workbook border side preserving style and factual color semantics. */
public record ExcelBorderSideSnapshot(ExcelBorderStyle style, ExcelColorSnapshot color) {
  public ExcelBorderSideSnapshot {
    Objects.requireNonNull(style, "style must not be null");
  }
}
