package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the fill currently applied to a cell. */
public record ExcelCellFillSnapshot(
    ExcelFillPattern pattern, String foregroundColor, String backgroundColor) {
  public ExcelCellFillSnapshot {
    Objects.requireNonNull(pattern, "pattern must not be null");
    foregroundColor = ExcelRgbColorSupport.normalizeRgbHex(foregroundColor, "foregroundColor");
    backgroundColor = ExcelRgbColorSupport.normalizeRgbHex(backgroundColor, "backgroundColor");
    if (pattern == ExcelFillPattern.NONE && (foregroundColor != null || backgroundColor != null)) {
      throw new IllegalArgumentException("fill pattern NONE does not accept colors");
    }
    if (pattern == ExcelFillPattern.SOLID && backgroundColor != null) {
      throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
    }
  }
}
