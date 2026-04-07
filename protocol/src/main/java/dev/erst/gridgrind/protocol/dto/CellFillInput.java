package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelFillPattern;

/** Protocol-facing fill patch used by {@link CellStyleInput}. */
public record CellFillInput(
    ExcelFillPattern pattern, String foregroundColor, String backgroundColor) {
  public CellFillInput {
    foregroundColor = ProtocolRgbColorSupport.normalizeRgbHex(foregroundColor, "foregroundColor");
    backgroundColor = ProtocolRgbColorSupport.normalizeRgbHex(backgroundColor, "backgroundColor");
    if (pattern == null && foregroundColor == null && backgroundColor == null) {
      throw new IllegalArgumentException("fill must set at least one attribute");
    }
    if (pattern == ExcelFillPattern.NONE && (foregroundColor != null || backgroundColor != null)) {
      throw new IllegalArgumentException("fill pattern NONE does not accept colors");
    }
    if (pattern == ExcelFillPattern.SOLID && backgroundColor != null) {
      throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
    }
    if (backgroundColor != null && pattern == null) {
      throw new IllegalArgumentException("fill backgroundColor requires an explicit pattern");
    }
  }
}
