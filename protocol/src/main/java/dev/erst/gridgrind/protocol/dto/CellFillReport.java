package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelFillPattern;
import java.util.Objects;

/** Effective fill facts reported with every analyzed cell. */
public record CellFillReport(
    ExcelFillPattern pattern, String foregroundColor, String backgroundColor) {
  public CellFillReport {
    Objects.requireNonNull(pattern, "pattern must not be null");
    foregroundColor = ProtocolRgbColorSupport.normalizeRgbHex(foregroundColor, "foregroundColor");
    backgroundColor = ProtocolRgbColorSupport.normalizeRgbHex(backgroundColor, "backgroundColor");
    if (pattern == ExcelFillPattern.NONE && (foregroundColor != null || backgroundColor != null)) {
      throw new IllegalArgumentException("fill pattern NONE does not accept colors");
    }
    if (pattern == ExcelFillPattern.SOLID && backgroundColor != null) {
      throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
    }
  }
}
