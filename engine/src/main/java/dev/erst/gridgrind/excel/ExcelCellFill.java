package dev.erst.gridgrind.excel;

/** Fill patch applied through {@link ExcelCellStyle}. */
public record ExcelCellFill(
    ExcelFillPattern pattern, String foregroundColor, String backgroundColor) {
  public ExcelCellFill {
    foregroundColor = ExcelRgbColorSupport.normalizeRgbHex(foregroundColor, "foregroundColor");
    backgroundColor = ExcelRgbColorSupport.normalizeRgbHex(backgroundColor, "backgroundColor");
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
