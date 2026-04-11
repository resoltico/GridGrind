package dev.erst.gridgrind.excel;

/** Fill patch applied through {@link ExcelCellStyle}. */
public record ExcelCellFill(
    ExcelFillPattern pattern,
    ExcelColor foregroundColor,
    ExcelColor backgroundColor,
    ExcelGradientFill gradient) {
  /** Creates a pattern fill from simple RGB foreground/background values. */
  public ExcelCellFill(ExcelFillPattern pattern, String foregroundColor, String backgroundColor) {
    this(
        pattern,
        foregroundColor == null ? null : new ExcelColor(foregroundColor),
        backgroundColor == null ? null : new ExcelColor(backgroundColor),
        null);
  }

  public ExcelCellFill {
    if (pattern == null && foregroundColor == null && backgroundColor == null && gradient == null) {
      throw new IllegalArgumentException("fill must set at least one attribute");
    }
    if (gradient != null
        && (pattern != null || foregroundColor != null || backgroundColor != null)) {
      throw new IllegalArgumentException(
          "gradient fills must not also set pattern, foregroundColor, or backgroundColor");
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
