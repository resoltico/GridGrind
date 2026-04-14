package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelFillPattern;

/** Protocol-facing fill patch used by {@link CellStyleInput}. */
public record CellFillInput(
    ExcelFillPattern pattern,
    String foregroundColor,
    Integer foregroundColorTheme,
    Integer foregroundColorIndexed,
    Double foregroundColorTint,
    String backgroundColor,
    Integer backgroundColorTheme,
    Integer backgroundColorIndexed,
    Double backgroundColorTint,
    CellGradientFillInput gradient) {
  /** Creates a pattern fill from simple RGB foreground/background values. */
  public CellFillInput(ExcelFillPattern pattern, String foregroundColor, String backgroundColor) {
    this(pattern, foregroundColor, null, null, null, backgroundColor, null, null, null, null);
  }

  public CellFillInput {
    foregroundColor = ProtocolRgbColorSupport.normalizeRgbHex(foregroundColor, "foregroundColor");
    backgroundColor = ProtocolRgbColorSupport.normalizeRgbHex(backgroundColor, "backgroundColor");
    requireColorBaseOrNull(
        foregroundColor,
        foregroundColorTheme,
        foregroundColorIndexed,
        foregroundColorTint,
        "foregroundColor");
    requireColorBaseOrNull(
        backgroundColor,
        backgroundColorTheme,
        backgroundColorIndexed,
        backgroundColorTint,
        "backgroundColor");
    if (pattern == null
        && foregroundColor == null
        && foregroundColorTheme == null
        && foregroundColorIndexed == null
        && backgroundColor == null
        && backgroundColorTheme == null
        && backgroundColorIndexed == null
        && gradient == null) {
      throw new IllegalArgumentException("fill must set at least one attribute");
    }
    if (gradient != null
        && (pattern != null
            || foregroundColor != null
            || foregroundColorTheme != null
            || foregroundColorIndexed != null
            || backgroundColor != null
            || backgroundColorTheme != null
            || backgroundColorIndexed != null)) {
      throw new IllegalArgumentException(
          "gradient fills must not also set pattern, foregroundColor, or backgroundColor");
    }
    if (pattern == ExcelFillPattern.NONE && (foregroundColor != null || backgroundColor != null)) {
      throw new IllegalArgumentException("fill pattern NONE does not accept colors");
    }
    if (pattern == ExcelFillPattern.NONE
        && (foregroundColorTheme != null
            || foregroundColorIndexed != null
            || backgroundColorTheme != null
            || backgroundColorIndexed != null)) {
      throw new IllegalArgumentException("fill pattern NONE does not accept colors");
    }
    if (pattern == ExcelFillPattern.SOLID
        && (backgroundColor != null
            || backgroundColorTheme != null
            || backgroundColorIndexed != null)) {
      throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
    }
    if ((backgroundColor != null || backgroundColorTheme != null || backgroundColorIndexed != null)
        && pattern == null) {
      throw new IllegalArgumentException("fill backgroundColor requires an explicit pattern");
    }
  }

  private static void requireColorBaseOrNull(
      String rgb, Integer theme, Integer indexed, Double tint, String fieldName) {
    if (theme != null && theme < 0) {
      throw new IllegalArgumentException(fieldName + "Theme must not be negative");
    }
    if (indexed != null && indexed < 0) {
      throw new IllegalArgumentException(fieldName + "Indexed must not be negative");
    }
    if (tint != null && !Double.isFinite(tint)) {
      throw new IllegalArgumentException(fieldName + "Tint must be finite");
    }
    if (rgb == null && theme == null && indexed == null && tint != null) {
      throw new IllegalArgumentException(
          fieldName
              + "Tint requires "
              + fieldName
              + ", "
              + fieldName
              + "Theme, or "
              + fieldName
              + "Indexed");
    }
  }
}
