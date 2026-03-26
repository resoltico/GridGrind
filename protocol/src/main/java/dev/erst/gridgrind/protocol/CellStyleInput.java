package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import java.util.Locale;

/** Protocol-facing style patch used for range and cell presentation changes. */
public record CellStyleInput(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    Boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment,
    String fontName,
    FontHeightInput fontHeight,
    String fontColor,
    Boolean underline,
    Boolean strikeout,
    String fillColor,
    CellBorderInput border) {
  public CellStyleInput {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    if (fontName != null && fontName.isBlank()) {
      throw new IllegalArgumentException("fontName must not be blank");
    }
    fontColor = normalizeRgbHex(fontColor, "fontColor");
    fillColor = normalizeRgbHex(fillColor, "fillColor");
    if (numberFormat == null
        && bold == null
        && italic == null
        && wrapText == null
        && horizontalAlignment == null
        && verticalAlignment == null
        && fontName == null
        && fontHeight == null
        && fontColor == null
        && underline == null
        && strikeout == null
        && fillColor == null
        && border == null) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }

  /** Converts this transport model into the engine style model. */
  public ExcelCellStyle toExcelCellStyle() {
    return new ExcelCellStyle(
        numberFormat,
        bold,
        italic,
        wrapText,
        horizontalAlignment,
        verticalAlignment,
        fontName,
        fontHeight == null ? null : fontHeight.toExcelFontHeight(),
        fontColor,
        underline,
        strikeout,
        fillColor,
        border == null ? null : border.toExcelBorder());
  }

  private static String normalizeRgbHex(String color, String fieldName) {
    if (color == null) {
      return null;
    }
    if (color.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    if (!color.matches("^#[0-9A-Fa-f]{6}$")) {
      throw new IllegalArgumentException(fieldName + " must match #RRGGBB");
    }
    return color.toUpperCase(Locale.ROOT);
  }
}
