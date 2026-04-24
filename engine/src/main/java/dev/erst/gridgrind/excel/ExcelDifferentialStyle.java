package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authorable conditional-formatting differential style patch. */
public record ExcelDifferentialStyle(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    ExcelFontHeight fontHeight,
    String fontColor,
    Boolean underline,
    Boolean strikeout,
    String fillColor,
    ExcelDifferentialBorder border) {
  public ExcelDifferentialStyle {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    fontColor = ExcelRgbColorSupport.normalizeRgbHex(fontColor, "fontColor").orElse(null);
    fillColor = ExcelRgbColorSupport.normalizeRgbHex(fillColor, "fillColor").orElse(null);
    if (java.util.stream.Stream.of(
            numberFormat,
            bold,
            italic,
            fontHeight,
            fontColor,
            underline,
            strikeout,
            fillColor,
            border)
        .allMatch(Objects::isNull)) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }
}
