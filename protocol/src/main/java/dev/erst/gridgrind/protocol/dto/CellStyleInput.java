package dev.erst.gridgrind.protocol.dto;

/** Protocol-facing style patch used for range and cell presentation changes. */
public record CellStyleInput(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    Boolean wrapText,
    HorizontalAlignment horizontalAlignment,
    VerticalAlignment verticalAlignment,
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
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor");
    fillColor = ProtocolRgbColorSupport.normalizeRgbHex(fillColor, "fillColor");
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
}
