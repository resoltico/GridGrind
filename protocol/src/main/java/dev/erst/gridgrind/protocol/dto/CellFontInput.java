package dev.erst.gridgrind.protocol.dto;

/** Protocol-facing font patch used by {@link CellStyleInput}. */
public record CellFontInput(
    Boolean bold,
    Boolean italic,
    String fontName,
    FontHeightInput fontHeight,
    String fontColor,
    Boolean underline,
    Boolean strikeout) {
  public CellFontInput {
    if (fontName != null && fontName.isBlank()) {
      throw new IllegalArgumentException("fontName must not be blank");
    }
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor");
    if (bold == null
        && italic == null
        && fontName == null
        && fontHeight == null
        && fontColor == null
        && underline == null
        && strikeout == null) {
      throw new IllegalArgumentException("font must set at least one attribute");
    }
  }
}
