package dev.erst.gridgrind.contract.dto;

/** Protocol-facing font patch used by {@link CellStyleInput}. */
public record CellFontInput(
    Boolean bold,
    Boolean italic,
    String fontName,
    FontHeightInput fontHeight,
    String fontColor,
    Integer fontColorTheme,
    Integer fontColorIndexed,
    Double fontColorTint,
    Boolean underline,
    Boolean strikeout) {
  /** Creates a font patch using only the legacy RGB color shape. */
  public CellFontInput(
      Boolean bold,
      Boolean italic,
      String fontName,
      FontHeightInput fontHeight,
      String fontColor,
      Boolean underline,
      Boolean strikeout) {
    this(bold, italic, fontName, fontHeight, fontColor, null, null, null, underline, strikeout);
  }

  public CellFontInput {
    if (fontName != null && fontName.isBlank()) {
      throw new IllegalArgumentException("fontName must not be blank");
    }
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor").orElse(null);
    if (fontColorTheme != null && fontColorTheme < 0) {
      throw new IllegalArgumentException("fontColorTheme must not be negative");
    }
    if (fontColorIndexed != null && fontColorIndexed < 0) {
      throw new IllegalArgumentException("fontColorIndexed must not be negative");
    }
    if (fontColorTint != null && !Double.isFinite(fontColorTint)) {
      throw new IllegalArgumentException("fontColorTint must be finite");
    }
    if (fontColor == null
        && fontColorTheme == null
        && fontColorIndexed == null
        && fontColorTint != null) {
      throw new IllegalArgumentException(
          "fontColorTint requires fontColor, fontColorTheme, or fontColorIndexed");
    }
    if (bold == null
        && italic == null
        && fontName == null
        && fontHeight == null
        && fontColor == null
        && fontColorTheme == null
        && fontColorIndexed == null
        && underline == null
        && strikeout == null) {
      throw new IllegalArgumentException("font must set at least one attribute");
    }
  }
}
