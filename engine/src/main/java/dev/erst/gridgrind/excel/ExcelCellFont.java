package dev.erst.gridgrind.excel;

/** Font patch applied through {@link ExcelCellStyle}. */
public record ExcelCellFont(
    Boolean bold,
    Boolean italic,
    String fontName,
    ExcelFontHeight fontHeight,
    String fontColor,
    Boolean underline,
    Boolean strikeout) {
  public ExcelCellFont {
    if (fontName != null && fontName.isBlank()) {
      throw new IllegalArgumentException("fontName must not be blank");
    }
    fontColor = ExcelRgbColorSupport.normalizeRgbHex(fontColor, "fontColor");
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
