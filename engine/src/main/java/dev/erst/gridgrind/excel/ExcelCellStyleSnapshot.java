package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the style currently applied to a cell. */
public record ExcelCellStyleSnapshot(
    String numberFormat,
    boolean bold,
    boolean italic,
    boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment,
    String fontName,
    ExcelFontHeight fontHeight,
    String fontColor,
    boolean underline,
    boolean strikeout,
    String fillColor,
    ExcelBorderStyle topBorderStyle,
    ExcelBorderStyle rightBorderStyle,
    ExcelBorderStyle bottomBorderStyle,
    ExcelBorderStyle leftBorderStyle) {
  public ExcelCellStyleSnapshot {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(horizontalAlignment, "horizontalAlignment must not be null");
    Objects.requireNonNull(verticalAlignment, "verticalAlignment must not be null");
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    Objects.requireNonNull(topBorderStyle, "topBorderStyle must not be null");
    Objects.requireNonNull(rightBorderStyle, "rightBorderStyle must not be null");
    Objects.requireNonNull(bottomBorderStyle, "bottomBorderStyle must not be null");
    Objects.requireNonNull(leftBorderStyle, "leftBorderStyle must not be null");
  }
}
