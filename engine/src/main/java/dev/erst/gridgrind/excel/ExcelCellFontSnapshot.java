package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the font currently applied to a cell. */
public record ExcelCellFontSnapshot(
    boolean bold,
    boolean italic,
    String fontName,
    ExcelFontHeight fontHeight,
    ExcelColorSnapshot fontColor,
    boolean underline,
    boolean strikeout) {
  public ExcelCellFontSnapshot {
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
  }
}
