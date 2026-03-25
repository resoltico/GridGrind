package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;

/** Protocol-facing style patch used for range and cell presentation changes. */
public record CellStyleInput(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    Boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment) {
  public CellStyleInput {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    if (numberFormat == null
        && bold == null
        && italic == null
        && wrapText == null
        && horizontalAlignment == null
        && verticalAlignment == null) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }

  /** Converts this transport model into the engine style model. */
  public ExcelCellStyle toExcelCellStyle() {
    return new ExcelCellStyle(
        numberFormat, bold, italic, wrapText, horizontalAlignment, verticalAlignment);
  }
}
