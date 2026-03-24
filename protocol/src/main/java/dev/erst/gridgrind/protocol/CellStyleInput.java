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
    HorizontalAlignmentInput horizontalAlignment,
    VerticalAlignmentInput verticalAlignment) {
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
        numberFormat,
        bold,
        italic,
        wrapText,
        horizontalAlignment == null ? null : horizontalAlignment.toEngine(),
        verticalAlignment == null ? null : verticalAlignment.toEngine());
  }

  /** Supported horizontal alignments for JSON input. */
  public enum HorizontalAlignmentInput {
    GENERAL,
    LEFT,
    CENTER,
    RIGHT,
    FILL,
    JUSTIFY,
    CENTER_SELECTION,
    DISTRIBUTED;

    ExcelHorizontalAlignment toEngine() {
      return ExcelHorizontalAlignment.valueOf(name());
    }
  }

  /** Supported vertical alignments for JSON input. */
  public enum VerticalAlignmentInput {
    TOP,
    CENTER,
    BOTTOM,
    JUSTIFY,
    DISTRIBUTED;

    ExcelVerticalAlignment toEngine() {
      return ExcelVerticalAlignment.valueOf(name());
    }
  }
}
