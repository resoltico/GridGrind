package dev.erst.gridgrind.excel;

/** Agent-facing style patch that can be applied to one cell or a rectangular range. */
public record ExcelCellStyle(
    String numberFormat,
    ExcelCellAlignment alignment,
    ExcelCellFont font,
    ExcelCellFill fill,
    ExcelBorder border,
    ExcelCellProtection protection) {
  public ExcelCellStyle {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    if (numberFormat == null
        && alignment == null
        && font == null
        && fill == null
        && border == null
        && protection == null) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }

  /** Creates a style patch that only changes the number format. */
  public static ExcelCellStyle numberFormat(String numberFormat) {
    return new ExcelCellStyle(numberFormat, null, null, null, null, null);
  }

  /** Creates a style patch that only changes font emphasis. */
  public static ExcelCellStyle emphasis(Boolean bold, Boolean italic) {
    return new ExcelCellStyle(
        null,
        null,
        new ExcelCellFont(bold, italic, null, null, (ExcelColor) null, null, null),
        null,
        null,
        null);
  }

  /** Creates a style patch that only changes horizontal and vertical alignment. */
  public static ExcelCellStyle alignment(
      ExcelHorizontalAlignment horizontalAlignment, ExcelVerticalAlignment verticalAlignment) {
    if (horizontalAlignment == null && verticalAlignment == null) {
      throw new IllegalArgumentException("alignment must set at least one attribute");
    }
    return new ExcelCellStyle(
        null,
        new ExcelCellAlignment(null, horizontalAlignment, verticalAlignment, null, null),
        null,
        null,
        null,
        null);
  }
}
