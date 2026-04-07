package dev.erst.gridgrind.excel;

/** Alignment patch applied through {@link ExcelCellStyle}. */
public record ExcelCellAlignment(
    Boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment,
    Integer textRotation,
    Integer indentation) {
  private static final int MAX_TEXT_ROTATION = 180;
  private static final int MAX_INDENTATION = 250;

  public ExcelCellAlignment {
    if (wrapText == null
        && horizontalAlignment == null
        && verticalAlignment == null
        && textRotation == null
        && indentation == null) {
      throw new IllegalArgumentException("alignment must set at least one attribute");
    }
    if (textRotation != null && (textRotation < 0 || textRotation > MAX_TEXT_ROTATION)) {
      throw new IllegalArgumentException(
          "textRotation must be between 0 and " + MAX_TEXT_ROTATION + " (inclusive)");
    }
    if (indentation != null && (indentation < 0 || indentation > MAX_INDENTATION)) {
      throw new IllegalArgumentException(
          "indentation must be between 0 and " + MAX_INDENTATION + " (inclusive)");
    }
  }
}
