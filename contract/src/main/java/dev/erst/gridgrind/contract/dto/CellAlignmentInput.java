package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;

/** Protocol-facing alignment patch used by {@link CellStyleInput}. */
public record CellAlignmentInput(
    Boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment,
    Integer textRotation,
    Integer indentation) {
  private static final int MAX_TEXT_ROTATION = 180;
  private static final int MAX_INDENTATION = 250;

  public CellAlignmentInput {
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
