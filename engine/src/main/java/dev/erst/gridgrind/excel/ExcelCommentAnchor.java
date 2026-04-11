package dev.erst.gridgrind.excel;

/** Mutable-workbook comment-anchor bounds. */
public record ExcelCommentAnchor(int firstColumn, int firstRow, int lastColumn, int lastRow) {
  public ExcelCommentAnchor {
    requireNonNegative(firstColumn, "firstColumn");
    requireNonNegative(firstRow, "firstRow");
    requireNonNegative(lastColumn, "lastColumn");
    requireNonNegative(lastRow, "lastRow");
    if (lastColumn < firstColumn) {
      throw new IllegalArgumentException("lastColumn must be greater than or equal to firstColumn");
    }
    if (lastRow < firstRow) {
      throw new IllegalArgumentException("lastRow must be greater than or equal to firstRow");
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }
}
