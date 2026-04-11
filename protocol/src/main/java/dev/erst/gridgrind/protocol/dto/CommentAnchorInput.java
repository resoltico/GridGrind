package dev.erst.gridgrind.protocol.dto;

/** Comment-anchor bounds used by rich comment authoring. */
public record CommentAnchorInput(int firstColumn, int firstRow, int lastColumn, int lastRow) {
  public CommentAnchorInput {
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
