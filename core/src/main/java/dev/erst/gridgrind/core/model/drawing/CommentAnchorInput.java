package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Comment-anchor bounds used by rich comment authoring. */
@CatalogType(
    id = "CommentAnchorInput",
    summary = "Explicit comment-anchor bounds measured in zero-based column and row indexes.",
    displayOrder = 41)
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
