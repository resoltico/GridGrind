package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the alignment currently applied to a cell. */
public record ExcelCellAlignmentSnapshot(
    boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment,
    int textRotation,
    int indentation) {
  public ExcelCellAlignmentSnapshot {
    Objects.requireNonNull(horizontalAlignment, "horizontalAlignment must not be null");
    Objects.requireNonNull(verticalAlignment, "verticalAlignment must not be null");
  }
}
