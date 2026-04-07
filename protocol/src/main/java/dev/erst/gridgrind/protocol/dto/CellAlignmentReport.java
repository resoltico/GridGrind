package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import java.util.Objects;

/** Effective alignment facts reported with every analyzed cell. */
public record CellAlignmentReport(
    boolean wrapText,
    ExcelHorizontalAlignment horizontalAlignment,
    ExcelVerticalAlignment verticalAlignment,
    int textRotation,
    int indentation) {
  public CellAlignmentReport {
    Objects.requireNonNull(horizontalAlignment, "horizontalAlignment must not be null");
    Objects.requireNonNull(verticalAlignment, "verticalAlignment must not be null");
  }
}
