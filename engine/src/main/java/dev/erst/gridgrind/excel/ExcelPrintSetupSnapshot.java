package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Immutable advanced page-setup facts loaded from one worksheet. */
public record ExcelPrintSetupSnapshot(
    ExcelPrintMarginsSnapshot margins,
    boolean printGridlines,
    boolean horizontallyCentered,
    boolean verticallyCentered,
    int paperSize,
    boolean draft,
    boolean blackAndWhite,
    int copies,
    boolean useFirstPageNumber,
    int firstPageNumber,
    List<Integer> rowBreaks,
    List<Integer> columnBreaks) {
  public ExcelPrintSetupSnapshot {
    Objects.requireNonNull(margins, "margins must not be null");
    if (paperSize < 0) {
      throw new IllegalArgumentException("paperSize must not be negative");
    }
    if (copies < 0) {
      throw new IllegalArgumentException("copies must not be negative");
    }
    if (firstPageNumber < 0) {
      throw new IllegalArgumentException("firstPageNumber must not be negative");
    }
    rowBreaks = copyIndexes(rowBreaks, "rowBreaks");
    columnBreaks = copyIndexes(columnBreaks, "columnBreaks");
  }

  private static List<Integer> copyIndexes(List<Integer> indexes, String fieldName) {
    List<Integer> copy =
        List.copyOf(Objects.requireNonNull(indexes, fieldName + " must not be null"));
    for (Integer index : copy) {
      if (index < 0) {
        throw new IllegalArgumentException(fieldName + " must not contain negative indexes");
      }
    }
    return copy;
  }
}
