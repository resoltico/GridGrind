package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Mutable-workbook advanced page-setup payload. */
public record ExcelPrintSetup(
    ExcelPrintMargins margins,
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
  /** Returns the default advanced page-setup payload for a sheet with no explicit setup. */
  public static ExcelPrintSetup defaults() {
    return new ExcelPrintSetup(
        new ExcelPrintMargins(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
        false,
        false,
        false,
        1,
        false,
        false,
        1,
        false,
        1,
        List.of(),
        List.of());
  }

  public ExcelPrintSetup {
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
