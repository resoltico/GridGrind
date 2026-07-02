package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable rectangular window bounds shared by workbook read commands. */
public record ExcelReadWindow(String topLeftAddress, int rowCount, int columnCount) {
  public ExcelReadWindow {
    Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
    if (topLeftAddress.isBlank()) {
      throw new IllegalArgumentException("topLeftAddress must not be blank");
    }
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be positive");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be positive");
    }
    long totalCells = (long) rowCount * columnCount;
    if (totalCells > WorkbookReadCommand.MAX_READ_CELLS) {
      throw new IllegalArgumentException(
          "rowCount * columnCount must be <= " + WorkbookReadCommand.MAX_READ_CELLS);
    }
  }
}
