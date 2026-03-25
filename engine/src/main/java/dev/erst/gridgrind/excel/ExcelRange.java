package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.util.CellReference;

/** Normalized rectangular cell range expressed with zero-based row and column bounds. */
record ExcelRange(int firstRow, int lastRow, int firstColumn, int lastColumn) {
  ExcelRange {
    if (firstRow < 0) {
      throw new IllegalArgumentException("firstRow must not be negative");
    }
    if (lastRow < firstRow) {
      throw new IllegalArgumentException("lastRow must not be less than firstRow");
    }
    if (firstColumn < 0) {
      throw new IllegalArgumentException("firstColumn must not be negative");
    }
    if (lastColumn < firstColumn) {
      throw new IllegalArgumentException("lastColumn must not be less than firstColumn");
    }
  }

  static ExcelRange parse(String range) {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }

    String[] parts = range.split(":", -1);
    if (parts.length == 1) {
      CellReference cell = parseCell(range, range);
      return new ExcelRange(cell.getRow(), cell.getRow(), cell.getCol(), cell.getCol());
    }
    if (parts.length == 2) {
      CellReference start = parseCell(parts[0], range);
      CellReference end = parseCell(parts[1], range);
      return new ExcelRange(
          Math.min(start.getRow(), end.getRow()),
          Math.max(start.getRow(), end.getRow()),
          Math.min(start.getCol(), end.getCol()),
          Math.max(start.getCol(), end.getCol()));
    }
    throw new InvalidRangeAddressException(
        range, new IllegalArgumentException("range must contain at most one ':'"));
  }

  /** Returns the number of rows in this range. */
  int rowCount() {
    return lastRow - firstRow + 1;
  }

  /** Returns the number of columns in this range. */
  int columnCount() {
    return lastColumn - firstColumn + 1;
  }

  private static CellReference parseCell(String address, String range) {
    if (address.isBlank()) {
      throw new InvalidRangeAddressException(
          range, new IllegalArgumentException("range endpoint must not be blank"));
    }
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidRangeAddressException(range, exception);
    }
  }
}
