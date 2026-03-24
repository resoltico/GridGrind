package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.util.CellReference;

/** Normalized rectangular cell range expressed with zero-based row and column bounds. */
final class ExcelRange {
  private final int firstRow;
  private final int lastRow;
  private final int firstColumn;
  private final int lastColumn;

  private ExcelRange(int firstRow, int lastRow, int firstColumn, int lastColumn) {
    this.firstRow = firstRow;
    this.lastRow = lastRow;
    this.firstColumn = firstColumn;
    this.lastColumn = lastColumn;
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

  int firstRow() {
    return firstRow;
  }

  int lastRow() {
    return lastRow;
  }

  int firstColumn() {
    return firstColumn;
  }

  int lastColumn() {
    return lastColumn;
  }

  int rowCount() {
    return lastRow - firstRow + 1;
  }

  int columnCount() {
    return lastColumn - firstColumn + 1;
  }

  private static CellReference parseCell(String address, String range) {
    if (address == null || address.isBlank()) {
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
