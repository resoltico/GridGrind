package dev.erst.gridgrind.excel;

/** Inclusive zero-based row band used by workbook-core structural row commands. */
public record ExcelRowSpan(int firstRowIndex, int lastRowIndex) {
  /** Maximum row index addressable in one `.xlsx` worksheet. */
  public static final int MAX_ROW_INDEX = 1_048_575; // LIM-008

  public ExcelRowSpan {
    if (firstRowIndex < 0) {
      throw new IllegalArgumentException("firstRowIndex must not be negative");
    }
    if (lastRowIndex < 0) {
      throw new IllegalArgumentException("lastRowIndex must not be negative");
    }
    if (lastRowIndex < firstRowIndex) {
      throw new IllegalArgumentException("lastRowIndex must not be less than firstRowIndex");
    }
    if (firstRowIndex > MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "firstRowIndex must not exceed " + MAX_ROW_INDEX + " (Excel row limit)");
    }
    if (lastRowIndex > MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "lastRowIndex must not exceed " + MAX_ROW_INDEX + " (Excel row limit)");
    }
  }

  /** Returns the number of rows covered by the inclusive span. */
  public int count() {
    return lastRowIndex - firstRowIndex + 1;
  }
}
