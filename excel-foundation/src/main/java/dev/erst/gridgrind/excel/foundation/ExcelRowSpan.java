package dev.erst.gridgrind.excel.foundation;

/** Inclusive zero-based row band used by workbook-core structural row commands. */
public record ExcelRowSpan(int firstRowIndex, int lastRowIndex) {
  /** Maximum row index addressable in one `.xlsx` worksheet. */
  public static final int MAX_ROW_INDEX = 1_048_575; // LIM-008

  public ExcelRowSpan {
    if (firstRowIndex < 0) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeNegative("firstRowIndex", firstRowIndex));
    }
    if (lastRowIndex < 0) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeNegative("lastRowIndex", lastRowIndex));
    }
    if (lastRowIndex < firstRowIndex) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeLessThan(
              "lastRowIndex", lastRowIndex, "firstRowIndex", firstRowIndex));
    }
    if (firstRowIndex > MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed("firstRowIndex", firstRowIndex, MAX_ROW_INDEX));
    }
    if (lastRowIndex > MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed("lastRowIndex", lastRowIndex, MAX_ROW_INDEX));
    }
  }

  /** Returns the number of rows covered by the inclusive span. */
  public int count() {
    return lastRowIndex - firstRowIndex + 1;
  }
}
