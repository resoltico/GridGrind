package dev.erst.gridgrind.excel.foundation;

/** Inclusive zero-based column band used by workbook-core structural column commands. */
public record ExcelColumnSpan(int firstColumnIndex, int lastColumnIndex) {
  /** Maximum column index addressable in one `.xlsx` worksheet. */
  public static final int MAX_COLUMN_INDEX = 16_383; // LIM-009

  public ExcelColumnSpan {
    if (firstColumnIndex < 0) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeNegative("firstColumnIndex", firstColumnIndex));
    }
    if (lastColumnIndex < 0) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeNegative("lastColumnIndex", lastColumnIndex));
    }
    if (lastColumnIndex < firstColumnIndex) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeLessThan(
              "lastColumnIndex", lastColumnIndex, "firstColumnIndex", firstColumnIndex));
    }
    if (firstColumnIndex > MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed("firstColumnIndex", firstColumnIndex, MAX_COLUMN_INDEX));
    }
    if (lastColumnIndex > MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed("lastColumnIndex", lastColumnIndex, MAX_COLUMN_INDEX));
    }
  }

  /** Returns the number of columns covered by the inclusive span. */
  public int count() {
    return lastColumnIndex - firstColumnIndex + 1;
  }
}
