package dev.erst.gridgrind.excel;

/** Inclusive zero-based column band used by workbook-core structural column commands. */
public record ExcelColumnSpan(int firstColumnIndex, int lastColumnIndex) {
  /** Maximum column index addressable in one `.xlsx` worksheet. */
  public static final int MAX_COLUMN_INDEX = 16_383; // LIM-009

  public ExcelColumnSpan {
    if (firstColumnIndex < 0) {
      throw new IllegalArgumentException("firstColumnIndex must not be negative");
    }
    if (lastColumnIndex < 0) {
      throw new IllegalArgumentException("lastColumnIndex must not be negative");
    }
    if (lastColumnIndex < firstColumnIndex) {
      throw new IllegalArgumentException("lastColumnIndex must not be less than firstColumnIndex");
    }
    if (firstColumnIndex > MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "firstColumnIndex must not exceed " + MAX_COLUMN_INDEX + " (Excel column limit)");
    }
    if (lastColumnIndex > MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "lastColumnIndex must not exceed " + MAX_COLUMN_INDEX + " (Excel column limit)");
    }
  }

  /** Returns the number of columns covered by the inclusive span. */
  public int count() {
    return lastColumnIndex - firstColumnIndex + 1;
  }
}
