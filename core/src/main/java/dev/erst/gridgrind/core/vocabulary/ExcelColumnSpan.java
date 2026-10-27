package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.limits.Limit;

/** Inclusive zero-based column band used by workbook-core structural column commands. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_COLUMNS_PER_WORKSHEET
})
public record ExcelColumnSpan(int firstColumnIndex, int lastColumnIndex) {
  /** Maximum column index addressable in one `.xlsx` worksheet. */
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
    if (firstColumnIndex > Limit.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed(
              "firstColumnIndex", firstColumnIndex, Limit.MAX_COLUMN_INDEX));
    }
    if (lastColumnIndex > Limit.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed(
              "lastColumnIndex", lastColumnIndex, Limit.MAX_COLUMN_INDEX));
    }
  }

  /** Returns the number of columns covered by the inclusive span. */
  public int count() {
    return lastColumnIndex - firstColumnIndex + 1;
  }
}
