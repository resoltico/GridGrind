package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.limits.Limit;

/** Inclusive zero-based row band used by workbook-core structural row commands. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_ROWS_PER_WORKSHEET
})
public record ExcelRowSpan(int firstRowIndex, int lastRowIndex) {
  /** Maximum row index addressable in one `.xlsx` worksheet. */
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
    if (firstRowIndex > Limit.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed("firstRowIndex", firstRowIndex, Limit.MAX_ROW_INDEX));
    }
    if (lastRowIndex > Limit.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotExceed("lastRowIndex", lastRowIndex, Limit.MAX_ROW_INDEX));
    }
  }

  /** Returns the number of rows covered by the inclusive span. */
  public int count() {
    return lastRowIndex - firstRowIndex + 1;
  }
}
