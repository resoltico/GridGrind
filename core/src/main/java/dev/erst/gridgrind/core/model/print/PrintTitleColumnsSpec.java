package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelIndexDisplay;

/** One requested repeating print-title column band in protocol form. */
@Union(discriminator = "type", catalog = true, displayOrder = 69)
public sealed interface PrintTitleColumnsSpec
    permits dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec.None,
        dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec.Band {

  /** Sheet has no repeating print-title columns. */
  @Variant(id = "NONE", summary = "Sheet has no repeating print-title columns.")
  record None() implements PrintTitleColumnsSpec {}

  /** Sheet repeats the provided inclusive zero-based column band on every printed page. */
  @Variant(
      id = "BAND",
      summary =
          "Sheet repeats the provided inclusive zero-based column band on every printed page.")
  record Band(int firstColumnIndex, int lastColumnIndex) implements PrintTitleColumnsSpec {
    public Band {
      requireBand(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");
      if (lastColumnIndex > Limit.MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotExceed(
                "lastColumnIndex", lastColumnIndex, Limit.MAX_COLUMN_INDEX));
      }
    }
  }

  private static void requireBand(
      int firstIndex, int lastIndex, String firstFieldName, String lastFieldName) {
    if (firstIndex < 0) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeNegative(firstFieldName, firstIndex));
    }
    if (lastIndex < 0) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeNegative(lastFieldName, lastIndex));
    }
    if (lastIndex < firstIndex) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeLessThan(
              lastFieldName, lastIndex, firstFieldName, firstIndex));
    }
  }
}
