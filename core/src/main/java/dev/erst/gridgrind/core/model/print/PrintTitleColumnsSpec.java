package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** One requested repeating print-title column band in protocol form. */
@Union(discriminator = "type", catalog = true, displayOrder = 69)
public sealed interface PrintTitleColumnsSpec
    permits dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec.None,
        dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec.Band {
  int MAX_COLUMN_INDEX = 16_383;

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
      if (lastColumnIndex > MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not exceed " + MAX_COLUMN_INDEX + ": " + lastColumnIndex);
      }
    }
  }

  private static void requireBand(
      int firstIndex, int lastIndex, String firstFieldName, String lastFieldName) {
    if (firstIndex < 0) {
      throw new IllegalArgumentException(firstFieldName + " must not be negative");
    }
    if (lastIndex < 0) {
      throw new IllegalArgumentException(lastFieldName + " must not be negative");
    }
    if (lastIndex < firstIndex) {
      throw new IllegalArgumentException(
          lastFieldName + " must not be less than " + firstFieldName);
    }
  }
}
