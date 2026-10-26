package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** One requested repeating print-title row band in protocol form. */
@Union(discriminator = "type", catalog = true, displayOrder = 68)
public sealed interface PrintTitleRowsSpec
    permits dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec.None,
        dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec.Band {
  int MAX_ROW_INDEX = 1_048_575;

  /** Sheet has no repeating print-title rows. */
  @Variant(id = "NONE", summary = "Sheet has no repeating print-title rows.")
  record None() implements PrintTitleRowsSpec {}

  /** Sheet repeats the provided inclusive zero-based row band on every printed page. */
  @Variant(
      id = "BAND",
      summary = "Sheet repeats the provided inclusive zero-based row band on every printed page.")
  record Band(int firstRowIndex, int lastRowIndex) implements PrintTitleRowsSpec {
    public Band {
      requireBand(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
      if (lastRowIndex > MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            "lastRowIndex must not exceed " + MAX_ROW_INDEX + ": " + lastRowIndex);
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
