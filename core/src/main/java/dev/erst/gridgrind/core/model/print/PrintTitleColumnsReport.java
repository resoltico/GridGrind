package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Repeating print-title column state captured from one sheet. */
@Union(discriminator = "type")
public sealed interface PrintTitleColumnsReport
    permits PrintTitleColumnsReport.None, PrintTitleColumnsReport.Band {
  /** Sheet has no repeating print-title columns. */
  @Variant(id = "NONE", summary = "Sheet has no repeating print-title columns.")
  record None() implements PrintTitleColumnsReport {}

  /** Sheet repeats the provided inclusive zero-based column band on every printed page. */
  @Variant(
      id = "BAND",
      summary =
          "Sheet repeats the provided inclusive zero-based column band on every printed page.")
  record Band(int firstColumnIndex, int lastColumnIndex) implements PrintTitleColumnsReport {}
}
