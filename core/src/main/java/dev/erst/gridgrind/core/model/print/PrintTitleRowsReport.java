package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Repeating print-title row state captured from one sheet. */
@Union(discriminator = "type")
public sealed interface PrintTitleRowsReport
    permits PrintTitleRowsReport.None, PrintTitleRowsReport.Band {
  /** Sheet has no repeating print-title rows. */
  @Variant(id = "NONE", summary = "Sheet has no repeating print-title rows.")
  record None() implements PrintTitleRowsReport {}

  /** Sheet repeats the provided inclusive zero-based row band on every printed page. */
  @Variant(
      id = "BAND",
      summary = "Sheet repeats the provided inclusive zero-based row band on every printed page.")
  record Band(int firstRowIndex, int lastRowIndex) implements PrintTitleRowsReport {}
}
