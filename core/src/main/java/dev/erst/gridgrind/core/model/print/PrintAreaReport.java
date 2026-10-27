package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.Objects;

/** Print-area state captured from one sheet. */
@Union(discriminator = "type")
public sealed interface PrintAreaReport permits PrintAreaReport.None, PrintAreaReport.Range {
  /** Sheet has no explicit print area. */
  @Variant(id = "NONE", summary = "Sheet has no explicit print area.")
  record None() implements PrintAreaReport {}

  /** Sheet prints the provided rectangular A1-style range. */
  @Variant(id = "RANGE", summary = "Sheet prints the provided rectangular A1-style range.")
  record Range(RangeLabel range) implements PrintAreaReport {
    public Range {
      Objects.requireNonNull(range, "range must not be null");
    }
  }
}
