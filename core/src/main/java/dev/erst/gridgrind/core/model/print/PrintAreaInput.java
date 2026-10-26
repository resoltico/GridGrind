package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import java.util.Objects;

/** One requested print-area state in protocol form. */
@Union(discriminator = "type", catalog = true, displayOrder = 66)
public sealed interface PrintAreaInput
    permits dev.erst.gridgrind.core.model.print.PrintAreaInput.None,
        dev.erst.gridgrind.core.model.print.PrintAreaInput.Range {
  /** Sheet has no explicit print area. */
  @Variant(id = "NONE", summary = "Sheet has no explicit print area.")
  record None() implements PrintAreaInput {}

  /** Sheet prints the provided rectangular A1-style range. */
  @Variant(id = "RANGE", summary = "Sheet prints the provided rectangular A1-style range.")
  record Range(RangeRef range) implements PrintAreaInput {
    public Range {
      Objects.requireNonNull(range, "range must not be null");
    }
  }
}
