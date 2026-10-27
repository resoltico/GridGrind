package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import java.util.Objects;

/** One requested print-area state in protocol form. */
@Union(discriminator = "type", catalog = true, displayOrder = 66)
public sealed interface PrintAreaSpec
    permits dev.erst.gridgrind.core.model.print.PrintAreaSpec.None,
        dev.erst.gridgrind.core.model.print.PrintAreaSpec.Range {
  /** Sheet has no explicit print area. */
  @Variant(id = "NONE", summary = "Sheet has no explicit print area.")
  record None() implements PrintAreaSpec {}

  /** Sheet prints the provided rectangular A1-style range. */
  @Variant(id = "RANGE", summary = "Sheet prints the provided rectangular A1-style range.")
  record Range(RangeRef range) implements PrintAreaSpec {
    public Range {
      Objects.requireNonNull(range, "range must not be null");
    }
  }
}
