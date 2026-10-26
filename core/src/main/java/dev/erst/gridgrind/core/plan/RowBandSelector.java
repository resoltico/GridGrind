package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects a row band or one row-band insertion point on one sheet. */
@Union(discriminator = "type", catalog = true, displayOrder = 29)
public sealed interface RowBandSelector extends Selector
    permits dev.erst.gridgrind.core.plan.RowBandSelector.Span,
        dev.erst.gridgrind.core.plan.RowBandSelector.Insertion {

  /** Selects one inclusive zero-based row span on one sheet. */
  @Variant(id = "ROW_BAND_SPAN", summary = "Select one inclusive zero-based row span on one sheet.")
  record Span(SheetName sheetName, int firstRowIndex, int lastRowIndex)
      implements RowBandSelector, ResolvedTarget {
    public Span {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      firstRowIndex =
          SelectorValueValidation.requireRowIndexWithinBounds(firstRowIndex, "firstRowIndex");
      lastRowIndex =
          SelectorValueValidation.requireRowIndexWithinBounds(lastRowIndex, "lastRowIndex");
      if (lastRowIndex < firstRowIndex) {
        throw new IllegalArgumentException("lastRowIndex must not be less than firstRowIndex");
      }
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one row insertion point and band count on one sheet. */
  @Variant(
      id = "ROW_BAND_INSERTION",
      summary = "Select one row insertion point plus row count on one sheet.")
  record Insertion(SheetName sheetName, int beforeRowIndex, int rowCount)
      implements RowBandSelector, ResolvedTarget {
    public Insertion {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      beforeRowIndex =
          SelectorValueValidation.requireRowIndexWithinBounds(beforeRowIndex, "beforeRowIndex");
      rowCount = SelectorValueValidation.requirePositive(rowCount, "rowCount");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
