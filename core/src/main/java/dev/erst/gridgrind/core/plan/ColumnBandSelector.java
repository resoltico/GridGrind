package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects a column band or one column-band insertion point on one sheet. */
@Union(discriminator = "type", catalog = true, displayOrder = 30)
public sealed interface ColumnBandSelector extends Selector
    permits dev.erst.gridgrind.core.plan.ColumnBandSelector.Span,
        dev.erst.gridgrind.core.plan.ColumnBandSelector.Insertion {

  /** Selects one inclusive zero-based column span on one sheet. */
  @Variant(
      id = "COLUMN_BAND_SPAN",
      summary = "Select one inclusive zero-based column span on one sheet.")
  record Span(SheetName sheetName, int firstColumnIndex, int lastColumnIndex)
      implements ColumnBandSelector, ResolvedTarget {
    public Span {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      firstColumnIndex =
          SelectorValueValidation.requireColumnIndexWithinBounds(
              firstColumnIndex, "firstColumnIndex");
      lastColumnIndex =
          SelectorValueValidation.requireColumnIndexWithinBounds(
              lastColumnIndex, "lastColumnIndex");
      if (lastColumnIndex < firstColumnIndex) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not be less than firstColumnIndex");
      }
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one column insertion point and band count on one sheet. */
  @Variant(
      id = "COLUMN_BAND_INSERTION",
      summary = "Select one column insertion point plus column count on one sheet.")
  record Insertion(SheetName sheetName, int beforeColumnIndex, int columnCount)
      implements ColumnBandSelector, ResolvedTarget {
    public Insertion {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      beforeColumnIndex =
          SelectorValueValidation.requireColumnIndexWithinBounds(
              beforeColumnIndex, "beforeColumnIndex");
      columnCount = SelectorValueValidation.requirePositive(columnCount, "columnCount");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
