package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects a row band or one row-band insertion point on one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RowBandSelector.Span.class, name = "ROW_BAND_SPAN"),
  @JsonSubTypes.Type(value = RowBandSelector.Insertion.class, name = "ROW_BAND_INSERTION")
})
public sealed interface RowBandSelector extends Selector
    permits RowBandSelector.Span, RowBandSelector.Insertion {

  /** Selects one inclusive zero-based row span on one sheet. */
  record Span(SheetName sheetName, int firstRowIndex, int lastRowIndex) implements RowBandSelector {
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
  record Insertion(SheetName sheetName, int beforeRowIndex, int rowCount)
      implements RowBandSelector {
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
