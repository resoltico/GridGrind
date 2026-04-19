package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Selects a row band or one row-band insertion point on one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RowBandSelector.Span.class, name = "SPAN"),
  @JsonSubTypes.Type(value = RowBandSelector.Insertion.class, name = "INSERTION")
})
public sealed interface RowBandSelector extends Selector
    permits RowBandSelector.Span, RowBandSelector.Insertion {

  /** Selects one inclusive zero-based row span on one sheet. */
  record Span(String sheetName, int firstRowIndex, int lastRowIndex) implements RowBandSelector {
    public Span {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      firstRowIndex = SelectorSupport.requireRowIndexWithinBounds(firstRowIndex, "firstRowIndex");
      lastRowIndex = SelectorSupport.requireRowIndexWithinBounds(lastRowIndex, "lastRowIndex");
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
  record Insertion(String sheetName, int beforeRowIndex, int rowCount) implements RowBandSelector {
    public Insertion {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      beforeRowIndex =
          SelectorSupport.requireRowIndexWithinBounds(beforeRowIndex, "beforeRowIndex");
      rowCount = SelectorSupport.requirePositive(rowCount, "rowCount");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
