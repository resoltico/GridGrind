package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Selects a column band or one column-band insertion point on one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ColumnBandSelector.Span.class, name = "COLUMN_BAND_SPAN"),
  @JsonSubTypes.Type(value = ColumnBandSelector.Insertion.class, name = "COLUMN_BAND_INSERTION")
})
public sealed interface ColumnBandSelector extends Selector
    permits ColumnBandSelector.Span, ColumnBandSelector.Insertion {

  /** Selects one inclusive zero-based column span on one sheet. */
  record Span(String sheetName, int firstColumnIndex, int lastColumnIndex)
      implements ColumnBandSelector {
    public Span {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      firstColumnIndex =
          SelectorSupport.requireColumnIndexWithinBounds(firstColumnIndex, "firstColumnIndex");
      lastColumnIndex =
          SelectorSupport.requireColumnIndexWithinBounds(lastColumnIndex, "lastColumnIndex");
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
  record Insertion(String sheetName, int beforeColumnIndex, int columnCount)
      implements ColumnBandSelector {
    public Insertion {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      beforeColumnIndex =
          SelectorSupport.requireColumnIndexWithinBounds(beforeColumnIndex, "beforeColumnIndex");
      columnCount = SelectorSupport.requirePositive(columnCount, "columnCount");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
