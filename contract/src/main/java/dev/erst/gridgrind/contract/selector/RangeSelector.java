package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more A1-style rectangular ranges on one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RangeSelector.AllOnSheet.class, name = "ALL_ON_SHEET"),
  @JsonSubTypes.Type(value = RangeSelector.ByRange.class, name = "BY_RANGE"),
  @JsonSubTypes.Type(value = RangeSelector.ByRanges.class, name = "BY_RANGES"),
  @JsonSubTypes.Type(value = RangeSelector.RectangularWindow.class, name = "RECTANGULAR_WINDOW")
})
public sealed interface RangeSelector extends Selector
    permits RangeSelector.AllOnSheet,
        RangeSelector.ByRange,
        RangeSelector.ByRanges,
        RangeSelector.RectangularWindow {

  /** Selects all matching ranged structures on one sheet. */
  record AllOnSheet(String sheetName) implements RangeSelector {
    public AllOnSheet {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one exact rectangular range on one sheet. */
  record ByRange(String sheetName, String range) implements RangeSelector {
    public ByRange {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      range = SelectorSupport.requireRange(range, "range");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more exact rectangular ranges on one sheet. */
  record ByRanges(String sheetName, List<String> ranges) implements RangeSelector {
    public ByRanges {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      ranges = SelectorSupport.copyDistinctRanges(ranges, "ranges");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one rectangular window anchored at a top-left cell. */
  record RectangularWindow(String sheetName, String topLeftAddress, int rowCount, int columnCount)
      implements RangeSelector {
    public RectangularWindow {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      topLeftAddress = SelectorSupport.requireAddress(topLeftAddress, "topLeftAddress");
      rowCount = SelectorSupport.requirePositive(rowCount, "rowCount");
      columnCount = SelectorSupport.requirePositive(columnCount, "columnCount");
      SelectorSupport.requireWindowSize(rowCount, columnCount);
    }

    /** Returns the exact A1-style rectangular range implied by this window. */
    public String range() {
      int firstRow = SelectorSupport.rowIndex(topLeftAddress);
      int firstColumn = SelectorSupport.columnIndex(topLeftAddress);
      int lastRow = firstRow + rowCount - 1;
      int lastColumn = firstColumn + columnCount - 1;
      return topLeftAddress + ":" + SelectorSupport.absoluteA1Address(lastRow, lastColumn);
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
