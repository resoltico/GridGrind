package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more A1-style rectangular ranges on one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RangeSelector.AllOnSheet.class, name = "RANGE_ALL_ON_SHEET"),
  @JsonSubTypes.Type(value = RangeSelector.ByRange.class, name = "RANGE_BY_RANGE"),
  @JsonSubTypes.Type(value = RangeSelector.ByRanges.class, name = "RANGE_BY_RANGES"),
  @JsonSubTypes.Type(
      value = RangeSelector.RectangularWindow.class,
      name = "RANGE_RECTANGULAR_WINDOW")
})
public sealed interface RangeSelector extends Selector
    permits RangeSelector.AllOnSheet,
        RangeSelector.ByRange,
        RangeSelector.ByRanges,
        RangeSelector.RectangularWindow {

  /** Selects all matching ranged structures on one sheet. */
  record AllOnSheet(String sheetName) implements RangeSelector {
    public AllOnSheet {
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one exact rectangular range on one sheet. */
  record ByRange(String sheetName, String range) implements RangeSelector {
    public ByRange {
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
      range = SelectorValueValidation.requireRange(range, "range");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }

    /** Returns the exact number of cells in this validated rectangular range. */
    public long cellCount() {
      String[] endpoints = range.split(":", -1);
      String first = endpoints[0];
      String last = endpoints.length == 1 ? first : endpoints[1];
      long rowCount =
          (long) SelectorAddressSupport.rowIndex(last)
              - SelectorAddressSupport.rowIndex(first)
              + 1L;
      long columnCount =
          (long) SelectorAddressSupport.columnIndex(last)
              - SelectorAddressSupport.columnIndex(first)
              + 1L;
      return Math.multiplyExact(rowCount, columnCount);
    }
  }

  /** Selects one or more exact rectangular ranges on one sheet. */
  record ByRanges(String sheetName, List<String> ranges) implements RangeSelector {
    public ByRanges {
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
      ranges = SelectorListValidation.copyDistinctRanges(ranges, "ranges");
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
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
      topLeftAddress = SelectorValueValidation.requireAddress(topLeftAddress, "topLeftAddress");
      rowCount = SelectorValueValidation.requirePositive(rowCount, "rowCount");
      columnCount = SelectorValueValidation.requirePositive(columnCount, "columnCount");
      SelectorValueValidation.requireWindowSize(rowCount, columnCount);
    }

    /** Returns the exact A1-style rectangular range implied by this window. */
    public String range() {
      int firstRow = SelectorAddressSupport.rowIndex(topLeftAddress);
      int firstColumn = SelectorAddressSupport.columnIndex(topLeftAddress);
      int lastRow = firstRow + rowCount - 1;
      int lastColumn = firstColumn + columnCount - 1;
      return topLeftAddress + ":" + SelectorAddressSupport.absoluteA1Address(lastRow, lastColumn);
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
