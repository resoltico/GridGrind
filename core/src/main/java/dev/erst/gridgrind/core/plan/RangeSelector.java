package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;

/** Selects one or more A1-style rectangular ranges on one sheet. */
@Union(discriminator = "type", catalog = true, displayOrder = 28)
public sealed interface RangeSelector extends Selector
    permits dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet,
        dev.erst.gridgrind.core.plan.RangeSelector.ByRange,
        dev.erst.gridgrind.core.plan.RangeSelector.ByRanges,
        dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow {

  /** Selects all matching ranged structures on one sheet. */
  @Variant(
      id = "RANGE_ALL_ON_SHEET",
      summary = "Select every matching range-backed structure on one sheet.")
  record AllOnSheet(SheetName sheetName) implements RangeSelector, ResolvedTarget {
    public AllOnSheet {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one exact rectangular range on one sheet. */
  @Variant(id = "RANGE_BY_RANGE", summary = "Select one exact rectangular range on one sheet.")
  record ByRange(SheetName sheetName, RangeRef range) implements RangeSelector, ResolvedTarget {
    public ByRange {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      range = SelectorValueValidation.requireOrderedRange(range, "range");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }

    /** Returns the exact number of cells in this validated rectangular range. */
    public long cellCount() {
      String[] endpoints = range.value().split(":", -1);
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
  @Variant(
      id = "RANGE_BY_RANGES",
      summary = "Select one or more exact rectangular ranges on one sheet.")
  record ByRanges(SheetName sheetName, List<RangeRef> ranges)
      implements RangeSelector, ResolvedTarget {
    public ByRanges {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      ranges = SelectorListValidation.copyDistinctRanges(ranges, "ranges");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one rectangular window anchored at a top-left cell. */
  @Variant(
      id = "RANGE_RECTANGULAR_WINDOW",
      summary = "Select one rectangular window anchored at one top-left cell.")
  record RectangularWindow(
      SheetName sheetName, CellAddress topLeftAddress, int rowCount, int columnCount)
      implements RangeSelector, ResolvedTarget {
    public RectangularWindow {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      java.util.Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      rowCount = SelectorValueValidation.requirePositive(rowCount, "rowCount");
      columnCount = SelectorValueValidation.requirePositive(columnCount, "columnCount");
      SelectorValueValidation.requireWindowSize(rowCount, columnCount);
      if (SelectorAddressSupport.rowIndex(topLeftAddress.value()) + rowCount - 1
              > Limit.MAX_ROW_INDEX
          || SelectorAddressSupport.columnIndex(topLeftAddress.value()) + columnCount - 1
              > Limit.MAX_COLUMN_INDEX) {
        throw new dev.erst.gridgrind.core.problem.InvalidRequestException(
            new dev.erst.gridgrind.core.problem.ActionableInvariantMessage(
                "window must end within Excel .xlsx bounds",
                "Choose a top-left address and dimensions that keep the entire window inside the worksheet.",
                java.util.Optional.empty()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            null);
      }
    }

    /** Returns the exact A1-style rectangular range implied by this window. */
    public RangeRef range() {
      int firstRow = SelectorAddressSupport.rowIndex(topLeftAddress.value());
      int firstColumn = SelectorAddressSupport.columnIndex(topLeftAddress.value());
      int lastRow = firstRow + rowCount - 1;
      int lastColumn = firstColumn + columnCount - 1;
      return new RangeRef(
          topLeftAddress + ":" + SelectorAddressSupport.absoluteA1Address(lastRow, lastColumn));
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
