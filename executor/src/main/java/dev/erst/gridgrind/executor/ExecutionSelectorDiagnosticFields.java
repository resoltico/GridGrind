package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;

/** Selector diagnostic field extraction for execution journal contexts. */
final class ExecutionSelectorDiagnosticFields {
  private ExecutionSelectorDiagnosticFields() {}

  static String sheetNameFor(Selector selector) {
    if (selector instanceof WorkbookSelector) {
      return null;
    }
    if (selector instanceof SheetSelector sheetSelector) {
      return sheetNameFor(sheetSelector);
    }
    if (selector instanceof CellSelector cellSelector) {
      return singleSheetName(cellSelector);
    }
    if (selector instanceof RangeSelector rangeSelector) {
      return singleSheetName(rangeSelector);
    }
    if (selector instanceof RowBandSelector.Span span) {
      return span.sheetName();
    }
    if (selector instanceof RowBandSelector.Insertion insertion) {
      return insertion.sheetName();
    }
    if (selector instanceof ColumnBandSelector.Span span) {
      return span.sheetName();
    }
    if (selector instanceof ColumnBandSelector.Insertion insertion) {
      return insertion.sheetName();
    }
    if (selector instanceof DrawingObjectSelector.AllOnSheet allOnSheet) {
      return allOnSheet.sheetName();
    }
    if (selector instanceof DrawingObjectSelector.ByName byName) {
      return byName.sheetName();
    }
    if (selector instanceof ChartSelector.AllOnSheet allOnSheet) {
      return allOnSheet.sheetName();
    }
    if (selector instanceof ChartSelector.ByName byName) {
      return byName.sheetName();
    }
    if (selector instanceof TableSelector tableSelector) {
      return sheetNameFor(tableSelector);
    }
    if (selector instanceof PivotTableSelector pivotTableSelector) {
      return sheetNameFor(pivotTableSelector);
    }
    if (selector instanceof NamedRangeSelector namedRangeSelector) {
      return singleSheetName(namedRangeSelector);
    }
    if (selector instanceof TableRowSelector tableRowSelector) {
      return sheetNameFor(tableRowSelector);
    }
    if (selector instanceof TableCellSelector.ByColumnName byColumnName) {
      return sheetNameFor(byColumnName.row());
    }
    return null;
  }

  static String addressFor(Selector selector) {
    if (selector instanceof CellSelector.ByAddress byAddress) {
      return byAddress.address();
    }
    if (selector instanceof CellSelector.ByQualifiedAddresses qualifiedAddresses) {
      return qualifiedAddresses.cells().size() == 1
          ? qualifiedAddresses.cells().getFirst().address()
          : null;
    }
    if (selector instanceof RangeSelector.RectangularWindow window) {
      return window.topLeftAddress();
    }
    return null;
  }

  static String rangeFor(Selector selector) {
    if (selector instanceof RangeSelector.ByRange byRange) {
      return byRange.range();
    }
    if (selector instanceof RangeSelector.ByRanges byRanges) {
      return byRanges.ranges().size() == 1 ? byRanges.ranges().getFirst() : null;
    }
    if (selector instanceof RangeSelector.RectangularWindow window) {
      return window.range();
    }
    return null;
  }

  static String namedRangeNameFor(Selector selector) {
    if (selector instanceof NamedRangeSelector namedRangeSelector) {
      return singleNamedRangeName(namedRangeSelector);
    }
    return null;
  }

  static String singleSheetName(CellSelector selector) {
    return switch (selector) {
      case CellSelector.AllUsedInSheet allUsedInSheet -> allUsedInSheet.sheetName();
      case CellSelector.ByAddress byAddress -> byAddress.sheetName();
      case CellSelector.ByAddresses byAddresses -> byAddresses.sheetName();
      case CellSelector.ByQualifiedAddresses qualifiedAddresses ->
          qualifiedAddresses.cells().stream()
                      .map(CellSelector.QualifiedAddress::sheetName)
                      .distinct()
                      .count()
                  == 1
              ? qualifiedAddresses.cells().getFirst().sheetName()
              : null;
    };
  }

  static String singleSheetName(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet -> allOnSheet.sheetName();
      case RangeSelector.ByRange byRange -> byRange.sheetName();
      case RangeSelector.ByRanges byRanges -> byRanges.sheetName();
      case RangeSelector.RectangularWindow window -> window.sheetName();
    };
  }

  static String singleSheetName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> null;
      case NamedRangeSelector.ByName _ -> null;
      case NamedRangeSelector.ByNames _ -> null;
      case NamedRangeSelector.WorkbookScope _ -> null;
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.sheetName();
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1 ? singleSheetName(anyOf.selectors().getFirst()) : null;
    };
  }

  static String singleSheetName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName _ -> null;
      case NamedRangeSelector.WorkbookScope _ -> null;
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.sheetName();
    };
  }

  static String singleNamedRangeName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> null;
      case NamedRangeSelector.ByName byName -> byName.name();
      case NamedRangeSelector.ByNames byNames ->
          byNames.names().size() == 1 ? byNames.names().getFirst() : null;
      case NamedRangeSelector.WorkbookScope workbookScope -> workbookScope.name();
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.name();
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1 ? singleNamedRangeName(anyOf.selectors().getFirst()) : null;
    };
  }

  static String singleNamedRangeName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> byName.name();
      case NamedRangeSelector.WorkbookScope workbookScope -> workbookScope.name();
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.name();
    };
  }

  private static String sheetNameFor(SheetSelector selector) {
    return switch (selector) {
      case SheetSelector.All _ -> null;
      case SheetSelector.ByName byName -> byName.name();
      case SheetSelector.ByNames byNames ->
          byNames.names().size() == 1 ? byNames.names().getFirst() : null;
    };
  }

  private static String sheetNameFor(TableSelector selector) {
    return switch (selector) {
      case TableSelector.All _ -> null;
      case TableSelector.ByName _ -> null;
      case TableSelector.ByNames _ -> null;
      case TableSelector.ByNameOnSheet byNameOnSheet -> byNameOnSheet.sheetName();
    };
  }

  private static String sheetNameFor(PivotTableSelector selector) {
    return switch (selector) {
      case PivotTableSelector.All _ -> null;
      case PivotTableSelector.ByName _ -> null;
      case PivotTableSelector.ByNames _ -> null;
      case PivotTableSelector.ByNameOnSheet byNameOnSheet -> byNameOnSheet.sheetName();
    };
  }

  private static String sheetNameFor(TableRowSelector selector) {
    return switch (selector) {
      case TableRowSelector.AllRows allRows -> sheetNameFor(allRows.table());
      case TableRowSelector.ByIndex byIndex -> sheetNameFor(byIndex.table());
      case TableRowSelector.ByKeyCell byKeyCell -> sheetNameFor(byKeyCell.table());
    };
  }
}
