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
import java.util.Optional;

/** Selector diagnostic field extraction for execution journal contexts. */
final class ExecutionSelectorDiagnosticFields {
  private ExecutionSelectorDiagnosticFields() {}

  static Optional<String> sheetNameFor(Selector selector) {
    if (selector instanceof WorkbookSelector) {
      return Optional.empty();
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
      return Optional.of(span.sheetName());
    }
    if (selector instanceof RowBandSelector.Insertion insertion) {
      return Optional.of(insertion.sheetName());
    }
    if (selector instanceof ColumnBandSelector.Span span) {
      return Optional.of(span.sheetName());
    }
    if (selector instanceof ColumnBandSelector.Insertion insertion) {
      return Optional.of(insertion.sheetName());
    }
    if (selector instanceof DrawingObjectSelector.AllOnSheet allOnSheet) {
      return Optional.of(allOnSheet.sheetName());
    }
    if (selector instanceof DrawingObjectSelector.ByName byName) {
      return Optional.of(byName.sheetName());
    }
    if (selector instanceof ChartSelector.AllOnSheet allOnSheet) {
      return Optional.of(allOnSheet.sheetName());
    }
    if (selector instanceof ChartSelector.ByName byName) {
      return Optional.of(byName.sheetName());
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
    return Optional.empty();
  }

  static Optional<String> addressFor(Selector selector) {
    if (selector instanceof CellSelector.ByAddress byAddress) {
      return Optional.of(byAddress.address());
    }
    if (selector instanceof CellSelector.ByQualifiedAddresses qualifiedAddresses) {
      return qualifiedAddresses.cells().size() == 1
          ? Optional.of(qualifiedAddresses.cells().getFirst().address())
          : Optional.empty();
    }
    if (selector instanceof RangeSelector.RectangularWindow window) {
      return Optional.of(window.topLeftAddress());
    }
    return Optional.empty();
  }

  static Optional<String> rangeFor(Selector selector) {
    if (selector instanceof RangeSelector.ByRange byRange) {
      return Optional.of(byRange.range());
    }
    if (selector instanceof RangeSelector.ByRanges byRanges) {
      return byRanges.ranges().size() == 1
          ? Optional.of(byRanges.ranges().getFirst())
          : Optional.empty();
    }
    if (selector instanceof RangeSelector.RectangularWindow window) {
      return Optional.of(window.range());
    }
    return Optional.empty();
  }

  static Optional<String> namedRangeNameFor(Selector selector) {
    if (selector instanceof NamedRangeSelector namedRangeSelector) {
      return singleNamedRangeName(namedRangeSelector);
    }
    return Optional.empty();
  }

  static Optional<String> singleSheetName(CellSelector selector) {
    return switch (selector) {
      case CellSelector.AllUsedInSheet allUsedInSheet -> Optional.of(allUsedInSheet.sheetName());
      case CellSelector.ByAddress byAddress -> Optional.of(byAddress.sheetName());
      case CellSelector.ByAddresses byAddresses -> Optional.of(byAddresses.sheetName());
      case CellSelector.ByQualifiedAddresses qualifiedAddresses ->
          qualifiedAddresses.cells().stream()
                      .map(CellSelector.QualifiedAddress::sheetName)
                      .distinct()
                      .count()
                  == 1
              ? Optional.of(qualifiedAddresses.cells().getFirst().sheetName())
              : Optional.empty();
    };
  }

  static Optional<String> singleSheetName(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet -> Optional.of(allOnSheet.sheetName());
      case RangeSelector.ByRange byRange -> Optional.of(byRange.sheetName());
      case RangeSelector.ByRanges byRanges -> Optional.of(byRanges.sheetName());
      case RangeSelector.RectangularWindow window -> Optional.of(window.sheetName());
    };
  }

  static Optional<String> singleSheetName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> Optional.empty();
      case NamedRangeSelector.ByName _ -> Optional.empty();
      case NamedRangeSelector.ByNames _ -> Optional.empty();
      case NamedRangeSelector.WorkbookScope _ -> Optional.empty();
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.sheetName());
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1
              ? singleSheetName(anyOf.selectors().getFirst())
              : Optional.empty();
    };
  }

  static Optional<String> singleSheetName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName _ -> Optional.empty();
      case NamedRangeSelector.WorkbookScope _ -> Optional.empty();
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.sheetName());
    };
  }

  static Optional<String> singleNamedRangeName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> Optional.empty();
      case NamedRangeSelector.ByName byName -> Optional.of(byName.name());
      case NamedRangeSelector.ByNames byNames ->
          byNames.names().size() == 1 ? Optional.of(byNames.names().getFirst()) : Optional.empty();
      case NamedRangeSelector.WorkbookScope workbookScope -> Optional.of(workbookScope.name());
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.name());
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1
              ? singleNamedRangeName(anyOf.selectors().getFirst())
              : Optional.empty();
    };
  }

  static Optional<String> singleNamedRangeName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> Optional.of(byName.name());
      case NamedRangeSelector.WorkbookScope workbookScope -> Optional.of(workbookScope.name());
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.name());
    };
  }

  private static Optional<String> sheetNameFor(SheetSelector selector) {
    return switch (selector) {
      case SheetSelector.All _ -> Optional.empty();
      case SheetSelector.ByName byName -> Optional.of(byName.name());
      case SheetSelector.ByNames byNames ->
          byNames.names().size() == 1 ? Optional.of(byNames.names().getFirst()) : Optional.empty();
    };
  }

  private static Optional<String> sheetNameFor(TableSelector selector) {
    return switch (selector) {
      case TableSelector.All _ -> Optional.empty();
      case TableSelector.ByName _ -> Optional.empty();
      case TableSelector.ByNames _ -> Optional.empty();
      case TableSelector.ByNameOnSheet byNameOnSheet -> Optional.of(byNameOnSheet.sheetName());
    };
  }

  private static Optional<String> sheetNameFor(PivotTableSelector selector) {
    return switch (selector) {
      case PivotTableSelector.All _ -> Optional.empty();
      case PivotTableSelector.ByName _ -> Optional.empty();
      case PivotTableSelector.ByNames _ -> Optional.empty();
      case PivotTableSelector.ByNameOnSheet byNameOnSheet -> Optional.of(byNameOnSheet.sheetName());
    };
  }

  private static Optional<String> sheetNameFor(TableRowSelector selector) {
    return switch (selector) {
      case TableRowSelector.AllRows allRows -> sheetNameFor(allRows.table());
      case TableRowSelector.ByIndex byIndex -> sheetNameFor(byIndex.table());
      case TableRowSelector.ByKeyCell byKeyCell -> sheetNameFor(byKeyCell.table());
    };
  }
}
