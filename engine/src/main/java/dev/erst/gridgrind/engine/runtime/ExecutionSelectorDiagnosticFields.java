package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import java.util.Optional;

/** Selector diagnostic field extraction for execution journal contexts. */
final class ExecutionSelectorDiagnosticFields {
  private ExecutionSelectorDiagnosticFields() {}

  static Optional<String> sheetNameFor(Selector selector) {
    return switch (selector) {
      case WorkbookSelector _ -> Optional.empty();
      case SheetSelector sheetSelector -> sheetNameFor(sheetSelector);
      case CellSelector cellSelector -> singleSheetName(cellSelector);
      case RangeSelector rangeSelector -> singleSheetName(rangeSelector);
      case RowBandSelector.Span span -> Optional.of(span.sheetName().value());
      case RowBandSelector.Insertion insertion -> Optional.of(insertion.sheetName().value());
      case ColumnBandSelector.Span span -> Optional.of(span.sheetName().value());
      case ColumnBandSelector.Insertion insertion -> Optional.of(insertion.sheetName().value());
      case DrawingObjectSelector.AllOnSheet allOnSheet ->
          Optional.of(allOnSheet.sheetName().value());
      case DrawingObjectSelector.ByName byName -> Optional.of(byName.sheetName().value());
      case ChartSelector.AllOnSheet allOnSheet -> Optional.of(allOnSheet.sheetName().value());
      case ChartSelector.ByName byName -> Optional.of(byName.sheetName().value());
      case TableSelector tableSelector -> sheetNameFor(tableSelector);
      case PivotTableSelector pivotTableSelector -> sheetNameFor(pivotTableSelector);
      case NamedRangeSelector namedRangeSelector -> singleSheetName(namedRangeSelector);
      case TableRowSelector tableRowSelector -> sheetNameFor(tableRowSelector);
      case TableCellSelector.ByColumnName byColumnName -> sheetNameFor(byColumnName.row());
    };
  }

  static Optional<String> addressFor(Selector selector) {
    if (selector instanceof CellSelector.ByAddress byAddress) {
      return Optional.of(byAddress.address().value());
    }
    if (selector instanceof RangeSelector.RectangularWindow window) {
      return Optional.of(window.topLeftAddress().value());
    }
    return Optional.empty();
  }

  static Optional<String> rangeFor(Selector selector) {
    if (selector instanceof RangeSelector.ByRange byRange) {
      return Optional.of(byRange.range().value());
    }
    if (selector instanceof RangeSelector.ByRanges byRanges) {
      return byRanges.ranges().size() == 1
          ? Optional.of(byRanges.ranges().getFirst().value())
          : Optional.empty();
    }
    if (selector instanceof RangeSelector.RectangularWindow window) {
      return Optional.of(window.range().value());
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
      case CellSelector.AllUsedInSheet allUsedInSheet ->
          Optional.of(allUsedInSheet.sheetName().value());
      case CellSelector.ByAddress byAddress -> Optional.of(byAddress.sheetName().value());
      case CellSelector.ByAddresses byAddresses -> Optional.of(byAddresses.sheetName().value());
    };
  }

  static Optional<String> singleSheetName(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet -> Optional.of(allOnSheet.sheetName().value());
      case RangeSelector.ByRange byRange -> Optional.of(byRange.sheetName().value());
      case RangeSelector.ByRanges byRanges -> Optional.of(byRanges.sheetName().value());
      case RangeSelector.RectangularWindow window -> Optional.of(window.sheetName().value());
    };
  }

  static Optional<String> singleSheetName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> Optional.empty();
      case NamedRangeSelector.ByName _ -> Optional.empty();
      case NamedRangeSelector.ByNames _ -> Optional.empty();
      case NamedRangeSelector.WorkbookScope _ -> Optional.empty();
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.sheetName().value());
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
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.sheetName().value());
    };
  }

  static Optional<String> singleNamedRangeName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> Optional.empty();
      case NamedRangeSelector.ByName byName -> Optional.of(byName.name().value());
      case NamedRangeSelector.ByNames byNames ->
          byNames.names().size() == 1
              ? Optional.of(byNames.names().getFirst().value())
              : Optional.empty();
      case NamedRangeSelector.WorkbookScope workbookScope ->
          Optional.of(workbookScope.name().value());
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.name().value());
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1
              ? singleNamedRangeName(anyOf.selectors().getFirst())
              : Optional.empty();
    };
  }

  static Optional<String> singleNamedRangeName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> Optional.of(byName.name().value());
      case NamedRangeSelector.WorkbookScope workbookScope ->
          Optional.of(workbookScope.name().value());
      case NamedRangeSelector.SheetScope sheetScope -> Optional.of(sheetScope.name().value());
    };
  }

  private static Optional<String> sheetNameFor(SheetSelector selector) {
    return switch (selector) {
      case SheetSelector.All _ -> Optional.empty();
      case SheetSelector.ByName byName -> Optional.of(byName.name().value());
      case SheetSelector.ByNames byNames ->
          byNames.names().size() == 1
              ? Optional.of(byNames.names().getFirst().value())
              : Optional.empty();
    };
  }

  private static Optional<String> sheetNameFor(TableSelector selector) {
    return switch (selector) {
      case TableSelector.All _ -> Optional.empty();
      case TableSelector.ByName _ -> Optional.empty();
      case TableSelector.ByNames _ -> Optional.empty();
      case TableSelector.ByNameOnSheet byNameOnSheet ->
          Optional.of(byNameOnSheet.sheetName().value());
    };
  }

  private static Optional<String> sheetNameFor(PivotTableSelector selector) {
    return switch (selector) {
      case PivotTableSelector.All _ -> Optional.empty();
      case PivotTableSelector.ByName _ -> Optional.empty();
      case PivotTableSelector.ByNames _ -> Optional.empty();
      case PivotTableSelector.ByNameOnSheet byNameOnSheet ->
          Optional.of(byNameOnSheet.sheetName().value());
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
