package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelCellSelection;
import dev.erst.gridgrind.excel.ExcelColumnSpan;
import dev.erst.gridgrind.excel.ExcelFormulaCellTarget;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelection;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelSheetSelection;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import java.util.List;

/** Central conversion seam between contract selectors and workbook-core selector primitives. */
final class SelectorConverter {
  private SelectorConverter() {}

  static String toSheetName(SheetSelector.ByName selector) {
    return selector.name();
  }

  static String toSheetName(DrawingObjectSelector.AllOnSheet selector) {
    return selector.sheetName();
  }

  static String toSheetName(ChartSelector.AllOnSheet selector) {
    return selector.sheetName();
  }

  static String toSheetName(DrawingObjectSelector.ByName selector) {
    return selector.sheetName();
  }

  static ExcelSheetSelection toExcelSheetSelection(SheetSelector selector) {
    return switch (selector) {
      case SheetSelector.All _ -> new ExcelSheetSelection.All();
      case SheetSelector.ByName byName -> new ExcelSheetSelection.Selected(List.of(byName.name()));
      case SheetSelector.ByNames byNames -> new ExcelSheetSelection.Selected(byNames.names());
    };
  }

  static ExcelNamedRangeSelection toExcelNamedRangeSelection(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> new ExcelNamedRangeSelection.All();
      case NamedRangeSelector.AnyOf anyOf ->
          new ExcelNamedRangeSelection.Selected(
              anyOf.selectors().stream()
                  .map(SelectorConverter::toExcelNamedRangeSelector)
                  .toList());
      case NamedRangeSelector.ByName byName ->
          new ExcelNamedRangeSelection.Selected(
              List.of(new ExcelNamedRangeSelector.ByName(byName.name())));
      case NamedRangeSelector.ByNames byNames ->
          new ExcelNamedRangeSelection.Selected(
              byNames.names().stream()
                  .map(name -> (ExcelNamedRangeSelector) new ExcelNamedRangeSelector.ByName(name))
                  .toList());
      case NamedRangeSelector.WorkbookScope workbookScope ->
          new ExcelNamedRangeSelection.Selected(
              List.of(new ExcelNamedRangeSelector.WorkbookScope(workbookScope.name())));
      case NamedRangeSelector.SheetScope sheetScope ->
          new ExcelNamedRangeSelection.Selected(
              List.of(
                  new ExcelNamedRangeSelector.SheetScope(
                      sheetScope.name(), sheetScope.sheetName())));
    };
  }

  private static ExcelNamedRangeSelector toExcelNamedRangeSelector(
      NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> new ExcelNamedRangeSelector.ByName(byName.name());
      case NamedRangeSelector.WorkbookScope workbookScope ->
          new ExcelNamedRangeSelector.WorkbookScope(workbookScope.name());
      case NamedRangeSelector.SheetScope sheetScope ->
          new ExcelNamedRangeSelector.SheetScope(sheetScope.name(), sheetScope.sheetName());
    };
  }

  static ExcelTableSelection toExcelTableSelection(TableSelector selector) {
    return switch (selector) {
      case TableSelector.All _ -> new ExcelTableSelection.All();
      case TableSelector.ByName byName -> new ExcelTableSelection.ByNames(List.of(byName.name()));
      case TableSelector.ByNames byNames -> new ExcelTableSelection.ByNames(byNames.names());
      case TableSelector.ByNameOnSheet byNameOnSheet ->
          new ExcelTableSelection.ByNames(List.of(byNameOnSheet.name()));
    };
  }

  static ExcelPivotTableSelection toExcelPivotTableSelection(PivotTableSelector selector) {
    return switch (selector) {
      case PivotTableSelector.All _ -> new ExcelPivotTableSelection.All();
      case PivotTableSelector.ByName byName ->
          new ExcelPivotTableSelection.ByNames(List.of(byName.name()));
      case PivotTableSelector.ByNames byNames ->
          new ExcelPivotTableSelection.ByNames(byNames.names());
      case PivotTableSelector.ByNameOnSheet byNameOnSheet ->
          new ExcelPivotTableSelection.ByNames(List.of(byNameOnSheet.name()));
    };
  }

  static SheetLocalCellSelection toSheetLocalCellSelection(CellSelector selector) {
    return switch (selector) {
      case CellSelector.AllUsedInSheet all ->
          new SheetLocalCellSelection(all.sheetName(), new ExcelCellSelection.AllUsedCells());
      case CellSelector.ByAddress byAddress ->
          new SheetLocalCellSelection(
              byAddress.sheetName(), new ExcelCellSelection.Selected(List.of(byAddress.address())));
      case CellSelector.ByAddresses byAddresses ->
          new SheetLocalCellSelection(
              byAddresses.sheetName(), new ExcelCellSelection.Selected(byAddresses.addresses()));
      case CellSelector.ByQualifiedAddresses byQualifiedAddresses ->
          throw new IllegalArgumentException(
              "selector must target one sheet; cross-sheet exact cell selectors are not supported here: "
                  + byQualifiedAddresses);
    };
  }

  static SheetLocalCellAddresses toSheetLocalCellAddresses(CellSelector selector) {
    return switch (selector) {
      case CellSelector.ByAddress byAddress ->
          new SheetLocalCellAddresses(byAddress.sheetName(), List.of(byAddress.address()));
      case CellSelector.ByAddresses byAddresses ->
          new SheetLocalCellAddresses(byAddresses.sheetName(), byAddresses.addresses());
      case CellSelector.AllUsedInSheet _ ->
          throw new IllegalArgumentException(
              "selector must provide exact addresses here; ALL_USED_IN_SHEET is not supported");
      case CellSelector.ByQualifiedAddresses _ ->
          throw new IllegalArgumentException(
              "selector must target one sheet with exact addresses here");
    };
  }

  static QualifiedCellAddresses toQualifiedCellAddresses(
      CellSelector.ByQualifiedAddresses selector) {
    return new QualifiedCellAddresses(
        selector.cells().stream()
            .map(cell -> new ExcelFormulaCellTarget(cell.sheetName(), cell.address()))
            .toList());
  }

  static SingleCellTarget toSingleCellTarget(CellSelector.ByAddress selector) {
    return new SingleCellTarget(selector.sheetName(), selector.address());
  }

  static SheetLocalRangeSelection toSheetLocalRangeSelection(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet ->
          new SheetLocalRangeSelection(allOnSheet.sheetName(), new ExcelRangeSelection.All());
      case RangeSelector.ByRange byRange ->
          new SheetLocalRangeSelection(
              byRange.sheetName(), new ExcelRangeSelection.Selected(List.of(byRange.range())));
      case RangeSelector.ByRanges byRanges ->
          new SheetLocalRangeSelection(
              byRanges.sheetName(), new ExcelRangeSelection.Selected(byRanges.ranges()));
      case RangeSelector.RectangularWindow window ->
          new SheetLocalRangeSelection(
              window.sheetName(), new ExcelRangeSelection.Selected(List.of(window.range())));
    };
  }

  static SingleRangeTarget toSingleRangeTarget(RangeSelector.ByRange selector) {
    return new SingleRangeTarget(selector.sheetName(), selector.range());
  }

  static SingleRangeTarget toSingleRangeTarget(RangeSelector.RectangularWindow selector) {
    return new SingleRangeTarget(selector.sheetName(), selector.range());
  }

  static String toSheetName(RangeSelector.AllOnSheet selector) {
    return selector.sheetName();
  }

  static String toSheetName(RowBandSelector.Span selector) {
    return selector.sheetName();
  }

  static String toSheetName(RowBandSelector.Insertion selector) {
    return selector.sheetName();
  }

  static String toSheetName(ColumnBandSelector.Span selector) {
    return selector.sheetName();
  }

  static String toSheetName(ColumnBandSelector.Insertion selector) {
    return selector.sheetName();
  }

  static ExcelRowSpan toExcelRowSpan(RowBandSelector.Span selector) {
    return new ExcelRowSpan(selector.firstRowIndex(), selector.lastRowIndex());
  }

  static ExcelColumnSpan toExcelColumnSpan(ColumnBandSelector.Span selector) {
    return new ExcelColumnSpan(selector.firstColumnIndex(), selector.lastColumnIndex());
  }

  static String toSheetName(TableSelector.ByNameOnSheet selector) {
    return selector.sheetName();
  }

  static String toSheetName(PivotTableSelector.ByNameOnSheet selector) {
    return selector.sheetName();
  }

  record SheetLocalCellSelection(String sheetName, ExcelCellSelection selection) {}

  record SheetLocalCellAddresses(String sheetName, List<String> addresses) {}

  record SheetLocalRangeSelection(String sheetName, ExcelRangeSelection selection) {}

  record SingleCellTarget(String sheetName, String address) {}

  record SingleRangeTarget(String sheetName, String range) {}

  record QualifiedCellAddresses(List<ExcelFormulaCellTarget> cells) {}
}
