package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelCellSelection;
import dev.erst.gridgrind.excel.ExcelChartSelection;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelection;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelSheetSelection;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import java.util.List;

/** Central conversion seam between contract selectors and workbook-core selector primitives. */
final class SelectorConverter {
  private SelectorConverter() {}

  static String toSheetName(SheetSelector.ByName selector) {
    return selector.name().value();
  }

  static String toSheetName(DrawingObjectSelector.AllOnSheet selector) {
    return selector.sheetName().value();
  }

  static String toSheetName(DrawingObjectSelector.ByName selector) {
    return selector.sheetName().value();
  }

  static ExcelChartSelection toExcelChartSelection(ChartSelector selector) {
    return switch (selector) {
      case ChartSelector.AllOnSheet allOnSheet ->
          new ExcelChartSelection.AllOnSheet(allOnSheet.sheetName().value());
      case ChartSelector.ByName byName ->
          new ExcelChartSelection.ByName(byName.sheetName().value(), byName.chartName());
    };
  }

  static ExcelSheetSelection toExcelSheetSelection(SheetSelector selector) {
    return switch (selector) {
      case SheetSelector.All _ -> new ExcelSheetSelection.All();
      case SheetSelector.ByName byName ->
          new ExcelSheetSelection.Selected(List.of(byName.name().value()));
      case SheetSelector.ByNames byNames ->
          new ExcelSheetSelection.Selected(
              byNames.names().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                  .toList());
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
                      sheetScope.name(), sheetScope.sheetName().value())));
    };
  }

  private static ExcelNamedRangeSelector toExcelNamedRangeSelector(
      NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> new ExcelNamedRangeSelector.ByName(byName.name());
      case NamedRangeSelector.WorkbookScope workbookScope ->
          new ExcelNamedRangeSelector.WorkbookScope(workbookScope.name());
      case NamedRangeSelector.SheetScope sheetScope ->
          new ExcelNamedRangeSelector.SheetScope(sheetScope.name(), sheetScope.sheetName().value());
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
          new SheetLocalCellSelection(
              all.sheetName().value(), new ExcelCellSelection.AllUsedCells());
      case CellSelector.ByAddress byAddress ->
          new SheetLocalCellSelection(
              byAddress.sheetName().value(),
              new ExcelCellSelection.Selected(
                  List.of(byAddress.address()).stream()
                      .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                      .toList()));
      case CellSelector.ByAddresses byAddresses ->
          new SheetLocalCellSelection(
              byAddresses.sheetName().value(),
              new ExcelCellSelection.Selected(
                  byAddresses.addresses().stream()
                      .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                      .toList()));
    };
  }

  static SheetLocalCellAddresses toSheetLocalCellAddresses(CellSelector selector) {
    return switch (selector) {
      case CellSelector.ByAddress byAddress ->
          new SheetLocalCellAddresses(
              byAddress.sheetName().value(),
              List.of(byAddress.address()).stream()
                  .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                  .toList());
      case CellSelector.ByAddresses byAddresses ->
          new SheetLocalCellAddresses(
              byAddresses.sheetName().value(),
              byAddresses.addresses().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                  .toList());
      case CellSelector.AllUsedInSheet _ ->
          throw new IllegalArgumentException(
              "selector must provide exact addresses here; ALL_USED_IN_SHEET is not supported");
    };
  }

  static SingleCellTarget toSingleCellTarget(CellSelector.ByAddress selector) {
    return new SingleCellTarget(selector.sheetName().value(), selector.address().value());
  }

  static SheetLocalRangeSelection toSheetLocalRangeSelection(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet ->
          new SheetLocalRangeSelection(
              allOnSheet.sheetName().value(), new ExcelRangeSelection.All());
      case RangeSelector.ByRange byRange ->
          new SheetLocalRangeSelection(
              byRange.sheetName().value(),
              new ExcelRangeSelection.Selected(
                  List.of(byRange.range()).stream()
                      .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                      .toList()));
      case RangeSelector.ByRanges byRanges ->
          new SheetLocalRangeSelection(
              byRanges.sheetName().value(),
              new ExcelRangeSelection.Selected(
                  byRanges.ranges().stream()
                      .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                      .toList()));
      case RangeSelector.RectangularWindow window ->
          new SheetLocalRangeSelection(
              window.sheetName().value(),
              new ExcelRangeSelection.Selected(
                  List.of(window.range()).stream()
                      .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                      .toList()));
    };
  }

  static SingleRangeTarget toSingleRangeTarget(RangeSelector.ByRange selector) {
    return new SingleRangeTarget(selector.sheetName().value(), selector.range().value());
  }

  static SingleRangeTarget toSingleRangeTarget(RangeSelector.RectangularWindow selector) {
    return new SingleRangeTarget(selector.sheetName().value(), selector.range().value());
  }

  static String toSheetName(RangeSelector.AllOnSheet selector) {
    return selector.sheetName().value();
  }

  static String toSheetName(RowBandSelector.Span selector) {
    return selector.sheetName().value();
  }

  static String toSheetName(RowBandSelector.Insertion selector) {
    return selector.sheetName().value();
  }

  static String toSheetName(ColumnBandSelector.Span selector) {
    return selector.sheetName().value();
  }

  static String toSheetName(ColumnBandSelector.Insertion selector) {
    return selector.sheetName().value();
  }

  static ExcelRowSpan toExcelRowSpan(RowBandSelector.Span selector) {
    return new ExcelRowSpan(selector.firstRowIndex(), selector.lastRowIndex());
  }

  static ExcelColumnSpan toExcelColumnSpan(ColumnBandSelector.Span selector) {
    return new ExcelColumnSpan(selector.firstColumnIndex(), selector.lastColumnIndex());
  }

  static String toSheetName(TableSelector.ByNameOnSheet selector) {
    return selector.sheetName().value();
  }

  static String toSheetName(PivotTableSelector.ByNameOnSheet selector) {
    return selector.sheetName().value();
  }

  record SheetLocalCellSelection(String sheetName, ExcelCellSelection selection) {}

  record SheetLocalCellAddresses(String sheetName, List<String> addresses) {
    SheetLocalCellAddresses {
      addresses = List.copyOf(addresses);
    }
  }

  record SheetLocalRangeSelection(String sheetName, ExcelRangeSelection selection) {}

  record SingleCellTarget(String sheetName, String address) {}

  record SingleRangeTarget(String sheetName, String range) {}
}
