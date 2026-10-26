package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelCellSelection;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
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
