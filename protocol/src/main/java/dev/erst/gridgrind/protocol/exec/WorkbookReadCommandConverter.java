package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;

/** Converts protocol read operations into workbook-core read commands. */
final class WorkbookReadCommandConverter {
  private WorkbookReadCommandConverter() {}

  /** Converts one protocol read into the matching workbook-core read command. */
  static WorkbookReadCommand toReadCommand(WorkbookReadOperation read) {
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary op ->
          new WorkbookReadCommand.GetWorkbookSummary(op.requestId());
      case WorkbookReadOperation.GetWorkbookProtection op ->
          new WorkbookReadCommand.GetWorkbookProtection(op.requestId());
      case WorkbookReadOperation.GetNamedRanges op ->
          new WorkbookReadCommand.GetNamedRanges(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
      case WorkbookReadOperation.GetSheetSummary op ->
          new WorkbookReadCommand.GetSheetSummary(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetCells op ->
          new WorkbookReadCommand.GetCells(op.requestId(), op.sheetName(), op.addresses());
      case WorkbookReadOperation.GetWindow op ->
          new WorkbookReadCommand.GetWindow(
              op.requestId(), op.sheetName(), op.topLeftAddress(), op.rowCount(), op.columnCount());
      case WorkbookReadOperation.GetMergedRegions op ->
          new WorkbookReadCommand.GetMergedRegions(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetHyperlinks op ->
          new WorkbookReadCommand.GetHyperlinks(
              op.requestId(), op.sheetName(), toExcelCellSelection(op.selection()));
      case WorkbookReadOperation.GetComments op ->
          new WorkbookReadCommand.GetComments(
              op.requestId(), op.sheetName(), toExcelCellSelection(op.selection()));
      case WorkbookReadOperation.GetDrawingObjects op ->
          new WorkbookReadCommand.GetDrawingObjects(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetCharts op ->
          new WorkbookReadCommand.GetCharts(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetPivotTables op ->
          new WorkbookReadCommand.GetPivotTables(
              op.requestId(), toExcelPivotTableSelection(op.selection()));
      case WorkbookReadOperation.GetDrawingObjectPayload op ->
          new WorkbookReadCommand.GetDrawingObjectPayload(
              op.requestId(), op.sheetName(), op.objectName());
      case WorkbookReadOperation.GetSheetLayout op ->
          new WorkbookReadCommand.GetSheetLayout(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetPrintLayout op ->
          new WorkbookReadCommand.GetPrintLayout(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetDataValidations op ->
          new WorkbookReadCommand.GetDataValidations(
              op.requestId(), op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookReadOperation.GetConditionalFormatting op ->
          new WorkbookReadCommand.GetConditionalFormatting(
              op.requestId(), op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookReadOperation.GetAutofilters op ->
          new WorkbookReadCommand.GetAutofilters(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetTables op ->
          new WorkbookReadCommand.GetTables(op.requestId(), toExcelTableSelection(op.selection()));
      case WorkbookReadOperation.GetFormulaSurface op ->
          new WorkbookReadCommand.GetFormulaSurface(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.GetSheetSchema op ->
          new WorkbookReadCommand.GetSheetSchema(
              op.requestId(), op.sheetName(), op.topLeftAddress(), op.rowCount(), op.columnCount());
      case WorkbookReadOperation.GetNamedRangeSurface op ->
          new WorkbookReadCommand.GetNamedRangeSurface(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeFormulaHealth op ->
          new WorkbookReadCommand.AnalyzeFormulaHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeDataValidationHealth op ->
          new WorkbookReadCommand.AnalyzeDataValidationHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth op ->
          new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeAutofilterHealth op ->
          new WorkbookReadCommand.AnalyzeAutofilterHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeTableHealth op ->
          new WorkbookReadCommand.AnalyzeTableHealth(
              op.requestId(), toExcelTableSelection(op.selection()));
      case WorkbookReadOperation.AnalyzePivotTableHealth op ->
          new WorkbookReadCommand.AnalyzePivotTableHealth(
              op.requestId(), toExcelPivotTableSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeHyperlinkHealth op ->
          new WorkbookReadCommand.AnalyzeHyperlinkHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeNamedRangeHealth op ->
          new WorkbookReadCommand.AnalyzeNamedRangeHealth(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeWorkbookFindings op ->
          new WorkbookReadCommand.AnalyzeWorkbookFindings(op.requestId());
    };
  }

  static ExcelTableSelection toExcelTableSelection(TableSelection selection) {
    return switch (selection) {
      case TableSelection.All _ -> new ExcelTableSelection.All();
      case TableSelection.ByNames byNames -> new ExcelTableSelection.ByNames(byNames.names());
    };
  }

  static ExcelPivotTableSelection toExcelPivotTableSelection(PivotTableSelection selection) {
    return switch (selection) {
      case PivotTableSelection.All _ -> new ExcelPivotTableSelection.All();
      case PivotTableSelection.ByNames byNames ->
          new ExcelPivotTableSelection.ByNames(byNames.names());
    };
  }

  private static ExcelNamedRangeSelection toExcelNamedRangeSelection(
      NamedRangeSelection selection) {
    return switch (selection) {
      case NamedRangeSelection.All _ -> new ExcelNamedRangeSelection.All();
      case NamedRangeSelection.Selected selected ->
          new ExcelNamedRangeSelection.Selected(
              selected.selectors().stream()
                  .map(WorkbookReadCommandConverter::toExcelNamedRangeSelector)
                  .toList());
    };
  }

  private static ExcelNamedRangeSelector toExcelNamedRangeSelector(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> new ExcelNamedRangeSelector.ByName(byName.name());
      case NamedRangeSelector.WorkbookScope workbookScope ->
          new ExcelNamedRangeSelector.WorkbookScope(workbookScope.name());
      case NamedRangeSelector.SheetScope sheetScope ->
          new ExcelNamedRangeSelector.SheetScope(sheetScope.name(), sheetScope.sheetName());
    };
  }

  private static ExcelSheetSelection toExcelSheetSelection(SheetSelection selection) {
    return switch (selection) {
      case SheetSelection.All _ -> new ExcelSheetSelection.All();
      case SheetSelection.Selected selected ->
          new ExcelSheetSelection.Selected(selected.sheetNames());
    };
  }

  private static ExcelRangeSelection toExcelRangeSelection(RangeSelection selection) {
    return switch (selection) {
      case RangeSelection.All _ -> new ExcelRangeSelection.All();
      case RangeSelection.Selected selected -> new ExcelRangeSelection.Selected(selected.ranges());
    };
  }

  private static ExcelCellSelection toExcelCellSelection(CellSelection selection) {
    return switch (selection) {
      case CellSelection.AllUsedCells _ -> new ExcelCellSelection.AllUsedCells();
      case CellSelection.Selected selected -> new ExcelCellSelection.Selected(selected.addresses());
    };
  }
}
