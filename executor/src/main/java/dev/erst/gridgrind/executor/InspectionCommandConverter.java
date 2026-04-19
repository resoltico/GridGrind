package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.WorkbookReadCommand;

/** Converts contract inspection steps into workbook-core read commands. */
final class InspectionCommandConverter {
  private InspectionCommandConverter() {}

  /** Converts one inspection step into the matching workbook-core read command. */
  static WorkbookReadCommand toReadCommand(InspectionStep step) {
    return toReadCommand(step.stepId(), step.target(), step.query());
  }

  /** Converts one inspection query and selector into the matching workbook-core read command. */
  static WorkbookReadCommand toReadCommand(String stepId, Selector target, InspectionQuery query) {
    return switch (query) {
      case InspectionQuery.GetWorkbookSummary _ ->
          new WorkbookReadCommand.GetWorkbookSummary(stepId);
      case InspectionQuery.GetPackageSecurity _ ->
          new WorkbookReadCommand.GetPackageSecurity(stepId);
      case InspectionQuery.GetWorkbookProtection _ ->
          new WorkbookReadCommand.GetWorkbookProtection(stepId);
      case InspectionQuery.GetNamedRanges _ ->
          new WorkbookReadCommand.GetNamedRanges(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionQuery.GetSheetSummary _ ->
          new WorkbookReadCommand.GetSheetSummary(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case InspectionQuery.GetCells _ -> {
        SelectorConverter.SheetLocalCellAddresses selection =
            SelectorConverter.toSheetLocalCellAddresses((CellSelector) target);
        yield new WorkbookReadCommand.GetCells(
            stepId, selection.sheetName(), selection.addresses());
      }
      case InspectionQuery.GetWindow _ -> {
        RangeSelector.RectangularWindow selector = (RangeSelector.RectangularWindow) target;
        yield new WorkbookReadCommand.GetWindow(
            stepId,
            selector.sheetName(),
            selector.topLeftAddress(),
            selector.rowCount(),
            selector.columnCount());
      }
      case InspectionQuery.GetMergedRegions _ ->
          new WorkbookReadCommand.GetMergedRegions(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case InspectionQuery.GetHyperlinks _ -> {
        SelectorConverter.SheetLocalCellSelection selection =
            SelectorConverter.toSheetLocalCellSelection((CellSelector) target);
        yield new WorkbookReadCommand.GetHyperlinks(
            stepId, selection.sheetName(), selection.selection());
      }
      case InspectionQuery.GetComments _ -> {
        SelectorConverter.SheetLocalCellSelection selection =
            SelectorConverter.toSheetLocalCellSelection((CellSelector) target);
        yield new WorkbookReadCommand.GetComments(
            stepId, selection.sheetName(), selection.selection());
      }
      case InspectionQuery.GetDrawingObjects _ ->
          new WorkbookReadCommand.GetDrawingObjects(
              stepId, SelectorConverter.toSheetName((DrawingObjectSelector.AllOnSheet) target));
      case InspectionQuery.GetCharts _ -> {
        String sheetName =
            switch (target) {
              case ChartSelector.AllOnSheet selector -> selector.sheetName();
              case SheetSelector.ByName selector -> selector.name();
              default -> throw new IllegalArgumentException("Unsupported chart inspection target");
            };
        yield new WorkbookReadCommand.GetCharts(stepId, sheetName);
      }
      case InspectionQuery.GetPivotTables _ ->
          new WorkbookReadCommand.GetPivotTables(
              stepId, SelectorConverter.toExcelPivotTableSelection((PivotTableSelector) target));
      case InspectionQuery.GetDrawingObjectPayload _ -> {
        DrawingObjectSelector.ByName selector = (DrawingObjectSelector.ByName) target;
        yield new WorkbookReadCommand.GetDrawingObjectPayload(
            stepId, selector.sheetName(), selector.objectName());
      }
      case InspectionQuery.GetSheetLayout _ ->
          new WorkbookReadCommand.GetSheetLayout(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case InspectionQuery.GetPrintLayout _ ->
          new WorkbookReadCommand.GetPrintLayout(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case InspectionQuery.GetDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookReadCommand.GetDataValidations(
            stepId, selection.sheetName(), selection.selection());
      }
      case InspectionQuery.GetConditionalFormatting _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookReadCommand.GetConditionalFormatting(
            stepId, selection.sheetName(), selection.selection());
      }
      case InspectionQuery.GetAutofilters _ ->
          new WorkbookReadCommand.GetAutofilters(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case InspectionQuery.GetTables _ ->
          new WorkbookReadCommand.GetTables(
              stepId, SelectorConverter.toExcelTableSelection((TableSelector) target));
      case InspectionQuery.GetFormulaSurface _ ->
          new WorkbookReadCommand.GetFormulaSurface(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionQuery.GetSheetSchema _ -> {
        RangeSelector.RectangularWindow selector = (RangeSelector.RectangularWindow) target;
        yield new WorkbookReadCommand.GetSheetSchema(
            stepId,
            selector.sheetName(),
            selector.topLeftAddress(),
            selector.rowCount(),
            selector.columnCount());
      }
      case InspectionQuery.GetNamedRangeSurface _ ->
          new WorkbookReadCommand.GetNamedRangeSurface(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionQuery.AnalyzeFormulaHealth _ ->
          new WorkbookReadCommand.AnalyzeFormulaHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionQuery.AnalyzeDataValidationHealth _ ->
          new WorkbookReadCommand.AnalyzeDataValidationHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionQuery.AnalyzeConditionalFormattingHealth _ ->
          new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionQuery.AnalyzeAutofilterHealth _ ->
          new WorkbookReadCommand.AnalyzeAutofilterHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionQuery.AnalyzeTableHealth _ ->
          new WorkbookReadCommand.AnalyzeTableHealth(
              stepId, SelectorConverter.toExcelTableSelection((TableSelector) target));
      case InspectionQuery.AnalyzePivotTableHealth _ ->
          new WorkbookReadCommand.AnalyzePivotTableHealth(
              stepId, SelectorConverter.toExcelPivotTableSelection((PivotTableSelector) target));
      case InspectionQuery.AnalyzeHyperlinkHealth _ ->
          new WorkbookReadCommand.AnalyzeHyperlinkHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionQuery.AnalyzeNamedRangeHealth _ ->
          new WorkbookReadCommand.AnalyzeNamedRangeHealth(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionQuery.AnalyzeWorkbookFindings _ ->
          new WorkbookReadCommand.AnalyzeWorkbookFindings(stepId);
    };
  }
}
