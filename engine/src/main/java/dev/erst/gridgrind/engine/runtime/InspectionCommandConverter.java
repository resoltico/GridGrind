package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionSurfaceQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
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
      case WorkbookIntrospectionQuery.GetWorkbookSummary _ ->
          new WorkbookReadCommand.GetWorkbookSummary(stepId);
      case WorkbookIntrospectionQuery.GetPackageSecurity _ ->
          new WorkbookReadCommand.GetPackageSecurity(stepId);
      case WorkbookIntrospectionQuery.GetWorkbookProtection _ ->
          new WorkbookReadCommand.GetWorkbookProtection(stepId);
      case WorkbookIntrospectionQuery.GetCustomXmlMappings _ ->
          new WorkbookReadCommand.GetCustomXmlMappings(stepId);
      case WorkbookIntrospectionQuery.ExportCustomXmlMapping exportCustomXmlMapping ->
          new WorkbookReadCommand.ExportCustomXmlMapping(
              stepId,
              toExcelCustomXmlMappingLocator(exportCustomXmlMapping.mapping()),
              exportCustomXmlMapping.validateSchema(),
              exportCustomXmlMapping.encoding());
      case WorkbookIntrospectionQuery.GetNamedRanges _ ->
          new WorkbookReadCommand.GetNamedRanges(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case SheetIntrospectionQuery.GetSheetSummary _ ->
          new WorkbookReadCommand.GetSheetSummary(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case SheetIntrospectionQuery.GetArrayFormulas _ ->
          new WorkbookReadCommand.GetArrayFormulas(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case SheetIntrospectionQuery.GetCells _ -> {
        SelectorConverter.SheetLocalCellAddresses selection =
            SelectorConverter.toSheetLocalCellAddresses((CellSelector) target);
        yield new WorkbookReadCommand.GetCells(
            stepId, selection.sheetName(), selection.addresses());
      }
      case SheetIntrospectionQuery.GetWindow _ -> {
        RangeSelector.RectangularWindow selector = (RangeSelector.RectangularWindow) target;
        yield new WorkbookReadCommand.GetWindow(
            stepId,
            selector.sheetName(),
            selector.topLeftAddress(),
            selector.rowCount(),
            selector.columnCount());
      }
      case SheetIntrospectionQuery.GetMergedRegions _ ->
          new WorkbookReadCommand.GetMergedRegions(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case SheetIntrospectionQuery.GetHyperlinks _ -> {
        SelectorConverter.SheetLocalCellSelection selection =
            SelectorConverter.toSheetLocalCellSelection((CellSelector) target);
        yield new WorkbookReadCommand.GetHyperlinks(
            stepId, selection.sheetName(), selection.selection());
      }
      case SheetIntrospectionQuery.GetComments _ -> {
        SelectorConverter.SheetLocalCellSelection selection =
            SelectorConverter.toSheetLocalCellSelection((CellSelector) target);
        yield new WorkbookReadCommand.GetComments(
            stepId, selection.sheetName(), selection.selection());
      }
      case WorkbookAssetIntrospectionQuery.GetDrawingObjects _ ->
          new WorkbookReadCommand.GetDrawingObjects(
              stepId, SelectorConverter.toSheetName((DrawingObjectSelector.AllOnSheet) target));
      case WorkbookAssetIntrospectionQuery.GetCharts _ -> {
        if (!(target instanceof ChartSelector selection)) {
          throw new IllegalArgumentException("Unsupported chart inspection target");
        }
        yield new WorkbookReadCommand.GetCharts(
            stepId, SelectorConverter.toExcelChartSelection(selection));
      }
      case WorkbookAssetIntrospectionQuery.GetPivotTables _ ->
          new WorkbookReadCommand.GetPivotTables(
              stepId, SelectorConverter.toExcelPivotTableSelection((PivotTableSelector) target));
      case WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload _ -> {
        DrawingObjectSelector.ByName selector = (DrawingObjectSelector.ByName) target;
        yield new WorkbookReadCommand.GetDrawingObjectPayload(
            stepId, selector.sheetName(), selector.objectName());
      }
      case SheetIntrospectionQuery.GetSheetLayout _ ->
          new WorkbookReadCommand.GetSheetLayout(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case SheetIntrospectionQuery.GetPrintLayout _ ->
          new WorkbookReadCommand.GetPrintLayout(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case SheetIntrospectionQuery.GetDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookReadCommand.GetDataValidations(
            stepId, selection.sheetName(), selection.selection());
      }
      case SheetIntrospectionQuery.GetConditionalFormatting _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookReadCommand.GetConditionalFormatting(
            stepId, selection.sheetName(), selection.selection());
      }
      case SheetIntrospectionQuery.GetAutofilters _ ->
          new WorkbookReadCommand.GetAutofilters(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case WorkbookAssetIntrospectionQuery.GetTables _ ->
          new WorkbookReadCommand.GetTables(
              stepId, SelectorConverter.toExcelTableSelection((TableSelector) target));
      case InspectionSurfaceQuery.GetFormulaSurface _ ->
          new WorkbookReadCommand.GetFormulaSurface(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionSurfaceQuery.GetSheetSchema _ -> {
        RangeSelector.RectangularWindow selector = (RangeSelector.RectangularWindow) target;
        yield new WorkbookReadCommand.GetSheetSchema(
            stepId,
            selector.sheetName(),
            selector.topLeftAddress(),
            selector.rowCount(),
            selector.columnCount());
      }
      case InspectionSurfaceQuery.GetNamedRangeSurface _ ->
          new WorkbookReadCommand.GetNamedRangeSurface(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionAnalysisQuery.AnalyzeFormulaHealth _ ->
          new WorkbookReadCommand.AnalyzeFormulaHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionAnalysisQuery.AnalyzeDataValidationHealth _ ->
          new WorkbookReadCommand.AnalyzeDataValidationHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth _ ->
          new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionAnalysisQuery.AnalyzeAutofilterHealth _ ->
          new WorkbookReadCommand.AnalyzeAutofilterHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionAnalysisQuery.AnalyzeTableHealth _ ->
          new WorkbookReadCommand.AnalyzeTableHealth(
              stepId, SelectorConverter.toExcelTableSelection((TableSelector) target));
      case InspectionAnalysisQuery.AnalyzePivotTableHealth _ ->
          new WorkbookReadCommand.AnalyzePivotTableHealth(
              stepId, SelectorConverter.toExcelPivotTableSelection((PivotTableSelector) target));
      case InspectionAnalysisQuery.AnalyzeHyperlinkHealth _ ->
          new WorkbookReadCommand.AnalyzeHyperlinkHealth(
              stepId, SelectorConverter.toExcelSheetSelection((SheetSelector) target));
      case InspectionAnalysisQuery.AnalyzeNamedRangeHealth _ ->
          new WorkbookReadCommand.AnalyzeNamedRangeHealth(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionAnalysisQuery.AnalyzeWorkbookFindings _ ->
          new WorkbookReadCommand.AnalyzeWorkbookFindings(stepId);
    };
  }

  private static dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator
      toExcelCustomXmlMappingLocator(
          dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator locator) {
    return new dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator(
        locator.mapId(), locator.name());
  }
}
