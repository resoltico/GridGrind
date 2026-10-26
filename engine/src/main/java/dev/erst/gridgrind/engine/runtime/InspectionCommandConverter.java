package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.InspectionAnalysisQuery;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.InspectionSurfaceQuery;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.WorkbookReadCommand;

/** Converts contract inspection steps into workbook-core read commands. */
final class InspectionCommandConverter {
  private InspectionCommandConverter() {}

  /** Converts one inspection step into the matching workbook-core read command. */
  static WorkbookReadCommand toReadCommand(InspectionStep step) {
    return toReadCommand(step.stepId(), step.target(), step.query());
  }

  /** Converts one inspection query and selector into the matching workbook-core read command. */
  static WorkbookReadCommand toReadCommand(StepId stepId, Selector target, InspectionQuery query) {
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
              exportCustomXmlMapping.mapping(),
              exportCustomXmlMapping.validateSchema(),
              exportCustomXmlMapping.encoding());
      case WorkbookIntrospectionQuery.GetNamedRanges _ ->
          new WorkbookReadCommand.GetNamedRanges(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case SheetIntrospectionQuery.GetSheetSummary _ ->
          new WorkbookReadCommand.GetSheetSummary(
              stepId, SelectorConverter.toSheetName((SheetSelector.ByName) target));
      case SheetIntrospectionQuery.GetArrayFormulas _ ->
          new WorkbookReadCommand.GetArrayFormulas(stepId, (SheetSelector) target);
      case SheetIntrospectionQuery.GetCells _ -> {
        SheetIntrospectionQuery.GetCells getCells = (SheetIntrospectionQuery.GetCells) query;
        SelectorConverter.SheetLocalCellAddresses selection =
            SelectorConverter.toSheetLocalCellAddresses((CellSelector) target);
        yield new WorkbookReadCommand.GetCells(
            stepId, selection.sheetName(), selection.addresses(), getCells.resolvedProjection());
      }
      case SheetIntrospectionQuery.GetWindow _ -> {
        SheetIntrospectionQuery.GetWindow getWindow = (SheetIntrospectionQuery.GetWindow) query;
        RangeSelector.RectangularWindow selector = (RangeSelector.RectangularWindow) target;
        yield new WorkbookReadCommand.GetWindow(
            stepId,
            selector.sheetName().value(),
            new dev.erst.gridgrind.excel.ExcelReadWindow(
                selector.topLeftAddress().value(), selector.rowCount(), selector.columnCount()),
            getWindow.resolvedProjection(),
            getWindow.includeBlanks());
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
        yield new WorkbookReadCommand.GetCharts(stepId, selection);
      }
      case WorkbookAssetIntrospectionQuery.GetPivotTables _ ->
          new WorkbookReadCommand.GetPivotTables(stepId, (PivotTableSelector) target);
      case WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload _ -> {
        DrawingObjectSelector.ByName selector = (DrawingObjectSelector.ByName) target;
        yield new WorkbookReadCommand.GetDrawingObjectPayload(
            stepId, selector.sheetName().value(), selector.objectName());
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
          new WorkbookReadCommand.GetTables(stepId, (TableSelector) target);
      case InspectionSurfaceQuery.GetFormulaSurface _ ->
          new WorkbookReadCommand.GetFormulaSurface(stepId, (SheetSelector) target);
      case InspectionSurfaceQuery.GetSheetSchema _ -> {
        InspectionSurfaceQuery.GetSheetSchema getSheetSchema =
            (InspectionSurfaceQuery.GetSheetSchema) query;
        RangeSelector.RectangularWindow selector = (RangeSelector.RectangularWindow) target;
        yield new WorkbookReadCommand.GetSheetSchema(
            stepId,
            selector.sheetName().value(),
            new dev.erst.gridgrind.excel.ExcelReadWindow(
                selector.topLeftAddress().value(), selector.rowCount(), selector.columnCount()),
            getSheetSchema.resolvedProjection());
      }
      case InspectionSurfaceQuery.GetNamedRangeSurface _ ->
          new WorkbookReadCommand.GetNamedRangeSurface(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionAnalysisQuery.AnalyzeFormulaHealth _ ->
          new WorkbookReadCommand.AnalyzeFormulaHealth(stepId, (SheetSelector) target);
      case InspectionAnalysisQuery.AnalyzeDataValidationHealth _ ->
          new WorkbookReadCommand.AnalyzeDataValidationHealth(stepId, (SheetSelector) target);
      case InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth _ ->
          new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
              stepId, (SheetSelector) target);
      case InspectionAnalysisQuery.AnalyzeAutofilterHealth _ ->
          new WorkbookReadCommand.AnalyzeAutofilterHealth(stepId, (SheetSelector) target);
      case InspectionAnalysisQuery.AnalyzeTableHealth _ ->
          new WorkbookReadCommand.AnalyzeTableHealth(stepId, (TableSelector) target);
      case InspectionAnalysisQuery.AnalyzePivotTableHealth _ ->
          new WorkbookReadCommand.AnalyzePivotTableHealth(stepId, (PivotTableSelector) target);
      case InspectionAnalysisQuery.AnalyzeHyperlinkHealth _ ->
          new WorkbookReadCommand.AnalyzeHyperlinkHealth(stepId, (SheetSelector) target);
      case InspectionAnalysisQuery.AnalyzeNamedRangeHealth _ ->
          new WorkbookReadCommand.AnalyzeNamedRangeHealth(
              stepId, SelectorConverter.toExcelNamedRangeSelection((NamedRangeSelector) target));
      case InspectionAnalysisQuery.AnalyzeWorkbookFindings _ ->
          new WorkbookReadCommand.AnalyzeWorkbookFindings(stepId);
    };
  }
}
