package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAnalysisResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;

/** Converts engine read results into protocol inspection-result variants. */
final class InspectionResultReadConverter {
  private InspectionResultReadConverter() {}

  static InspectionResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult workbookSummary ->
          new WorkbookInspectionResult.WorkbookSummaryResult(
              workbookSummary.stepId(), workbookSummary.workbook());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.PackageSecurityResult packageSecurity ->
          new WorkbookInspectionResult.PackageSecurityResult(
              packageSecurity.stepId(), packageSecurity.security());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookProtectionResult protection ->
          new WorkbookInspectionResult.WorkbookProtectionResult(
              protection.stepId(), protection.protection());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlMappingsResult customXmlMappings ->
          new WorkbookInspectionResult.CustomXmlMappingsResult(
              customXmlMappings.stepId(), customXmlMappings.mappings());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlExportResult customXmlExport ->
          new WorkbookInspectionResult.CustomXmlExportResult(
              customXmlExport.stepId(), customXmlExport.export());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.NamedRangesResult namedRanges ->
          new WorkbookInspectionResult.NamedRangesResult(
              namedRanges.stepId(), namedRanges.namedRanges());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult sheetSummary ->
          new SheetInspectionResult.SheetSummaryResult(sheetSummary.stepId(), sheetSummary.sheet());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.ArrayFormulasResult arrayFormulas ->
          new SheetInspectionResult.ArrayFormulasResult(
              arrayFormulas.stepId(), arrayFormulas.arrayFormulas());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.CellsResult cells ->
          new SheetInspectionResult.CellsResult(
              cells.stepId(), new SheetLabel(cells.sheetName()), cells.cells());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult window ->
          new SheetInspectionResult.WindowResult(
              window.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toWindowReport(
                  window.window(), window.projection(), window.includeBlanks(), window.date1904()));
      case dev.erst.gridgrind.excel.WorkbookSheetResult.MergedRegionsResult mergedRegions ->
          new SheetInspectionResult.MergedRegionsResult(
              mergedRegions.stepId(),
              new SheetLabel(mergedRegions.sheetName()),
              mergedRegions.mergedRegions());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.HyperlinksResult hyperlinks ->
          new SheetInspectionResult.HyperlinksResult(
              hyperlinks.stepId(), new SheetLabel(hyperlinks.sheetName()), hyperlinks.hyperlinks());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.CommentsResult comments ->
          new SheetInspectionResult.CommentsResult(
              comments.stepId(), new SheetLabel(comments.sheetName()), comments.comments());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectsResult drawingObjects ->
          new WorkbookAssetInspectionResult.DrawingObjectsResult(
              drawingObjects.stepId(),
              new SheetLabel(drawingObjects.sheetName()),
              drawingObjects.drawingObjects());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult charts ->
          new WorkbookAssetInspectionResult.ChartsResult(
              charts.stepId(), new SheetLabel(charts.sheetName()), charts.charts());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult pivotTables ->
          new WorkbookAssetInspectionResult.PivotTablesResult(
              pivotTables.stepId(), pivotTables.pivotTables());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectPayloadResult
              drawingPayload ->
          new WorkbookAssetInspectionResult.DrawingObjectPayloadResult(
              drawingPayload.stepId(),
              new SheetLabel(drawingPayload.sheetName()),
              drawingPayload.payload());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult sheetLayout ->
          new SheetInspectionResult.SheetLayoutResult(sheetLayout.stepId(), sheetLayout.layout());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult printLayout ->
          new SheetInspectionResult.PrintLayoutResult(
              printLayout.stepId(), printLayout.printLayout());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.DataValidationsResult dataValidations ->
          new SheetInspectionResult.DataValidationsResult(
              dataValidations.stepId(),
              new SheetLabel(dataValidations.sheetName()),
              dataValidations.validations());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult
              conditionalFormatting ->
          new SheetInspectionResult.ConditionalFormattingResult(
              conditionalFormatting.stepId(),
              new SheetLabel(conditionalFormatting.sheetName()),
              conditionalFormatting.conditionalFormattingBlocks());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult autofilters ->
          new SheetInspectionResult.AutofiltersResult(
              autofilters.stepId(),
              new SheetLabel(autofilters.sheetName()),
              autofilters.autofilters());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult tables ->
          new WorkbookAssetInspectionResult.TablesResult(tables.stepId(), tables.tables());
      case dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurfaceResult formulaSurface ->
          new WorkbookSurfaceInspectionResult.FormulaSurfaceResult(
              formulaSurface.stepId(),
              InspectionResultSurfaceReportSupport.toFormulaSurfaceReport(
                  formulaSurface.surface()));
      case dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult sheetSchema ->
          new WorkbookSurfaceInspectionResult.SheetSchemaResult(
              sheetSchema.stepId(),
              InspectionResultSurfaceReportSupport.toSheetSchemaReport(
                  sheetSchema.surface(), sheetSchema.projection(), sheetSchema.date1904()));
      case dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceResult
              namedRangeSurface ->
          new WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult(
              namedRangeSurface.stepId(),
              InspectionResultSurfaceReportSupport.toNamedRangeSurfaceReport(
                  namedRangeSurface.surface()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.FormulaHealthResult formulaHealth ->
          new WorkbookAnalysisResult.FormulaHealthResult(
              formulaHealth.stepId(), formulaHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.DataValidationHealthResult
              dataValidationHealth ->
          new WorkbookAnalysisResult.DataValidationHealthResult(
              dataValidationHealth.stepId(), dataValidationHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.ConditionalFormattingHealthResult
              conditionalFormattingHealth ->
          new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
              conditionalFormattingHealth.stepId(), conditionalFormattingHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.AutofilterHealthResult
              autofilterHealth ->
          new WorkbookAnalysisResult.AutofilterHealthResult(
              autofilterHealth.stepId(), autofilterHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.TableHealthResult tableHealth ->
          new WorkbookAnalysisResult.TableHealthResult(
              tableHealth.stepId(), tableHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.PivotTableHealthResult
              pivotTableHealth ->
          new WorkbookAnalysisResult.PivotTableHealthResult(
              pivotTableHealth.stepId(), pivotTableHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.HyperlinkHealthResult hyperlinkHealth ->
          new WorkbookAnalysisResult.HyperlinkHealthResult(
              hyperlinkHealth.stepId(), hyperlinkHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.NamedRangeHealthResult
              namedRangeHealth ->
          new WorkbookAnalysisResult.NamedRangeHealthResult(
              namedRangeHealth.stepId(), namedRangeHealth.analysis());
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.WorkbookFindingsResult
              workbookFindings ->
          new WorkbookAnalysisResult.WorkbookFindingsResult(
              workbookFindings.stepId(), workbookFindings.analysis());
    };
  }
}
