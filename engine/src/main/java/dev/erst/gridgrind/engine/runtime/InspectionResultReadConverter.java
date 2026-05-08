package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.MergedRegionReport;
import dev.erst.gridgrind.contract.dto.SheetSummaryReport;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookSurfaceInspectionResult;

/** Converts engine read results into protocol inspection-result variants. */
final class InspectionResultReadConverter {
  private InspectionResultReadConverter() {}

  static InspectionResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult workbookSummary ->
          new WorkbookInspectionResult.WorkbookSummaryResult(
              workbookSummary.stepId(),
              InspectionResultWorkbookCoreReportSupport.toWorkbookSummary(
                  workbookSummary.workbook()));
      case dev.erst.gridgrind.excel.WorkbookCoreResult.PackageSecurityResult packageSecurity ->
          new WorkbookInspectionResult.PackageSecurityResult(
              packageSecurity.stepId(),
              InspectionResultWorkbookCoreReportSupport.toOoxmlPackageSecurityReport(
                  packageSecurity.security()));
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookProtectionResult protection ->
          new WorkbookInspectionResult.WorkbookProtectionResult(
              protection.stepId(),
              InspectionResultWorkbookCoreReportSupport.toWorkbookProtectionReport(
                  protection.protection()));
      case dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlMappingsResult customXmlMappings ->
          new WorkbookInspectionResult.CustomXmlMappingsResult(
              customXmlMappings.stepId(),
              customXmlMappings.mappings().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlMappingReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.CustomXmlExportResult customXmlExport ->
          new WorkbookInspectionResult.CustomXmlExportResult(
              customXmlExport.stepId(),
              InspectionResultWorkbookStructureReportSupport.toCustomXmlExportReport(
                  customXmlExport.export()));
      case dev.erst.gridgrind.excel.WorkbookCoreResult.NamedRangesResult namedRanges ->
          new WorkbookInspectionResult.NamedRangesResult(
              namedRanges.stepId(),
              namedRanges.namedRanges().stream()
                  .map(InspectionResultWorkbookCoreReportSupport::toNamedRangeReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult sheetSummary ->
          new SheetInspectionResult.SheetSummaryResult(
              sheetSummary.stepId(),
              new SheetSummaryReport(
                  sheetSummary.sheet().sheetName(),
                  sheetSummary.sheet().visibility(),
                  InspectionResultWorkbookLayoutReportSupport.toSheetProtectionReport(
                      sheetSummary.sheet().protection()),
                  sheetSummary.sheet().physicalRowCount(),
                  sheetSummary.sheet().lastRowIndex(),
                  sheetSummary.sheet().lastColumnIndex()));
      case dev.erst.gridgrind.excel.WorkbookSheetResult.ArrayFormulasResult arrayFormulas ->
          new SheetInspectionResult.ArrayFormulasResult(
              arrayFormulas.stepId(),
              arrayFormulas.arrayFormulas().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toArrayFormulaReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.CellsResult cells ->
          new SheetInspectionResult.CellsResult(
              cells.stepId(),
              cells.sheetName(),
              cells.cells().stream().map(InspectionResultCellReportSupport::toCellReport).toList());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult window ->
          new SheetInspectionResult.WindowResult(
              window.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toWindowReport(window.window()));
      case dev.erst.gridgrind.excel.WorkbookSheetResult.MergedRegionsResult mergedRegions ->
          new SheetInspectionResult.MergedRegionsResult(
              mergedRegions.stepId(),
              mergedRegions.sheetName(),
              mergedRegions.mergedRegions().stream()
                  .map(region -> new MergedRegionReport(region.range()))
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.HyperlinksResult hyperlinks ->
          new SheetInspectionResult.HyperlinksResult(
              hyperlinks.stepId(),
              hyperlinks.sheetName(),
              hyperlinks.hyperlinks().stream()
                  .map(InspectionResultWorkbookLayoutReportSupport::toCellHyperlinkReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookSheetResult.CommentsResult comments ->
          new SheetInspectionResult.CommentsResult(
              comments.stepId(),
              comments.sheetName(),
              comments.comments().stream()
                  .map(InspectionResultWorkbookLayoutReportSupport::toCellCommentReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectsResult drawingObjects ->
          new WorkbookAssetInspectionResult.DrawingObjectsResult(
              drawingObjects.stepId(),
              drawingObjects.sheetName(),
              drawingObjects.drawingObjects().stream()
                  .map(InspectionResultDrawingReportSupport::toDrawingObjectReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult charts ->
          new WorkbookAssetInspectionResult.ChartsResult(
              charts.stepId(),
              charts.sheetName(),
              charts.charts().stream()
                  .map(InspectionResultDrawingReportSupport::toChartReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult pivotTables ->
          new WorkbookAssetInspectionResult.PivotTablesResult(
              pivotTables.stepId(),
              pivotTables.pivotTables().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toPivotTableReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectPayloadResult
              drawingPayload ->
          new WorkbookAssetInspectionResult.DrawingObjectPayloadResult(
              drawingPayload.stepId(),
              drawingPayload.sheetName(),
              InspectionResultDrawingReportSupport.toDrawingObjectPayloadReport(
                  drawingPayload.payload()));
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult sheetLayout ->
          new SheetInspectionResult.SheetLayoutResult(
              sheetLayout.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toSheetLayoutReport(
                  sheetLayout.layout()));
      case dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult printLayout ->
          new SheetInspectionResult.PrintLayoutResult(
              printLayout.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toPrintLayoutReport(printLayout));
      case dev.erst.gridgrind.excel.WorkbookRuleResult.DataValidationsResult dataValidations ->
          new SheetInspectionResult.DataValidationsResult(
              dataValidations.stepId(),
              dataValidations.sheetName(),
              dataValidations.validations().stream()
                  .map(InspectionResultValidationReportSupport::toDataValidationEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult
              conditionalFormatting ->
          new SheetInspectionResult.ConditionalFormattingResult(
              conditionalFormatting.stepId(),
              conditionalFormatting.sheetName(),
              conditionalFormatting.conditionalFormattingBlocks().stream()
                  .map(InspectionResultValidationReportSupport::toConditionalFormattingEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.AutofiltersResult autofilters ->
          new SheetInspectionResult.AutofiltersResult(
              autofilters.stepId(),
              autofilters.sheetName(),
              autofilters.autofilters().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toAutofilterEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookRuleResult.TablesResult tables ->
          new WorkbookAssetInspectionResult.TablesResult(
              tables.stepId(),
              tables.tables().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toTableEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurfaceResult formulaSurface ->
          new WorkbookSurfaceInspectionResult.FormulaSurfaceResult(
              formulaSurface.stepId(),
              InspectionResultSurfaceReportSupport.toFormulaSurfaceReport(
                  formulaSurface.surface()));
      case dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult sheetSchema ->
          new WorkbookSurfaceInspectionResult.SheetSchemaResult(
              sheetSchema.stepId(),
              InspectionResultSurfaceReportSupport.toSheetSchemaReport(sheetSchema.surface()));
      case dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceResult
              namedRangeSurface ->
          new WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult(
              namedRangeSurface.stepId(),
              InspectionResultSurfaceReportSupport.toNamedRangeSurfaceReport(
                  namedRangeSurface.surface()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.FormulaHealthResult formulaHealth ->
          new WorkbookAnalysisResult.FormulaHealthResult(
              formulaHealth.stepId(),
              InspectionResultAnalysisReportSupport.toFormulaHealthReport(
                  formulaHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.DataValidationHealthResult
              dataValidationHealth ->
          new WorkbookAnalysisResult.DataValidationHealthResult(
              dataValidationHealth.stepId(),
              InspectionResultAnalysisReportSupport.toDataValidationHealthReport(
                  dataValidationHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.ConditionalFormattingHealthResult
              conditionalFormattingHealth ->
          new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
              conditionalFormattingHealth.stepId(),
              InspectionResultAnalysisReportSupport.toConditionalFormattingHealthReport(
                  conditionalFormattingHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.AutofilterHealthResult
              autofilterHealth ->
          new WorkbookAnalysisResult.AutofilterHealthResult(
              autofilterHealth.stepId(),
              InspectionResultAnalysisReportSupport.toAutofilterHealthReport(
                  autofilterHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.TableHealthResult tableHealth ->
          new WorkbookAnalysisResult.TableHealthResult(
              tableHealth.stepId(),
              InspectionResultAnalysisReportSupport.toTableHealthReport(tableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.PivotTableHealthResult
              pivotTableHealth ->
          new WorkbookAnalysisResult.PivotTableHealthResult(
              pivotTableHealth.stepId(),
              InspectionResultAnalysisReportSupport.toPivotTableHealthReport(
                  pivotTableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.HyperlinkHealthResult hyperlinkHealth ->
          new WorkbookAnalysisResult.HyperlinkHealthResult(
              hyperlinkHealth.stepId(),
              InspectionResultAnalysisReportSupport.toHyperlinkHealthReport(
                  hyperlinkHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.NamedRangeHealthResult
              namedRangeHealth ->
          new WorkbookAnalysisResult.NamedRangeHealthResult(
              namedRangeHealth.stepId(),
              InspectionResultAnalysisReportSupport.toNamedRangeHealthReport(
                  namedRangeHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookAnalysisResult.WorkbookFindingsResult
              workbookFindings ->
          new WorkbookAnalysisResult.WorkbookFindingsResult(
              workbookFindings.stepId(),
              InspectionResultAnalysisReportSupport.toWorkbookFindingsReport(
                  workbookFindings.analysis()));
    };
  }
}
