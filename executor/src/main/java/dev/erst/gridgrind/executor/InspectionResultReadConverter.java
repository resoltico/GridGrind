package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.query.InspectionResult;

/** Converts engine read results into protocol inspection-result variants. */
final class InspectionResultReadConverter {
  private InspectionResultReadConverter() {}

  static InspectionResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult workbookSummary ->
          new InspectionResult.WorkbookSummaryResult(
              workbookSummary.stepId(),
              InspectionResultWorkbookCoreReportSupport.toWorkbookSummary(
                  workbookSummary.workbook()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PackageSecurityResult packageSecurity ->
          new InspectionResult.PackageSecurityResult(
              packageSecurity.stepId(),
              InspectionResultWorkbookCoreReportSupport.toOoxmlPackageSecurityReport(
                  packageSecurity.security()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookProtectionResult protection ->
          new InspectionResult.WorkbookProtectionResult(
              protection.stepId(),
              InspectionResultWorkbookCoreReportSupport.toWorkbookProtectionReport(
                  protection.protection()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.CustomXmlMappingsResult customXmlMappings ->
          new InspectionResult.CustomXmlMappingsResult(
              customXmlMappings.stepId(),
              customXmlMappings.mappings().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toCustomXmlMappingReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CustomXmlExportResult customXmlExport ->
          new InspectionResult.CustomXmlExportResult(
              customXmlExport.stepId(),
              InspectionResultWorkbookStructureReportSupport.toCustomXmlExportReport(
                  customXmlExport.export()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult namedRanges ->
          new InspectionResult.NamedRangesResult(
              namedRanges.stepId(),
              namedRanges.namedRanges().stream()
                  .map(InspectionResultWorkbookCoreReportSupport::toNamedRangeReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult sheetSummary ->
          new InspectionResult.SheetSummaryResult(
              sheetSummary.stepId(),
              new GridGrindResponse.SheetSummaryReport(
                  sheetSummary.sheet().sheetName(),
                  sheetSummary.sheet().visibility(),
                  InspectionResultWorkbookLayoutReportSupport.toSheetProtectionReport(
                      sheetSummary.sheet().protection()),
                  sheetSummary.sheet().physicalRowCount(),
                  sheetSummary.sheet().lastRowIndex(),
                  sheetSummary.sheet().lastColumnIndex()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.ArrayFormulasResult arrayFormulas ->
          new InspectionResult.ArrayFormulasResult(
              arrayFormulas.stepId(),
              arrayFormulas.arrayFormulas().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toArrayFormulaReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult cells ->
          new InspectionResult.CellsResult(
              cells.stepId(),
              cells.sheetName(),
              cells.cells().stream().map(InspectionResultCellReportSupport::toCellReport).toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult window ->
          new InspectionResult.WindowResult(
              window.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toWindowReport(window.window()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult mergedRegions ->
          new InspectionResult.MergedRegionsResult(
              mergedRegions.stepId(),
              mergedRegions.sheetName(),
              mergedRegions.mergedRegions().stream()
                  .map(region -> new GridGrindResponse.MergedRegionReport(region.range()))
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult hyperlinks ->
          new InspectionResult.HyperlinksResult(
              hyperlinks.stepId(),
              hyperlinks.sheetName(),
              hyperlinks.hyperlinks().stream()
                  .map(InspectionResultWorkbookLayoutReportSupport::toCellHyperlinkReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult comments ->
          new InspectionResult.CommentsResult(
              comments.stepId(),
              comments.sheetName(),
              comments.comments().stream()
                  .map(InspectionResultWorkbookLayoutReportSupport::toCellCommentReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectsResult drawingObjects ->
          new InspectionResult.DrawingObjectsResult(
              drawingObjects.stepId(),
              drawingObjects.sheetName(),
              drawingObjects.drawingObjects().stream()
                  .map(InspectionResultDrawingReportSupport::toDrawingObjectReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult charts ->
          new InspectionResult.ChartsResult(
              charts.stepId(),
              charts.sheetName(),
              charts.charts().stream()
                  .map(InspectionResultDrawingReportSupport::toChartReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult pivotTables ->
          new InspectionResult.PivotTablesResult(
              pivotTables.stepId(),
              pivotTables.pivotTables().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toPivotTableReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectPayloadResult drawingPayload ->
          new InspectionResult.DrawingObjectPayloadResult(
              drawingPayload.stepId(),
              drawingPayload.sheetName(),
              InspectionResultDrawingReportSupport.toDrawingObjectPayloadReport(
                  drawingPayload.payload()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult sheetLayout ->
          new InspectionResult.SheetLayoutResult(
              sheetLayout.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toSheetLayoutReport(
                  sheetLayout.layout()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult printLayout ->
          new InspectionResult.PrintLayoutResult(
              printLayout.stepId(),
              InspectionResultWorkbookLayoutReportSupport.toPrintLayoutReport(printLayout));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationsResult dataValidations ->
          new InspectionResult.DataValidationsResult(
              dataValidations.stepId(),
              dataValidations.sheetName(),
              dataValidations.validations().stream()
                  .map(InspectionResultValidationReportSupport::toDataValidationEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult
              conditionalFormatting ->
          new InspectionResult.ConditionalFormattingResult(
              conditionalFormatting.stepId(),
              conditionalFormatting.sheetName(),
              conditionalFormatting.conditionalFormattingBlocks().stream()
                  .map(InspectionResultValidationReportSupport::toConditionalFormattingEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult autofilters ->
          new InspectionResult.AutofiltersResult(
              autofilters.stepId(),
              autofilters.sheetName(),
              autofilters.autofilters().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toAutofilterEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult tables ->
          new InspectionResult.TablesResult(
              tables.stepId(),
              tables.tables().stream()
                  .map(InspectionResultWorkbookStructureReportSupport::toTableEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult formulaSurface ->
          new InspectionResult.FormulaSurfaceResult(
              formulaSurface.stepId(),
              InspectionResultAnalysisReportSupport.toFormulaSurfaceReport(
                  formulaSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult sheetSchema ->
          new InspectionResult.SheetSchemaResult(
              sheetSchema.stepId(),
              InspectionResultAnalysisReportSupport.toSheetSchemaReport(sheetSchema.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface ->
          new InspectionResult.NamedRangeSurfaceResult(
              namedRangeSurface.stepId(),
              InspectionResultAnalysisReportSupport.toNamedRangeSurfaceReport(
                  namedRangeSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult formulaHealth ->
          new InspectionResult.FormulaHealthResult(
              formulaHealth.stepId(),
              InspectionResultAnalysisReportSupport.toFormulaHealthReport(
                  formulaHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationHealthResult
              dataValidationHealth ->
          new InspectionResult.DataValidationHealthResult(
              dataValidationHealth.stepId(),
              InspectionResultAnalysisReportSupport.toDataValidationHealthReport(
                  dataValidationHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult
              conditionalFormattingHealth ->
          new InspectionResult.ConditionalFormattingHealthResult(
              conditionalFormattingHealth.stepId(),
              InspectionResultAnalysisReportSupport.toConditionalFormattingHealthReport(
                  conditionalFormattingHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofilterHealthResult autofilterHealth ->
          new InspectionResult.AutofilterHealthResult(
              autofilterHealth.stepId(),
              InspectionResultAnalysisReportSupport.toAutofilterHealthReport(
                  autofilterHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.TableHealthResult tableHealth ->
          new InspectionResult.TableHealthResult(
              tableHealth.stepId(),
              InspectionResultAnalysisReportSupport.toTableHealthReport(tableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.PivotTableHealthResult pivotTableHealth ->
          new InspectionResult.PivotTableHealthResult(
              pivotTableHealth.stepId(),
              InspectionResultAnalysisReportSupport.toPivotTableHealthReport(
                  pivotTableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult hyperlinkHealth ->
          new InspectionResult.HyperlinkHealthResult(
              hyperlinkHealth.stepId(),
              InspectionResultAnalysisReportSupport.toHyperlinkHealthReport(
                  hyperlinkHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult namedRangeHealth ->
          new InspectionResult.NamedRangeHealthResult(
              namedRangeHealth.stepId(),
              InspectionResultAnalysisReportSupport.toNamedRangeHealthReport(
                  namedRangeHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult workbookFindings ->
          new InspectionResult.WorkbookFindingsResult(
              workbookFindings.stepId(),
              InspectionResultAnalysisReportSupport.toWorkbookFindingsReport(
                  workbookFindings.analysis()));
    };
  }
}
