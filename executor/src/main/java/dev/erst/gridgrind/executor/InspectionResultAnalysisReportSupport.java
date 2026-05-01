package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports;
import dev.erst.gridgrind.contract.dto.GridGrindSchemaAndFormulaReports;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.excel.WorkbookAnalysis;

/** Converts workbook-analysis snapshots into protocol health and findings reports. */
final class InspectionResultAnalysisReportSupport {
  private InspectionResultAnalysisReportSupport() {}

  static GridGrindSchemaAndFormulaReports.FormulaSurfaceReport toFormulaSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurface analysis) {
    return new GridGrindSchemaAndFormulaReports.FormulaSurfaceReport(
        analysis.totalFormulaCellCount(),
        analysis.sheets().stream()
            .map(
                sheet ->
                    new GridGrindSchemaAndFormulaReports.SheetFormulaSurfaceReport(
                        sheet.sheetName(),
                        sheet.formulaCellCount(),
                        sheet.distinctFormulaCount(),
                        sheet.formulas().stream()
                            .map(
                                formula ->
                                    new GridGrindSchemaAndFormulaReports.FormulaPatternReport(
                                        formula.formula(),
                                        formula.occurrenceCount(),
                                        formula.addresses()))
                            .toList()))
            .toList());
  }

  static GridGrindSchemaAndFormulaReports.SheetSchemaReport toSheetSchemaReport(
      dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchema analysis) {
    return new GridGrindSchemaAndFormulaReports.SheetSchemaReport(
        analysis.sheetName(),
        analysis.topLeftAddress(),
        analysis.rowCount(),
        analysis.columnCount(),
        analysis.dataRowCount(),
        analysis.columns().stream()
            .map(
                column ->
                    new GridGrindSchemaAndFormulaReports.SchemaColumnReport(
                        column.columnIndex(),
                        column.columnAddress(),
                        column.headerDisplayValue(),
                        column.populatedCellCount(),
                        column.blankCellCount(),
                        column.observedTypes().stream()
                            .map(
                                typeCount ->
                                    new GridGrindSchemaAndFormulaReports.TypeCountReport(
                                        typeCount.type(), typeCount.count()))
                            .toList(),
                        column.dominantType()))
            .toList());
  }

  static GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport toNamedRangeSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurface analysis) {
    return new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport(
        analysis.workbookScopedCount(),
        analysis.sheetScopedCount(),
        analysis.rangeBackedCount(),
        analysis.formulaBackedCount(),
        analysis.namedRanges().stream()
            .map(
                entry ->
                    new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceEntryReport(
                        entry.name(),
                        InspectionResultWorkbookCoreReportSupport.toNamedRangeScope(entry.scope()),
                        entry.refersToFormula(),
                        switch (entry.kind()) {
                          case RANGE ->
                              GridGrindSchemaAndFormulaReports.NamedRangeBackingKind.RANGE;
                          case FORMULA ->
                              GridGrindSchemaAndFormulaReports.NamedRangeBackingKind.FORMULA;
                        }))
            .toList());
  }

  static GridGrindAnalysisReports.FormulaHealthReport toFormulaHealthReport(
      WorkbookAnalysis.FormulaHealth analysis) {
    return new GridGrindAnalysisReports.FormulaHealthReport(
        analysis.checkedFormulaCellCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static DataValidationHealthReport toDataValidationHealthReport(
      WorkbookAnalysis.DataValidationHealth analysis) {
    return new DataValidationHealthReport(
        analysis.checkedValidationCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static ConditionalFormattingHealthReport toConditionalFormattingHealthReport(
      WorkbookAnalysis.ConditionalFormattingHealth analysis) {
    return new ConditionalFormattingHealthReport(
        analysis.checkedConditionalFormattingBlockCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static AutofilterHealthReport toAutofilterHealthReport(
      WorkbookAnalysis.AutofilterHealth analysis) {
    return new AutofilterHealthReport(
        analysis.checkedAutofilterCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static TableHealthReport toTableHealthReport(WorkbookAnalysis.TableHealth analysis) {
    return new TableHealthReport(
        analysis.checkedTableCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static PivotTableHealthReport toPivotTableHealthReport(
      WorkbookAnalysis.PivotTableHealth analysis) {
    return new PivotTableHealthReport(
        analysis.checkedPivotTableCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static GridGrindAnalysisReports.HyperlinkHealthReport toHyperlinkHealthReport(
      WorkbookAnalysis.HyperlinkHealth analysis) {
    return new GridGrindAnalysisReports.HyperlinkHealthReport(
        analysis.checkedHyperlinkCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static GridGrindAnalysisReports.NamedRangeHealthReport toNamedRangeHealthReport(
      WorkbookAnalysis.NamedRangeHealth analysis) {
    return new GridGrindAnalysisReports.NamedRangeHealthReport(
        analysis.checkedNamedRangeCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static GridGrindAnalysisReports.WorkbookFindingsReport toWorkbookFindingsReport(
      WorkbookAnalysis.WorkbookFindings analysis) {
    return new GridGrindAnalysisReports.WorkbookFindingsReport(
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindAnalysisReports.AnalysisSummaryReport toAnalysisSummaryReport(
      WorkbookAnalysis.AnalysisSummary summary) {
    return new GridGrindAnalysisReports.AnalysisSummaryReport(
        summary.totalCount(), summary.errorCount(), summary.warningCount(), summary.infoCount());
  }

  private static GridGrindAnalysisReports.AnalysisFindingReport toAnalysisFindingReport(
      WorkbookAnalysis.AnalysisFinding finding) {
    return new GridGrindAnalysisReports.AnalysisFindingReport(
        finding.code(),
        finding.severity(),
        finding.title(),
        finding.message(),
        toAnalysisLocationReport(finding.location()),
        finding.evidence());
  }

  private static GridGrindAnalysisReports.AnalysisLocationReport toAnalysisLocationReport(
      WorkbookAnalysis.AnalysisLocation location) {
    return switch (location) {
      case WorkbookAnalysis.AnalysisLocation.Workbook _ ->
          new GridGrindAnalysisReports.AnalysisLocationReport.Workbook();
      case WorkbookAnalysis.AnalysisLocation.Sheet sheet ->
          new GridGrindAnalysisReports.AnalysisLocationReport.Sheet(sheet.sheetName());
      case WorkbookAnalysis.AnalysisLocation.Cell cell ->
          new GridGrindAnalysisReports.AnalysisLocationReport.Cell(
              cell.sheetName(), cell.address());
      case WorkbookAnalysis.AnalysisLocation.Range range ->
          new GridGrindAnalysisReports.AnalysisLocationReport.Range(
              range.sheetName(), range.range());
      case WorkbookAnalysis.AnalysisLocation.NamedRange namedRange ->
          new GridGrindAnalysisReports.AnalysisLocationReport.NamedRange(
              namedRange.name(),
              InspectionResultWorkbookCoreReportSupport.toNamedRangeScope(namedRange.scope()));
    };
  }
}
