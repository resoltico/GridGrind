package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.excel.WorkbookAnalysis;

/** Converts workbook-analysis snapshots into protocol health and findings reports. */
final class InspectionResultAnalysisReportSupport {
  private InspectionResultAnalysisReportSupport() {}

  static GridGrindResponse.FormulaSurfaceReport toFormulaSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface analysis) {
    return new GridGrindResponse.FormulaSurfaceReport(
        analysis.totalFormulaCellCount(),
        analysis.sheets().stream()
            .map(
                sheet ->
                    new GridGrindResponse.SheetFormulaSurfaceReport(
                        sheet.sheetName(),
                        sheet.formulaCellCount(),
                        sheet.distinctFormulaCount(),
                        sheet.formulas().stream()
                            .map(
                                formula ->
                                    new GridGrindResponse.FormulaPatternReport(
                                        formula.formula(),
                                        formula.occurrenceCount(),
                                        formula.addresses()))
                            .toList()))
            .toList());
  }

  static GridGrindResponse.SheetSchemaReport toSheetSchemaReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema analysis) {
    return new GridGrindResponse.SheetSchemaReport(
        analysis.sheetName(),
        analysis.topLeftAddress(),
        analysis.rowCount(),
        analysis.columnCount(),
        analysis.dataRowCount(),
        analysis.columns().stream()
            .map(
                column ->
                    new GridGrindResponse.SchemaColumnReport(
                        column.columnIndex(),
                        column.columnAddress(),
                        column.headerDisplayValue(),
                        column.populatedCellCount(),
                        column.blankCellCount(),
                        column.observedTypes().stream()
                            .map(
                                typeCount ->
                                    new GridGrindResponse.TypeCountReport(
                                        typeCount.type(), typeCount.count()))
                            .toList(),
                        column.dominantType()))
            .toList());
  }

  static GridGrindResponse.NamedRangeSurfaceReport toNamedRangeSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface analysis) {
    return new GridGrindResponse.NamedRangeSurfaceReport(
        analysis.workbookScopedCount(),
        analysis.sheetScopedCount(),
        analysis.rangeBackedCount(),
        analysis.formulaBackedCount(),
        analysis.namedRanges().stream()
            .map(
                entry ->
                    new GridGrindResponse.NamedRangeSurfaceEntryReport(
                        entry.name(),
                        InspectionResultConverter.toNamedRangeScope(entry.scope()),
                        entry.refersToFormula(),
                        switch (entry.kind()) {
                          case RANGE -> GridGrindResponse.NamedRangeBackingKind.RANGE;
                          case FORMULA -> GridGrindResponse.NamedRangeBackingKind.FORMULA;
                        }))
            .toList());
  }

  static GridGrindResponse.FormulaHealthReport toFormulaHealthReport(
      WorkbookAnalysis.FormulaHealth analysis) {
    return new GridGrindResponse.FormulaHealthReport(
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

  static GridGrindResponse.HyperlinkHealthReport toHyperlinkHealthReport(
      WorkbookAnalysis.HyperlinkHealth analysis) {
    return new GridGrindResponse.HyperlinkHealthReport(
        analysis.checkedHyperlinkCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static GridGrindResponse.NamedRangeHealthReport toNamedRangeHealthReport(
      WorkbookAnalysis.NamedRangeHealth analysis) {
    return new GridGrindResponse.NamedRangeHealthReport(
        analysis.checkedNamedRangeCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static GridGrindResponse.WorkbookFindingsReport toWorkbookFindingsReport(
      WorkbookAnalysis.WorkbookFindings analysis) {
    return new GridGrindResponse.WorkbookFindingsReport(
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.AnalysisSummaryReport toAnalysisSummaryReport(
      WorkbookAnalysis.AnalysisSummary summary) {
    return new GridGrindResponse.AnalysisSummaryReport(
        summary.totalCount(), summary.errorCount(), summary.warningCount(), summary.infoCount());
  }

  private static GridGrindResponse.AnalysisFindingReport toAnalysisFindingReport(
      WorkbookAnalysis.AnalysisFinding finding) {
    return new GridGrindResponse.AnalysisFindingReport(
        toAnalysisFindingCode(finding.code()),
        toAnalysisSeverity(finding.severity()),
        finding.title(),
        finding.message(),
        toAnalysisLocationReport(finding.location()),
        finding.evidence());
  }

  private static GridGrindResponse.AnalysisLocationReport toAnalysisLocationReport(
      WorkbookAnalysis.AnalysisLocation location) {
    return switch (location) {
      case WorkbookAnalysis.AnalysisLocation.Workbook _ ->
          new GridGrindResponse.AnalysisLocationReport.Workbook();
      case WorkbookAnalysis.AnalysisLocation.Sheet sheet ->
          new GridGrindResponse.AnalysisLocationReport.Sheet(sheet.sheetName());
      case WorkbookAnalysis.AnalysisLocation.Cell cell ->
          new GridGrindResponse.AnalysisLocationReport.Cell(cell.sheetName(), cell.address());
      case WorkbookAnalysis.AnalysisLocation.Range range ->
          new GridGrindResponse.AnalysisLocationReport.Range(range.sheetName(), range.range());
      case WorkbookAnalysis.AnalysisLocation.NamedRange namedRange ->
          new GridGrindResponse.AnalysisLocationReport.NamedRange(
              namedRange.name(), InspectionResultConverter.toNamedRangeScope(namedRange.scope()));
    };
  }

  private static AnalysisFindingCode toAnalysisFindingCode(
      WorkbookAnalysis.AnalysisFindingCode code) {
    return AnalysisFindingCode.valueOf(code.name());
  }

  private static AnalysisSeverity toAnalysisSeverity(WorkbookAnalysis.AnalysisSeverity severity) {
    return AnalysisSeverity.valueOf(severity.name());
  }
}
