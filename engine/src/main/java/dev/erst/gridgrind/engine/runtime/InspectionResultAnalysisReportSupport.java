package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.AnalysisFindingReport;
import dev.erst.gridgrind.contract.dto.AnalysisLocationReport;
import dev.erst.gridgrind.contract.dto.AnalysisSummaryReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.FormulaHealthReport;
import dev.erst.gridgrind.contract.dto.HyperlinkHealthReport;
import dev.erst.gridgrind.contract.dto.NamedRangeHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.WorkbookFindingsReport;
import dev.erst.gridgrind.excel.WorkbookAnalysis;

/** Converts workbook-analysis snapshots into protocol health and findings reports. */
final class InspectionResultAnalysisReportSupport {
  private InspectionResultAnalysisReportSupport() {}

  static FormulaHealthReport toFormulaHealthReport(WorkbookAnalysis.FormulaHealth analysis) {
    return new FormulaHealthReport(
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

  static HyperlinkHealthReport toHyperlinkHealthReport(WorkbookAnalysis.HyperlinkHealth analysis) {
    return new HyperlinkHealthReport(
        analysis.checkedHyperlinkCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static NamedRangeHealthReport toNamedRangeHealthReport(
      WorkbookAnalysis.NamedRangeHealth analysis) {
    return new NamedRangeHealthReport(
        analysis.checkedNamedRangeCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  static WorkbookFindingsReport toWorkbookFindingsReport(
      WorkbookAnalysis.WorkbookFindings analysis) {
    return new WorkbookFindingsReport(
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(InspectionResultAnalysisReportSupport::toAnalysisFindingReport)
            .toList());
  }

  private static AnalysisSummaryReport toAnalysisSummaryReport(
      WorkbookAnalysis.AnalysisSummary summary) {
    return new AnalysisSummaryReport(
        summary.totalCount(), summary.errorCount(), summary.warningCount(), summary.infoCount());
  }

  private static AnalysisFindingReport toAnalysisFindingReport(
      WorkbookAnalysis.AnalysisFinding finding) {
    return new AnalysisFindingReport(
        finding.code(),
        finding.severity(),
        finding.title(),
        finding.message(),
        toAnalysisLocationReport(finding.location()),
        finding.evidence());
  }

  private static AnalysisLocationReport toAnalysisLocationReport(
      WorkbookAnalysis.AnalysisLocation location) {
    return switch (location) {
      case WorkbookAnalysis.AnalysisLocation.Workbook _ -> new AnalysisLocationReport.Workbook();
      case WorkbookAnalysis.AnalysisLocation.Sheet sheet ->
          new AnalysisLocationReport.Sheet(sheet.sheetName());
      case WorkbookAnalysis.AnalysisLocation.Cell cell ->
          new AnalysisLocationReport.Cell(cell.sheetName(), cell.address());
      case WorkbookAnalysis.AnalysisLocation.Range range ->
          new AnalysisLocationReport.Range(range.sheetName(), range.range());
      case WorkbookAnalysis.AnalysisLocation.NamedRange namedRange ->
          new AnalysisLocationReport.NamedRange(
              namedRange.name(),
              InspectionResultWorkbookCoreReportSupport.toNamedRangeScope(namedRange.scope()));
    };
  }
}
