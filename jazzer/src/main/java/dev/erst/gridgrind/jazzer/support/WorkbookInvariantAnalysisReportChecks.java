package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.AnalysisFindingReport;
import dev.erst.gridgrind.contract.dto.AnalysisLocationReport;
import dev.erst.gridgrind.contract.dto.AnalysisSummaryReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.FormulaHealthReport;
import dev.erst.gridgrind.contract.dto.FormulaSurfaceReport;
import dev.erst.gridgrind.contract.dto.HyperlinkHealthReport;
import dev.erst.gridgrind.contract.dto.NamedRangeHealthReport;
import dev.erst.gridgrind.contract.dto.NamedRangeSurfaceReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.SheetSchemaReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.WorkbookFindingsReport;
import java.util.List;

/** Owns invariant checks for analysis summaries, findings, and report-style analysis payloads. */
final class WorkbookInvariantAnalysisReportChecks {
  private WorkbookInvariantAnalysisReportChecks() {}

  static void requireFormulaSurfaceShape(FormulaSurfaceReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.totalFormulaCellCount() >= 0, "totalFormulaCellCount must not be negative");
    analysis
        .sheets()
        .forEach(
            sheet -> {
              WorkbookInvariantChecks.require(
                  sheet.sheetName() != null, "formula surface sheetName must not be null");
              WorkbookInvariantChecks.require(
                  !sheet.sheetName().isBlank(), "formula surface sheetName must not be blank");
              WorkbookInvariantChecks.require(
                  sheet.formulaCellCount() >= 0, "formulaCellCount must not be negative");
              WorkbookInvariantChecks.require(
                  sheet.distinctFormulaCount() >= 0, "distinctFormulaCount must not be negative");
              sheet
                  .formulas()
                  .forEach(
                      formula -> {
                        WorkbookInvariantChecks.require(
                            formula.formula() != null, "formula pattern must not be null");
                        WorkbookInvariantChecks.require(
                            !formula.formula().isBlank(), "formula pattern must not be blank");
                        WorkbookInvariantChecks.require(
                            formula.occurrenceCount() > 0,
                            "occurrenceCount must be greater than 0");
                        WorkbookInvariantChecks.require(
                            formula.addresses() != null, "formula addresses must not be null");
                      });
            });
  }

  static void requireSheetSchemaShape(SheetSchemaReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.sheetName() != null, "schema sheetName must not be null");
    WorkbookInvariantChecks.require(
        !analysis.sheetName().isBlank(), "schema sheetName must not be blank");
    WorkbookInvariantChecks.require(
        analysis.topLeftAddress() != null, "schema topLeftAddress must not be null");
    WorkbookInvariantChecks.require(
        !analysis.topLeftAddress().isBlank(), "schema topLeftAddress must not be blank");
    WorkbookInvariantChecks.require(
        analysis.rowCount() > 0, "schema rowCount must be greater than 0");
    WorkbookInvariantChecks.require(
        analysis.columnCount() > 0, "schema columnCount must be greater than 0");
    WorkbookInvariantChecks.require(
        analysis.dataRowCount() >= 0, "schema dataRowCount must not be negative");
    analysis
        .columns()
        .forEach(
            column -> {
              WorkbookInvariantChecks.require(
                  column.columnIndex() >= 0, "schema columnIndex must not be negative");
              WorkbookInvariantChecks.require(
                  column.columnAddress() != null, "schema columnAddress must not be null");
              WorkbookInvariantChecks.require(
                  !column.columnAddress().isBlank(), "schema columnAddress must not be blank");
              WorkbookInvariantChecks.require(
                  column.headerDisplayValue() != null,
                  "schema headerDisplayValue must not be null");
              WorkbookInvariantChecks.require(
                  column.populatedCellCount() >= 0,
                  "schema populatedCellCount must not be negative");
              WorkbookInvariantChecks.require(
                  column.blankCellCount() >= 0, "schema blankCellCount must not be negative");
              column
                  .observedTypes()
                  .forEach(
                      typeCount -> {
                        WorkbookInvariantChecks.require(
                            typeCount.type() != null, "type count type must not be null");
                        WorkbookInvariantChecks.require(
                            !typeCount.type().isBlank(), "type count type must not be blank");
                        WorkbookInvariantChecks.require(
                            typeCount.count() > 0, "type count must be greater than 0");
                      });
            });
  }

  static void requireNamedRangeSurfaceShape(NamedRangeSurfaceReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.workbookScopedCount() >= 0, "workbookScopedCount must not be negative");
    WorkbookInvariantChecks.require(
        analysis.sheetScopedCount() >= 0, "sheetScopedCount must not be negative");
    WorkbookInvariantChecks.require(
        analysis.rangeBackedCount() >= 0, "rangeBackedCount must not be negative");
    WorkbookInvariantChecks.require(
        analysis.formulaBackedCount() >= 0, "formulaBackedCount must not be negative");
    analysis
        .namedRanges()
        .forEach(
            namedRange -> {
              WorkbookInvariantChecks.require(
                  namedRange.name() != null, "named range name must not be null");
              WorkbookInvariantChecks.require(
                  !namedRange.name().isBlank(), "named range name must not be blank");
              WorkbookInvariantChecks.require(
                  namedRange.scope() != null, "named range scope must not be null");
              WorkbookInvariantChecks.require(
                  namedRange.refersToFormula() != null,
                  "named range refersToFormula must not be null");
              WorkbookInvariantChecks.require(
                  namedRange.kind() != null, "named range kind must not be null");
            });
  }

  static void requireFormulaHealthShape(FormulaHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedFormulaCellCount() >= 0, "checkedFormulaCellCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireDataValidationHealthShape(
      dev.erst.gridgrind.contract.dto.DataValidationHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedValidationCount() >= 0, "checkedValidationCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireConditionalFormattingHealthShape(ConditionalFormattingHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedConditionalFormattingBlockCount() >= 0,
        "checkedConditionalFormattingBlockCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireAutofilterHealthShape(AutofilterHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedAutofilterCount() >= 0, "checkedAutofilterCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireTableHealthShape(TableHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedTableCount() >= 0, "checkedTableCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requirePivotTableHealthShape(PivotTableHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedPivotTableCount() >= 0, "checkedPivotTableCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireHyperlinkHealthShape(HyperlinkHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedHyperlinkCount() >= 0, "checkedHyperlinkCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireNamedRangeHealthShape(NamedRangeHealthReport analysis) {
    WorkbookInvariantChecks.require(
        analysis.checkedNamedRangeCount() >= 0, "checkedNamedRangeCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireWorkbookFindingsShape(WorkbookFindingsReport analysis) {
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  static void requireAnalysisSummaryShape(
      AnalysisSummaryReport summary, List<AnalysisFindingReport> findings) {
    WorkbookInvariantChecks.require(summary != null, "analysis summary must not be null");
    WorkbookInvariantChecks.require(findings != null, "analysis findings must not be null");
    WorkbookInvariantChecks.require(
        summary.totalCount() >= 0, "analysis totalCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.errorCount() >= 0, "analysis errorCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.warningCount() >= 0, "analysis warningCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.infoCount() >= 0, "analysis infoCount must not be negative");
    WorkbookInvariantChecks.require(
        summary.totalCount() == findings.size(), "analysis totalCount must match findings size");
    WorkbookInvariantChecks.require(
        summary.totalCount() == summary.errorCount() + summary.warningCount() + summary.infoCount(),
        "analysis totalCount must equal error + warning + info");
    findings.forEach(WorkbookInvariantAnalysisReportChecks::requireAnalysisFindingShape);
  }

  static void requireAnalysisFindingShape(AnalysisFindingReport finding) {
    WorkbookInvariantChecks.require(
        finding.code() != null, "analysis finding code must not be null");
    WorkbookInvariantChecks.require(
        finding.severity() != null, "analysis finding severity must not be null");
    WorkbookInvariantChecks.requireNonBlank(finding.title(), "analysis title");
    WorkbookInvariantChecks.requireNonBlank(finding.message(), "analysis message");
    WorkbookInvariantChecks.require(
        finding.location() != null, "analysis location must not be null");
    WorkbookInvariantChecks.require(
        finding.evidence() != null, "analysis evidence must not be null");
    finding
        .evidence()
        .forEach(
            evidence -> WorkbookInvariantChecks.requireNonBlank(evidence, "analysis evidence"));

    switch (finding.location()) {
      case AnalysisLocationReport.Workbook _ -> {}
      case AnalysisLocationReport.Sheet sheet ->
          WorkbookInvariantChecks.requireNonBlank(sheet.sheetName(), "analysis sheetName");
      case AnalysisLocationReport.Cell cell -> {
        WorkbookInvariantChecks.requireNonBlank(cell.sheetName(), "analysis sheetName");
        WorkbookInvariantChecks.requireNonBlank(cell.address(), "analysis address");
      }
      case AnalysisLocationReport.Range range -> {
        WorkbookInvariantChecks.requireNonBlank(range.sheetName(), "analysis sheetName");
        WorkbookInvariantChecks.requireNonBlank(range.range(), "analysis range");
      }
      case AnalysisLocationReport.NamedRange namedRange -> {
        WorkbookInvariantChecks.requireNonBlank(namedRange.name(), "analysis named range");
        WorkbookInvariantChecks.require(
            namedRange.scope() != null, "analysis named range scope must not be null");
      }
    }
  }
}
