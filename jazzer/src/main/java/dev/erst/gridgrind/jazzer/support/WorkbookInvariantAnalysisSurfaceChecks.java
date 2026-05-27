package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.AnalysisFindingReport;
import dev.erst.gridgrind.contract.dto.AnalysisSummaryReport;
import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.CellCommentReport;
import dev.erst.gridgrind.contract.dto.CellHyperlinkReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideReport;
import dev.erst.gridgrind.contract.dto.DifferentialStyleReport;
import dev.erst.gridgrind.contract.dto.FormulaHealthReport;
import dev.erst.gridgrind.contract.dto.FormulaSurfaceReport;
import dev.erst.gridgrind.contract.dto.HyperlinkHealthReport;
import dev.erst.gridgrind.contract.dto.NamedRangeHealthReport;
import dev.erst.gridgrind.contract.dto.NamedRangeSurfaceReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.SheetLayoutReport;
import dev.erst.gridgrind.contract.dto.SheetSchemaReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WindowReport;
import dev.erst.gridgrind.contract.dto.WorkbookFindingsReport;
import java.util.List;

/**
 * Owns invariant checks for analysis payloads, structured workbook reports, and layout surfaces.
 */
final class WorkbookInvariantAnalysisSurfaceChecks {
  private WorkbookInvariantAnalysisSurfaceChecks() {}

  static void requireWindowShape(WindowReport window) {
    WorkbookInvariantAnalysisLayoutChecks.requireWindowShape(window);
  }

  static void requireHyperlinkEntryShape(CellHyperlinkReport hyperlink) {
    WorkbookInvariantAnalysisLayoutChecks.requireHyperlinkEntryShape(hyperlink);
  }

  static void requireCommentEntryShape(CellCommentReport comment) {
    WorkbookInvariantAnalysisLayoutChecks.requireCommentEntryShape(comment);
  }

  static void requireSheetLayoutShape(SheetLayoutReport layout) {
    WorkbookInvariantAnalysisLayoutChecks.requireSheetLayoutShape(layout);
  }

  static void requirePrintLayoutShape(PrintLayoutReport layout) {
    WorkbookInvariantAnalysisLayoutChecks.requirePrintLayoutShape(layout);
  }

  static void requireDataValidationEntryShape(DataValidationEntryReport validation) {
    WorkbookInvariantAnalysisFormattingChecks.requireDataValidationEntryShape(validation);
  }

  static void requireAutofilterEntryShape(AutofilterEntryReport autofilter) {
    WorkbookInvariantAnalysisLayoutChecks.requireAutofilterEntryShape(autofilter);
  }

  static void requireConditionalFormattingEntryShape(
      ConditionalFormattingEntryReport conditionalFormatting) {
    WorkbookInvariantAnalysisFormattingChecks.requireConditionalFormattingEntryShape(
        conditionalFormatting);
  }

  static void requireTableEntryShape(TableEntryReport table) {
    WorkbookInvariantAnalysisLayoutChecks.requireTableEntryShape(table);
  }

  static void requirePivotTableShape(PivotTableReport pivotTable) {
    WorkbookInvariantAnalysisLayoutChecks.requirePivotTableShape(pivotTable);
  }

  static void requireConditionalFormattingRuleShape(ConditionalFormattingRuleReport rule) {
    WorkbookInvariantAnalysisFormattingChecks.requireConditionalFormattingRuleShape(rule);
  }

  static void requireConditionalFormattingThresholdShape(
      ConditionalFormattingThresholdReport threshold) {
    WorkbookInvariantAnalysisFormattingChecks.requireConditionalFormattingThresholdShape(threshold);
  }

  static void requireDifferentialStyleShape(DifferentialStyleReport style) {
    WorkbookInvariantAnalysisFormattingChecks.requireDifferentialStyleShape(style);
  }

  static void requireDifferentialBorderShape(DifferentialBorderReport border) {
    WorkbookInvariantAnalysisFormattingChecks.requireDifferentialBorderShape(border);
  }

  static void requireDifferentialBorderSideShape(DifferentialBorderSideReport side) {
    WorkbookInvariantAnalysisFormattingChecks.requireDifferentialBorderSideShape(side);
  }

  static void requireTableStyleShape(TableStyleReport style) {
    WorkbookInvariantAnalysisFormattingChecks.requireTableStyleShape(style);
  }

  static void requireSupportedDataValidationShape(
      DataValidationEntryReport.DataValidationDefinitionReport validation) {
    WorkbookInvariantAnalysisFormattingChecks.requireSupportedDataValidationShape(validation);
  }

  static void requireComparisonRuleShape(Object operator, String formula1) {
    WorkbookInvariantAnalysisFormattingChecks.requireComparisonRuleShape(operator, formula1);
  }

  static void requireFormulaSurfaceShape(FormulaSurfaceReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireFormulaSurfaceShape(analysis);
  }

  static void requireSheetSchemaShape(SheetSchemaReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireSheetSchemaShape(analysis);
  }

  static void requireNamedRangeSurfaceShape(NamedRangeSurfaceReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireNamedRangeSurfaceShape(analysis);
  }

  static void requireFormulaHealthShape(FormulaHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireFormulaHealthShape(analysis);
  }

  static void requireDataValidationHealthShape(
      dev.erst.gridgrind.contract.dto.DataValidationHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireDataValidationHealthShape(analysis);
  }

  static void requireConditionalFormattingHealthShape(ConditionalFormattingHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireConditionalFormattingHealthShape(analysis);
  }

  static void requireAutofilterHealthShape(AutofilterHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireAutofilterHealthShape(analysis);
  }

  static void requireTableHealthShape(TableHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireTableHealthShape(analysis);
  }

  static void requirePivotTableHealthShape(PivotTableHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requirePivotTableHealthShape(analysis);
  }

  static void requireHyperlinkHealthShape(HyperlinkHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireHyperlinkHealthShape(analysis);
  }

  static void requireNamedRangeHealthShape(NamedRangeHealthReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireNamedRangeHealthShape(analysis);
  }

  static void requireWorkbookFindingsShape(WorkbookFindingsReport analysis) {
    WorkbookInvariantAnalysisReportChecks.requireWorkbookFindingsShape(analysis);
  }

  static void requireAnalysisSummaryShape(
      AnalysisSummaryReport summary, List<AnalysisFindingReport> findings) {
    WorkbookInvariantAnalysisReportChecks.requireAnalysisSummaryShape(summary, findings);
  }

  static void requireAnalysisFindingShape(AnalysisFindingReport finding) {
    WorkbookInvariantAnalysisReportChecks.requireAnalysisFindingShape(finding);
  }
}
