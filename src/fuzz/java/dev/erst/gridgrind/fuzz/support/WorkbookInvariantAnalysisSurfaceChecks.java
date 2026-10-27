package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.formula.FormulaHealthReport;
import dev.erst.gridgrind.core.model.formula.FormulaSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableHealthReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderSideReport;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterHealthReport;
import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableHealthReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.HyperlinkHealthReport;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookFindingsReport;
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
      dev.erst.gridgrind.core.model.validation.DataValidationHealthReport analysis) {
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
