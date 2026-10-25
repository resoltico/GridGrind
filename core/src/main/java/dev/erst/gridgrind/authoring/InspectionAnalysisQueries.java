package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;

/** Canonical analysis-query factories kept internal to the Java authoring surface. */
final class InspectionAnalysisQueries {
  private InspectionAnalysisQueries() {}

  static InspectionAnalysisQuery.AnalyzeFormulaHealth formulaHealth() {
    return new InspectionAnalysisQuery.AnalyzeFormulaHealth();
  }

  static InspectionAnalysisQuery.AnalyzeDataValidationHealth dataValidationHealth() {
    return new InspectionAnalysisQuery.AnalyzeDataValidationHealth();
  }

  static InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth conditionalFormattingHealth() {
    return new InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth();
  }

  static InspectionAnalysisQuery.AnalyzeAutofilterHealth autofilterHealth() {
    return new InspectionAnalysisQuery.AnalyzeAutofilterHealth();
  }

  static InspectionAnalysisQuery.AnalyzeTableHealth tableHealth() {
    return new InspectionAnalysisQuery.AnalyzeTableHealth();
  }

  static InspectionAnalysisQuery.AnalyzePivotTableHealth pivotTableHealth() {
    return new InspectionAnalysisQuery.AnalyzePivotTableHealth();
  }

  static InspectionAnalysisQuery.AnalyzeHyperlinkHealth hyperlinkHealth() {
    return new InspectionAnalysisQuery.AnalyzeHyperlinkHealth();
  }

  static InspectionAnalysisQuery.AnalyzeNamedRangeHealth namedRangeHealth() {
    return new InspectionAnalysisQuery.AnalyzeNamedRangeHealth();
  }

  static InspectionAnalysisQuery.AnalyzeWorkbookFindings workbookFindings() {
    return new InspectionAnalysisQuery.AnalyzeWorkbookFindings();
  }
}
