package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionQuery;

/** Canonical inspection-query helpers kept internal to the focused Java authoring surface. */
final class Queries {
  private Queries() {}

  static InspectionQuery.GetWorkbookSummary workbookSummary() {
    return new InspectionQuery.GetWorkbookSummary();
  }

  static InspectionQuery.GetPackageSecurity packageSecurity() {
    return new InspectionQuery.GetPackageSecurity();
  }

  static InspectionQuery.GetWorkbookProtection workbookProtection() {
    return new InspectionQuery.GetWorkbookProtection();
  }

  static InspectionQuery.GetNamedRanges namedRanges() {
    return new InspectionQuery.GetNamedRanges();
  }

  static InspectionQuery.GetSheetSummary sheetSummary() {
    return new InspectionQuery.GetSheetSummary();
  }

  static InspectionQuery.GetCells cells() {
    return new InspectionQuery.GetCells();
  }

  static InspectionQuery.GetWindow window() {
    return new InspectionQuery.GetWindow();
  }

  static InspectionQuery.GetMergedRegions mergedRegions() {
    return new InspectionQuery.GetMergedRegions();
  }

  static InspectionQuery.GetHyperlinks hyperlinks() {
    return new InspectionQuery.GetHyperlinks();
  }

  static InspectionQuery.GetComments comments() {
    return new InspectionQuery.GetComments();
  }

  static InspectionQuery.GetDrawingObjects drawingObjects() {
    return new InspectionQuery.GetDrawingObjects();
  }

  static InspectionQuery.GetCharts charts() {
    return new InspectionQuery.GetCharts();
  }

  static InspectionQuery.GetPivotTables pivotTables() {
    return new InspectionQuery.GetPivotTables();
  }

  static InspectionQuery.GetDrawingObjectPayload drawingObjectPayload() {
    return new InspectionQuery.GetDrawingObjectPayload();
  }

  static InspectionQuery.GetSheetLayout sheetLayout() {
    return new InspectionQuery.GetSheetLayout();
  }

  static InspectionQuery.GetPrintLayout printLayout() {
    return new InspectionQuery.GetPrintLayout();
  }

  static InspectionQuery.GetDataValidations dataValidations() {
    return new InspectionQuery.GetDataValidations();
  }

  static InspectionQuery.GetConditionalFormatting conditionalFormatting() {
    return new InspectionQuery.GetConditionalFormatting();
  }

  static InspectionQuery.GetAutofilters autofilters() {
    return new InspectionQuery.GetAutofilters();
  }

  static InspectionQuery.GetTables tables() {
    return new InspectionQuery.GetTables();
  }

  static InspectionQuery.GetFormulaSurface formulaSurface() {
    return new InspectionQuery.GetFormulaSurface();
  }

  static InspectionQuery.GetSheetSchema sheetSchema() {
    return new InspectionQuery.GetSheetSchema();
  }

  static InspectionQuery.GetNamedRangeSurface namedRangeSurface() {
    return new InspectionQuery.GetNamedRangeSurface();
  }

  static InspectionQuery.AnalyzeFormulaHealth formulaHealth() {
    return new InspectionQuery.AnalyzeFormulaHealth();
  }

  static InspectionQuery.AnalyzeDataValidationHealth dataValidationHealth() {
    return new InspectionQuery.AnalyzeDataValidationHealth();
  }

  static InspectionQuery.AnalyzeConditionalFormattingHealth conditionalFormattingHealth() {
    return new InspectionQuery.AnalyzeConditionalFormattingHealth();
  }

  static InspectionQuery.AnalyzeAutofilterHealth autofilterHealth() {
    return new InspectionQuery.AnalyzeAutofilterHealth();
  }

  static InspectionQuery.AnalyzeTableHealth tableHealth() {
    return new InspectionQuery.AnalyzeTableHealth();
  }

  static InspectionQuery.AnalyzePivotTableHealth pivotTableHealth() {
    return new InspectionQuery.AnalyzePivotTableHealth();
  }

  static InspectionQuery.AnalyzeHyperlinkHealth hyperlinkHealth() {
    return new InspectionQuery.AnalyzeHyperlinkHealth();
  }

  static InspectionQuery.AnalyzeNamedRangeHealth namedRangeHealth() {
    return new InspectionQuery.AnalyzeNamedRangeHealth();
  }

  static InspectionQuery.AnalyzeWorkbookFindings workbookFindings() {
    return new InspectionQuery.AnalyzeWorkbookFindings();
  }
}
