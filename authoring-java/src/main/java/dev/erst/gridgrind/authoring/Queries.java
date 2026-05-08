package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.InspectionSurfaceQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;

/** Canonical inspection-query helpers kept internal to the focused Java authoring surface. */
final class Queries {
  private Queries() {}

  static WorkbookIntrospectionQuery.GetWorkbookSummary workbookSummary() {
    return new WorkbookIntrospectionQuery.GetWorkbookSummary();
  }

  static WorkbookIntrospectionQuery.GetPackageSecurity packageSecurity() {
    return new WorkbookIntrospectionQuery.GetPackageSecurity();
  }

  static WorkbookIntrospectionQuery.GetWorkbookProtection workbookProtection() {
    return new WorkbookIntrospectionQuery.GetWorkbookProtection();
  }

  static WorkbookIntrospectionQuery.GetNamedRanges namedRanges() {
    return new WorkbookIntrospectionQuery.GetNamedRanges();
  }

  static SheetIntrospectionQuery.GetSheetSummary sheetSummary() {
    return new SheetIntrospectionQuery.GetSheetSummary();
  }

  static SheetIntrospectionQuery.GetCells cells() {
    return new SheetIntrospectionQuery.GetCells();
  }

  static SheetIntrospectionQuery.GetWindow window() {
    return new SheetIntrospectionQuery.GetWindow();
  }

  static SheetIntrospectionQuery.GetMergedRegions mergedRegions() {
    return new SheetIntrospectionQuery.GetMergedRegions();
  }

  static SheetIntrospectionQuery.GetHyperlinks hyperlinks() {
    return new SheetIntrospectionQuery.GetHyperlinks();
  }

  static SheetIntrospectionQuery.GetComments comments() {
    return new SheetIntrospectionQuery.GetComments();
  }

  static WorkbookAssetIntrospectionQuery.GetDrawingObjects drawingObjects() {
    return new WorkbookAssetIntrospectionQuery.GetDrawingObjects();
  }

  static WorkbookAssetIntrospectionQuery.GetCharts charts() {
    return new WorkbookAssetIntrospectionQuery.GetCharts();
  }

  static WorkbookAssetIntrospectionQuery.GetPivotTables pivotTables() {
    return new WorkbookAssetIntrospectionQuery.GetPivotTables();
  }

  static WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload drawingObjectPayload() {
    return new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload();
  }

  static SheetIntrospectionQuery.GetSheetLayout sheetLayout() {
    return new SheetIntrospectionQuery.GetSheetLayout();
  }

  static SheetIntrospectionQuery.GetPrintLayout printLayout() {
    return new SheetIntrospectionQuery.GetPrintLayout();
  }

  static SheetIntrospectionQuery.GetDataValidations dataValidations() {
    return new SheetIntrospectionQuery.GetDataValidations();
  }

  static SheetIntrospectionQuery.GetConditionalFormatting conditionalFormatting() {
    return new SheetIntrospectionQuery.GetConditionalFormatting();
  }

  static SheetIntrospectionQuery.GetAutofilters autofilters() {
    return new SheetIntrospectionQuery.GetAutofilters();
  }

  static WorkbookAssetIntrospectionQuery.GetTables tables() {
    return new WorkbookAssetIntrospectionQuery.GetTables();
  }

  static InspectionSurfaceQuery.GetFormulaSurface formulaSurface() {
    return new InspectionSurfaceQuery.GetFormulaSurface();
  }

  static InspectionSurfaceQuery.GetSheetSchema sheetSchema() {
    return new InspectionSurfaceQuery.GetSheetSchema();
  }

  static InspectionSurfaceQuery.GetNamedRangeSurface namedRangeSurface() {
    return new InspectionSurfaceQuery.GetNamedRangeSurface();
  }

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
