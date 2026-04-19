package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionQuery;

/** Canonical inspection-query builders for the Java authoring layer. */
public final class Queries {
  private Queries() {}

  /** Returns the workbook-summary inspection query. */
  public static InspectionQuery.GetWorkbookSummary workbookSummary() {
    return new InspectionQuery.GetWorkbookSummary();
  }

  /** Returns the package-security inspection query. */
  public static InspectionQuery.GetPackageSecurity packageSecurity() {
    return new InspectionQuery.GetPackageSecurity();
  }

  /** Returns the workbook-protection inspection query. */
  public static InspectionQuery.GetWorkbookProtection workbookProtection() {
    return new InspectionQuery.GetWorkbookProtection();
  }

  /** Returns the named-range inventory inspection query. */
  public static InspectionQuery.GetNamedRanges namedRanges() {
    return new InspectionQuery.GetNamedRanges();
  }

  /** Returns the sheet-summary inspection query. */
  public static InspectionQuery.GetSheetSummary sheetSummary() {
    return new InspectionQuery.GetSheetSummary();
  }

  /** Returns the cell inspection query. */
  public static InspectionQuery.GetCells cells() {
    return new InspectionQuery.GetCells();
  }

  /** Returns the rectangular-window inspection query. */
  public static InspectionQuery.GetWindow window() {
    return new InspectionQuery.GetWindow();
  }

  /** Returns the merged-regions inspection query. */
  public static InspectionQuery.GetMergedRegions mergedRegions() {
    return new InspectionQuery.GetMergedRegions();
  }

  /** Returns the hyperlink inspection query. */
  public static InspectionQuery.GetHyperlinks hyperlinks() {
    return new InspectionQuery.GetHyperlinks();
  }

  /** Returns the comment inspection query. */
  public static InspectionQuery.GetComments comments() {
    return new InspectionQuery.GetComments();
  }

  /** Returns the drawing-object inventory query. */
  public static InspectionQuery.GetDrawingObjects drawingObjects() {
    return new InspectionQuery.GetDrawingObjects();
  }

  /** Returns the chart inventory query. */
  public static InspectionQuery.GetCharts charts() {
    return new InspectionQuery.GetCharts();
  }

  /** Returns the pivot-table inventory query. */
  public static InspectionQuery.GetPivotTables pivotTables() {
    return new InspectionQuery.GetPivotTables();
  }

  /** Returns the drawing-object payload query. */
  public static InspectionQuery.GetDrawingObjectPayload drawingObjectPayload() {
    return new InspectionQuery.GetDrawingObjectPayload();
  }

  /** Returns the sheet-layout inspection query. */
  public static InspectionQuery.GetSheetLayout sheetLayout() {
    return new InspectionQuery.GetSheetLayout();
  }

  /** Returns the print-layout inspection query. */
  public static InspectionQuery.GetPrintLayout printLayout() {
    return new InspectionQuery.GetPrintLayout();
  }

  /** Returns the data-validation inspection query. */
  public static InspectionQuery.GetDataValidations dataValidations() {
    return new InspectionQuery.GetDataValidations();
  }

  /** Returns the conditional-formatting inspection query. */
  public static InspectionQuery.GetConditionalFormatting conditionalFormatting() {
    return new InspectionQuery.GetConditionalFormatting();
  }

  /** Returns the autofilter inspection query. */
  public static InspectionQuery.GetAutofilters autofilters() {
    return new InspectionQuery.GetAutofilters();
  }

  /** Returns the table inventory query. */
  public static InspectionQuery.GetTables tables() {
    return new InspectionQuery.GetTables();
  }

  /** Returns the formula-surface inspection query. */
  public static InspectionQuery.GetFormulaSurface formulaSurface() {
    return new InspectionQuery.GetFormulaSurface();
  }

  /** Returns the sheet-schema inspection query. */
  public static InspectionQuery.GetSheetSchema sheetSchema() {
    return new InspectionQuery.GetSheetSchema();
  }

  /** Returns the named-range-surface inspection query. */
  public static InspectionQuery.GetNamedRangeSurface namedRangeSurface() {
    return new InspectionQuery.GetNamedRangeSurface();
  }

  /** Returns the formula-health analysis query. */
  public static InspectionQuery.AnalyzeFormulaHealth formulaHealth() {
    return new InspectionQuery.AnalyzeFormulaHealth();
  }

  /** Returns the data-validation-health analysis query. */
  public static InspectionQuery.AnalyzeDataValidationHealth dataValidationHealth() {
    return new InspectionQuery.AnalyzeDataValidationHealth();
  }

  /** Returns the conditional-formatting-health analysis query. */
  public static InspectionQuery.AnalyzeConditionalFormattingHealth conditionalFormattingHealth() {
    return new InspectionQuery.AnalyzeConditionalFormattingHealth();
  }

  /** Returns the autofilter-health analysis query. */
  public static InspectionQuery.AnalyzeAutofilterHealth autofilterHealth() {
    return new InspectionQuery.AnalyzeAutofilterHealth();
  }

  /** Returns the table-health analysis query. */
  public static InspectionQuery.AnalyzeTableHealth tableHealth() {
    return new InspectionQuery.AnalyzeTableHealth();
  }

  /** Returns the pivot-table-health analysis query. */
  public static InspectionQuery.AnalyzePivotTableHealth pivotTableHealth() {
    return new InspectionQuery.AnalyzePivotTableHealth();
  }

  /** Returns the hyperlink-health analysis query. */
  public static InspectionQuery.AnalyzeHyperlinkHealth hyperlinkHealth() {
    return new InspectionQuery.AnalyzeHyperlinkHealth();
  }

  /** Returns the named-range-health analysis query. */
  public static InspectionQuery.AnalyzeNamedRangeHealth namedRangeHealth() {
    return new InspectionQuery.AnalyzeNamedRangeHealth();
  }

  /** Returns the aggregate workbook-findings analysis query. */
  public static InspectionQuery.AnalyzeWorkbookFindings workbookFindings() {
    return new InspectionQuery.AnalyzeWorkbookFindings();
  }
}
