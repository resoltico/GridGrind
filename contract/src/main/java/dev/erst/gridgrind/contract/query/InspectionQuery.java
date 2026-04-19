package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Ordered post-mutation inspection queries that introspect or analyze workbook state. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = InspectionQuery.GetWorkbookSummary.class,
      name = "GET_WORKBOOK_SUMMARY"),
  @JsonSubTypes.Type(
      value = InspectionQuery.GetPackageSecurity.class,
      name = "GET_PACKAGE_SECURITY"),
  @JsonSubTypes.Type(
      value = InspectionQuery.GetWorkbookProtection.class,
      name = "GET_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(value = InspectionQuery.GetNamedRanges.class, name = "GET_NAMED_RANGES"),
  @JsonSubTypes.Type(value = InspectionQuery.GetSheetSummary.class, name = "GET_SHEET_SUMMARY"),
  @JsonSubTypes.Type(value = InspectionQuery.GetCells.class, name = "GET_CELLS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetWindow.class, name = "GET_WINDOW"),
  @JsonSubTypes.Type(value = InspectionQuery.GetMergedRegions.class, name = "GET_MERGED_REGIONS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetHyperlinks.class, name = "GET_HYPERLINKS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetComments.class, name = "GET_COMMENTS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetDrawingObjects.class, name = "GET_DRAWING_OBJECTS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetCharts.class, name = "GET_CHARTS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetPivotTables.class, name = "GET_PIVOT_TABLES"),
  @JsonSubTypes.Type(
      value = InspectionQuery.GetDrawingObjectPayload.class,
      name = "GET_DRAWING_OBJECT_PAYLOAD"),
  @JsonSubTypes.Type(value = InspectionQuery.GetSheetLayout.class, name = "GET_SHEET_LAYOUT"),
  @JsonSubTypes.Type(value = InspectionQuery.GetPrintLayout.class, name = "GET_PRINT_LAYOUT"),
  @JsonSubTypes.Type(
      value = InspectionQuery.GetDataValidations.class,
      name = "GET_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = InspectionQuery.GetConditionalFormatting.class,
      name = "GET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = InspectionQuery.GetAutofilters.class, name = "GET_AUTOFILTERS"),
  @JsonSubTypes.Type(value = InspectionQuery.GetTables.class, name = "GET_TABLES"),
  @JsonSubTypes.Type(value = InspectionQuery.GetFormulaSurface.class, name = "GET_FORMULA_SURFACE"),
  @JsonSubTypes.Type(value = InspectionQuery.GetSheetSchema.class, name = "GET_SHEET_SCHEMA"),
  @JsonSubTypes.Type(
      value = InspectionQuery.GetNamedRangeSurface.class,
      name = "GET_NAMED_RANGE_SURFACE"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeFormulaHealth.class,
      name = "ANALYZE_FORMULA_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeDataValidationHealth.class,
      name = "ANALYZE_DATA_VALIDATION_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeConditionalFormattingHealth.class,
      name = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeAutofilterHealth.class,
      name = "ANALYZE_AUTOFILTER_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeTableHealth.class,
      name = "ANALYZE_TABLE_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzePivotTableHealth.class,
      name = "ANALYZE_PIVOT_TABLE_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeHyperlinkHealth.class,
      name = "ANALYZE_HYPERLINK_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeNamedRangeHealth.class,
      name = "ANALYZE_NAMED_RANGE_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionQuery.AnalyzeWorkbookFindings.class,
      name = "ANALYZE_WORKBOOK_FINDINGS")
})
public sealed interface InspectionQuery
    permits InspectionQuery.Introspection, InspectionQuery.Analysis {

  /**
   * Maximum number of cells ({@code rowCount * columnCount}) permitted in one rectangular-window
   * selector. Requests exceeding this limit are rejected during plan validation to prevent
   * out-of-memory failures during serialization of large cell grids. See docs/LIMITATIONS.md
   * LIM-001.
   */
  int MAX_WINDOW_CELLS = 250_000; // LIM-001

  /** Marker for raw workbook-fact queries with no higher-level interpretation. */
  sealed interface Introspection extends InspectionQuery
      permits GetWorkbookSummary,
          GetPackageSecurity,
          GetWorkbookProtection,
          GetNamedRanges,
          GetSheetSummary,
          GetCells,
          GetWindow,
          GetMergedRegions,
          GetHyperlinks,
          GetComments,
          GetDrawingObjects,
          GetCharts,
          GetPivotTables,
          GetDrawingObjectPayload,
          GetSheetLayout,
          GetPrintLayout,
          GetDataValidations,
          GetConditionalFormatting,
          GetAutofilters,
          GetTables,
          GetFormulaSurface,
          GetSheetSchema,
          GetNamedRangeSurface {}

  /** Marker for derived workbook analysis queries. */
  sealed interface Analysis extends InspectionQuery
      permits AnalyzeFormulaHealth,
          AnalyzeDataValidationHealth,
          AnalyzeConditionalFormattingHealth,
          AnalyzeAutofilterHealth,
          AnalyzeTableHealth,
          AnalyzePivotTableHealth,
          AnalyzeHyperlinkHealth,
          AnalyzeNamedRangeHealth,
          AnalyzeWorkbookFindings {}

  record GetWorkbookSummary() implements Introspection {}

  record GetPackageSecurity() implements Introspection {}

  record GetWorkbookProtection() implements Introspection {}

  record GetNamedRanges() implements Introspection {}

  record GetSheetSummary() implements Introspection {}

  record GetCells() implements Introspection {}

  record GetWindow() implements Introspection {}

  record GetMergedRegions() implements Introspection {}

  record GetHyperlinks() implements Introspection {}

  record GetComments() implements Introspection {}

  record GetDrawingObjects() implements Introspection {}

  record GetCharts() implements Introspection {}

  record GetPivotTables() implements Introspection {}

  record GetDrawingObjectPayload() implements Introspection {}

  record GetSheetLayout() implements Introspection {}

  record GetPrintLayout() implements Introspection {}

  record GetDataValidations() implements Introspection {}

  record GetConditionalFormatting() implements Introspection {}

  record GetAutofilters() implements Introspection {}

  record GetTables() implements Introspection {}

  record GetFormulaSurface() implements Introspection {}

  record GetSheetSchema() implements Introspection {}

  record GetNamedRangeSurface() implements Introspection {}

  record AnalyzeFormulaHealth() implements Analysis {}

  record AnalyzeDataValidationHealth() implements Analysis {}

  record AnalyzeConditionalFormattingHealth() implements Analysis {}

  record AnalyzeAutofilterHealth() implements Analysis {}

  record AnalyzeTableHealth() implements Analysis {}

  record AnalyzePivotTableHealth() implements Analysis {}

  record AnalyzeHyperlinkHealth() implements Analysis {}

  record AnalyzeNamedRangeHealth() implements Analysis {}

  record AnalyzeWorkbookFindings() implements Analysis {}

  /** Returns the stable SCREAMING_SNAKE_CASE discriminator for one inspection query. */
  default String queryType() {
    return switch (this) {
      case GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case GetPackageSecurity _ -> "GET_PACKAGE_SECURITY";
      case GetWorkbookProtection _ -> "GET_WORKBOOK_PROTECTION";
      case GetNamedRanges _ -> "GET_NAMED_RANGES";
      case GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case GetCells _ -> "GET_CELLS";
      case GetWindow _ -> "GET_WINDOW";
      case GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case GetHyperlinks _ -> "GET_HYPERLINKS";
      case GetComments _ -> "GET_COMMENTS";
      case GetDrawingObjects _ -> "GET_DRAWING_OBJECTS";
      case GetCharts _ -> "GET_CHARTS";
      case GetPivotTables _ -> "GET_PIVOT_TABLES";
      case GetDrawingObjectPayload _ -> "GET_DRAWING_OBJECT_PAYLOAD";
      case GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case GetPrintLayout _ -> "GET_PRINT_LAYOUT";
      case GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case GetAutofilters _ -> "GET_AUTOFILTERS";
      case GetTables _ -> "GET_TABLES";
      case GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
      case AnalyzeFormulaHealth _ -> "ANALYZE_FORMULA_HEALTH";
      case AnalyzeDataValidationHealth _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case AnalyzeConditionalFormattingHealth _ -> "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case AnalyzeAutofilterHealth _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case AnalyzeTableHealth _ -> "ANALYZE_TABLE_HEALTH";
      case AnalyzePivotTableHealth _ -> "ANALYZE_PIVOT_TABLE_HEALTH";
      case AnalyzeHyperlinkHealth _ -> "ANALYZE_HYPERLINK_HEALTH";
      case AnalyzeNamedRangeHealth _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case AnalyzeWorkbookFindings _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }
}
