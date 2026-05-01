package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.excel.foundation.ExcelReadLimits;

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
  @JsonSubTypes.Type(
      value = InspectionQuery.GetCustomXmlMappings.class,
      name = "GET_CUSTOM_XML_MAPPINGS"),
  @JsonSubTypes.Type(
      value = InspectionQuery.ExportCustomXmlMapping.class,
      name = "EXPORT_CUSTOM_XML_MAPPING"),
  @JsonSubTypes.Type(value = InspectionQuery.GetNamedRanges.class, name = "GET_NAMED_RANGES"),
  @JsonSubTypes.Type(value = InspectionQuery.GetSheetSummary.class, name = "GET_SHEET_SUMMARY"),
  @JsonSubTypes.Type(value = InspectionQuery.GetArrayFormulas.class, name = "GET_ARRAY_FORMULAS"),
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
  int MAX_WINDOW_CELLS = ExcelReadLimits.MAX_WINDOW_CELLS; // LIM-001

  /** Marker for raw workbook-fact queries with no higher-level interpretation. */
  sealed interface Introspection extends InspectionQuery
      permits GetWorkbookSummary,
          GetPackageSecurity,
          GetWorkbookProtection,
          GetCustomXmlMappings,
          ExportCustomXmlMapping,
          GetNamedRanges,
          GetSheetSummary,
          GetArrayFormulas,
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

  record GetCustomXmlMappings() implements Introspection {}

  record ExportCustomXmlMapping(
      dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator mapping,
      boolean validateSchema,
      String encoding)
      implements Introspection {
    /** Creates one export query that uses UTF-8 explicitly for the extracted XML payload. */
    public ExportCustomXmlMapping(
        dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator mapping, boolean validateSchema) {
      this(mapping, validateSchema, java.nio.charset.StandardCharsets.UTF_8.name());
    }

    public ExportCustomXmlMapping {
      java.util.Objects.requireNonNull(mapping, "mapping must not be null");
      java.util.Objects.requireNonNull(encoding, "encoding must not be null");
      if (encoding.isBlank()) {
        throw new IllegalArgumentException("encoding must not be blank");
      }
    }
  }

  record GetNamedRanges() implements Introspection {}

  record GetSheetSummary() implements Introspection {}

  record GetArrayFormulas() implements Introspection {}

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
    return GridGrindProtocolTypeNames.inspectionQueryTypeName(
        getClass().asSubclass(InspectionQuery.class));
  }
}
