package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Ordered result payload returned for one requested inspection result. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = InspectionResult.WorkbookSummaryResult.class,
      name = "GET_WORKBOOK_SUMMARY"),
  @JsonSubTypes.Type(
      value = InspectionResult.PackageSecurityResult.class,
      name = "GET_PACKAGE_SECURITY"),
  @JsonSubTypes.Type(
      value = InspectionResult.WorkbookProtectionResult.class,
      name = "GET_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(value = InspectionResult.NamedRangesResult.class, name = "GET_NAMED_RANGES"),
  @JsonSubTypes.Type(value = InspectionResult.SheetSummaryResult.class, name = "GET_SHEET_SUMMARY"),
  @JsonSubTypes.Type(value = InspectionResult.CellsResult.class, name = "GET_CELLS"),
  @JsonSubTypes.Type(value = InspectionResult.WindowResult.class, name = "GET_WINDOW"),
  @JsonSubTypes.Type(
      value = InspectionResult.MergedRegionsResult.class,
      name = "GET_MERGED_REGIONS"),
  @JsonSubTypes.Type(value = InspectionResult.HyperlinksResult.class, name = "GET_HYPERLINKS"),
  @JsonSubTypes.Type(value = InspectionResult.CommentsResult.class, name = "GET_COMMENTS"),
  @JsonSubTypes.Type(
      value = InspectionResult.DrawingObjectsResult.class,
      name = "GET_DRAWING_OBJECTS"),
  @JsonSubTypes.Type(value = InspectionResult.ChartsResult.class, name = "GET_CHARTS"),
  @JsonSubTypes.Type(value = InspectionResult.PivotTablesResult.class, name = "GET_PIVOT_TABLES"),
  @JsonSubTypes.Type(
      value = InspectionResult.DrawingObjectPayloadResult.class,
      name = "GET_DRAWING_OBJECT_PAYLOAD"),
  @JsonSubTypes.Type(value = InspectionResult.SheetLayoutResult.class, name = "GET_SHEET_LAYOUT"),
  @JsonSubTypes.Type(value = InspectionResult.PrintLayoutResult.class, name = "GET_PRINT_LAYOUT"),
  @JsonSubTypes.Type(
      value = InspectionResult.DataValidationsResult.class,
      name = "GET_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = InspectionResult.ConditionalFormattingResult.class,
      name = "GET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = InspectionResult.AutofiltersResult.class, name = "GET_AUTOFILTERS"),
  @JsonSubTypes.Type(value = InspectionResult.TablesResult.class, name = "GET_TABLES"),
  @JsonSubTypes.Type(
      value = InspectionResult.FormulaSurfaceResult.class,
      name = "GET_FORMULA_SURFACE"),
  @JsonSubTypes.Type(value = InspectionResult.SheetSchemaResult.class, name = "GET_SHEET_SCHEMA"),
  @JsonSubTypes.Type(
      value = InspectionResult.NamedRangeSurfaceResult.class,
      name = "GET_NAMED_RANGE_SURFACE"),
  @JsonSubTypes.Type(
      value = InspectionResult.FormulaHealthResult.class,
      name = "ANALYZE_FORMULA_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.DataValidationHealthResult.class,
      name = "ANALYZE_DATA_VALIDATION_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.ConditionalFormattingHealthResult.class,
      name = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.AutofilterHealthResult.class,
      name = "ANALYZE_AUTOFILTER_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.TableHealthResult.class,
      name = "ANALYZE_TABLE_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.PivotTableHealthResult.class,
      name = "ANALYZE_PIVOT_TABLE_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.HyperlinkHealthResult.class,
      name = "ANALYZE_HYPERLINK_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.NamedRangeHealthResult.class,
      name = "ANALYZE_NAMED_RANGE_HEALTH"),
  @JsonSubTypes.Type(
      value = InspectionResult.WorkbookFindingsResult.class,
      name = "ANALYZE_WORKBOOK_FINDINGS")
})
public sealed interface InspectionResult
    permits InspectionResult.Introspection, InspectionResult.Analysis {

  /** Stable caller-provided identifier copied from the matching read operation. */
  String stepId();

  /** Marker for fact-only inspection results. */
  sealed interface Introspection extends InspectionResult
      permits WorkbookSummaryResult,
          PackageSecurityResult,
          WorkbookProtectionResult,
          NamedRangesResult,
          SheetSummaryResult,
          CellsResult,
          WindowResult,
          MergedRegionsResult,
          HyperlinksResult,
          CommentsResult,
          DrawingObjectsResult,
          ChartsResult,
          PivotTablesResult,
          DrawingObjectPayloadResult,
          SheetLayoutResult,
          PrintLayoutResult,
          DataValidationsResult,
          ConditionalFormattingResult,
          AutofiltersResult,
          TablesResult,
          FormulaSurfaceResult,
          SheetSchemaResult,
          NamedRangeSurfaceResult {}

  /** Marker for derived workbook analysis results. */
  sealed interface Analysis extends InspectionResult
      permits FormulaHealthResult,
          DataValidationHealthResult,
          ConditionalFormattingHealthResult,
          AutofilterHealthResult,
          TableHealthResult,
          PivotTableHealthResult,
          HyperlinkHealthResult,
          NamedRangeHealthResult,
          WorkbookFindingsResult {}

  /** Returns workbook-level summary facts. */
  record WorkbookSummaryResult(String stepId, GridGrindResponse.WorkbookSummary workbook)
      implements Introspection {
    public WorkbookSummaryResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts. */
  record PackageSecurityResult(String stepId, OoxmlPackageSecurityReport security)
      implements Introspection {
    public PackageSecurityResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(security, "security must not be null");
    }
  }

  /** Returns workbook-level protection facts. */
  record WorkbookProtectionResult(String stepId, WorkbookProtectionReport protection)
      implements Introspection {
    public WorkbookProtectionResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Returns named ranges selected by the originating read operation. */
  record NamedRangesResult(String stepId, List<GridGrindResponse.NamedRangeReport> namedRanges)
      implements Introspection {
    public NamedRangesResult {
      stepId = requireNonBlank(stepId, "stepId");
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }

  /** Returns summary facts for one sheet. */
  record SheetSummaryResult(String stepId, GridGrindResponse.SheetSummaryReport sheet)
      implements Introspection {
    public SheetSummaryResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  record CellsResult(String stepId, String sheetName, List<GridGrindResponse.CellReport> cells)
      implements Introspection {
    public CellsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      cells = copyValues(cells, "cells");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  record WindowResult(String stepId, GridGrindResponse.WindowReport window)
      implements Introspection {
    public WindowResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(window, "window must not be null");
    }
  }

  /** Returns every merged region present on one sheet. */
  record MergedRegionsResult(
      String stepId, String sheetName, List<GridGrindResponse.MergedRegionReport> mergedRegions)
      implements Introspection {
    public MergedRegionsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      mergedRegions = copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record HyperlinksResult(
      String stepId, String sheetName, List<GridGrindResponse.CellHyperlinkReport> hyperlinks)
      implements Introspection {
    public HyperlinksResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      hyperlinks = copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record CommentsResult(
      String stepId, String sheetName, List<GridGrindResponse.CellCommentReport> comments)
      implements Introspection {
    public CommentsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      comments = copyValues(comments, "comments");
    }
  }

  /** Returns factual drawing-object metadata for one sheet. */
  record DrawingObjectsResult(
      String stepId, String sheetName, List<DrawingObjectReport> drawingObjects)
      implements Introspection {
    public DrawingObjectsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      drawingObjects = copyValues(drawingObjects, "drawingObjects");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  record ChartsResult(String stepId, String sheetName, List<ChartReport> charts)
      implements Introspection {
    public ChartsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      charts = copyValues(charts, "charts");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  record PivotTablesResult(String stepId, List<PivotTableReport> pivotTables)
      implements Introspection {
    public PivotTablesResult {
      stepId = requireNonBlank(stepId, "stepId");
      pivotTables = copyValues(pivotTables, "pivotTables");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object. */
  record DrawingObjectPayloadResult(
      String stepId, String sheetName, DrawingObjectPayloadReport payload)
      implements Introspection {
    public DrawingObjectPayloadResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(payload, "payload must not be null");
    }
  }

  /** Returns layout facts such as pane state, zoom, and explicit sizing. */
  record SheetLayoutResult(String stepId, GridGrindResponse.SheetLayoutReport layout)
      implements Introspection {
    public SheetLayoutResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns supported print-layout facts for one sheet. */
  record PrintLayoutResult(String stepId, PrintLayoutReport layout) implements Introspection {
    public PrintLayoutResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record DataValidationsResult(
      String stepId, String sheetName, List<DataValidationEntryReport> validations)
      implements Introspection {
    public DataValidationsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      validations = copyValues(validations, "validations");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record ConditionalFormattingResult(
      String stepId,
      String sheetName,
      List<ConditionalFormattingEntryReport> conditionalFormattingBlocks)
      implements Introspection {
    public ConditionalFormattingResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      conditionalFormattingBlocks =
          copyValues(conditionalFormattingBlocks, "conditionalFormattingBlocks");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record AutofiltersResult(String stepId, String sheetName, List<AutofilterEntryReport> autofilters)
      implements Introspection {
    public AutofiltersResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      autofilters = copyValues(autofilters, "autofilters");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record TablesResult(String stepId, List<TableEntryReport> tables) implements Introspection {
    public TablesResult {
      stepId = requireNonBlank(stepId, "stepId");
      tables = copyValues(tables, "tables");
    }
  }

  /** Returns grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceResult(String stepId, GridGrindResponse.FormulaSurfaceReport analysis)
      implements Introspection {
    public FormulaSurfaceResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns inferred schema facts for one sheet window. */
  record SheetSchemaResult(String stepId, GridGrindResponse.SheetSchemaReport analysis)
      implements Introspection {
    public SheetSchemaResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns high-level characterization of named ranges. */
  record NamedRangeSurfaceResult(String stepId, GridGrindResponse.NamedRangeSurfaceReport analysis)
      implements Introspection {
    public NamedRangeSurfaceResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns formula-health findings. */
  record FormulaHealthResult(String stepId, GridGrindResponse.FormulaHealthReport analysis)
      implements Analysis {
    public FormulaHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings. */
  record DataValidationHealthResult(String stepId, DataValidationHealthReport analysis)
      implements Analysis {
    public DataValidationHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings. */
  record ConditionalFormattingHealthResult(
      String stepId, ConditionalFormattingHealthReport analysis) implements Analysis {
    public ConditionalFormattingHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings. */
  record AutofilterHealthResult(String stepId, AutofilterHealthReport analysis)
      implements Analysis {
    public AutofilterHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings. */
  record TableHealthResult(String stepId, TableHealthReport analysis) implements Analysis {
    public TableHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns pivot-table-health findings. */
  record PivotTableHealthResult(String stepId, PivotTableHealthReport analysis)
      implements Analysis {
    public PivotTableHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings. */
  record HyperlinkHealthResult(String stepId, GridGrindResponse.HyperlinkHealthReport analysis)
      implements Analysis {
    public HyperlinkHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings. */
  record NamedRangeHealthResult(String stepId, GridGrindResponse.NamedRangeHealthReport analysis)
      implements Analysis {
    public NamedRangeHealthResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns aggregated workbook findings. */
  record WorkbookFindingsResult(String stepId, GridGrindResponse.WorkbookFindingsReport analysis)
      implements Analysis {
    public WorkbookFindingsResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
