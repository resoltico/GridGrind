package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** Ordered result payload returned for one requested workbook read. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = WorkbookReadResult.WorkbookSummaryResult.class,
      name = "GET_WORKBOOK_SUMMARY"),
  @JsonSubTypes.Type(value = WorkbookReadResult.NamedRangesResult.class, name = "GET_NAMED_RANGES"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.SheetSummaryResult.class,
      name = "GET_SHEET_SUMMARY"),
  @JsonSubTypes.Type(value = WorkbookReadResult.CellsResult.class, name = "GET_CELLS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.WindowResult.class, name = "GET_WINDOW"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.MergedRegionsResult.class,
      name = "GET_MERGED_REGIONS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.HyperlinksResult.class, name = "GET_HYPERLINKS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.CommentsResult.class, name = "GET_COMMENTS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.SheetLayoutResult.class, name = "GET_SHEET_LAYOUT"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.DataValidationsResult.class,
      name = "GET_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.ConditionalFormattingResult.class,
      name = "GET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = WorkbookReadResult.AutofiltersResult.class, name = "GET_AUTOFILTERS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.TablesResult.class, name = "GET_TABLES"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.FormulaSurfaceResult.class,
      name = "GET_FORMULA_SURFACE"),
  @JsonSubTypes.Type(value = WorkbookReadResult.SheetSchemaResult.class, name = "GET_SHEET_SCHEMA"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.NamedRangeSurfaceResult.class,
      name = "GET_NAMED_RANGE_SURFACE"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.FormulaHealthResult.class,
      name = "ANALYZE_FORMULA_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.DataValidationHealthResult.class,
      name = "ANALYZE_DATA_VALIDATION_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.ConditionalFormattingHealthResult.class,
      name = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.AutofilterHealthResult.class,
      name = "ANALYZE_AUTOFILTER_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.TableHealthResult.class,
      name = "ANALYZE_TABLE_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.HyperlinkHealthResult.class,
      name = "ANALYZE_HYPERLINK_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.NamedRangeHealthResult.class,
      name = "ANALYZE_NAMED_RANGE_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.WorkbookFindingsResult.class,
      name = "ANALYZE_WORKBOOK_FINDINGS")
})
public sealed interface WorkbookReadResult
    permits WorkbookReadResult.Introspection, WorkbookReadResult.Analysis {

  /** Stable caller-provided identifier copied from the matching read operation. */
  String requestId();

  /** Marker for fact-only workbook reads. */
  sealed interface Introspection extends WorkbookReadResult
      permits WorkbookSummaryResult,
          NamedRangesResult,
          SheetSummaryResult,
          CellsResult,
          WindowResult,
          MergedRegionsResult,
          HyperlinksResult,
          CommentsResult,
          SheetLayoutResult,
          DataValidationsResult,
          ConditionalFormattingResult,
          AutofiltersResult,
          TablesResult,
          FormulaSurfaceResult,
          SheetSchemaResult,
          NamedRangeSurfaceResult {}

  /** Marker for derived workbook analysis results. */
  sealed interface Analysis extends WorkbookReadResult
      permits FormulaHealthResult,
          DataValidationHealthResult,
          ConditionalFormattingHealthResult,
          AutofilterHealthResult,
          TableHealthResult,
          HyperlinkHealthResult,
          NamedRangeHealthResult,
          WorkbookFindingsResult {}

  /** Returns workbook-level summary facts. */
  record WorkbookSummaryResult(String requestId, GridGrindResponse.WorkbookSummary workbook)
      implements Introspection {
    public WorkbookSummaryResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns named ranges selected by the originating read operation. */
  record NamedRangesResult(String requestId, List<GridGrindResponse.NamedRangeReport> namedRanges)
      implements Introspection {
    public NamedRangesResult {
      requestId = requireNonBlank(requestId, "requestId");
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }

  /** Returns summary facts for one sheet. */
  record SheetSummaryResult(String requestId, GridGrindResponse.SheetSummaryReport sheet)
      implements Introspection {
    public SheetSummaryResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  record CellsResult(String requestId, String sheetName, List<GridGrindResponse.CellReport> cells)
      implements Introspection {
    public CellsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      cells = copyValues(cells, "cells");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  record WindowResult(String requestId, GridGrindResponse.WindowReport window)
      implements Introspection {
    public WindowResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(window, "window must not be null");
    }
  }

  /** Returns every merged region present on one sheet. */
  record MergedRegionsResult(
      String requestId, String sheetName, List<GridGrindResponse.MergedRegionReport> mergedRegions)
      implements Introspection {
    public MergedRegionsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      mergedRegions = copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record HyperlinksResult(
      String requestId, String sheetName, List<GridGrindResponse.CellHyperlinkReport> hyperlinks)
      implements Introspection {
    public HyperlinksResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      hyperlinks = copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record CommentsResult(
      String requestId, String sheetName, List<GridGrindResponse.CellCommentReport> comments)
      implements Introspection {
    public CommentsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      comments = copyValues(comments, "comments");
    }
  }

  /** Returns layout facts such as freeze panes and explicit sizing. */
  record SheetLayoutResult(String requestId, GridGrindResponse.SheetLayoutReport layout)
      implements Introspection {
    public SheetLayoutResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record DataValidationsResult(
      String requestId, String sheetName, List<DataValidationEntryReport> validations)
      implements Introspection {
    public DataValidationsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      validations = copyValues(validations, "validations");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record ConditionalFormattingResult(
      String requestId,
      String sheetName,
      List<ConditionalFormattingEntryReport> conditionalFormattingBlocks)
      implements Introspection {
    public ConditionalFormattingResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      conditionalFormattingBlocks =
          copyValues(conditionalFormattingBlocks, "conditionalFormattingBlocks");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record AutofiltersResult(
      String requestId, String sheetName, List<AutofilterEntryReport> autofilters)
      implements Introspection {
    public AutofiltersResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      autofilters = copyValues(autofilters, "autofilters");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record TablesResult(String requestId, List<TableEntryReport> tables) implements Introspection {
    public TablesResult {
      requestId = requireNonBlank(requestId, "requestId");
      tables = copyValues(tables, "tables");
    }
  }

  /** Returns grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceResult(String requestId, GridGrindResponse.FormulaSurfaceReport analysis)
      implements Introspection {
    public FormulaSurfaceResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns inferred schema facts for one sheet window. */
  record SheetSchemaResult(String requestId, GridGrindResponse.SheetSchemaReport analysis)
      implements Introspection {
    public SheetSchemaResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns high-level characterization of named ranges. */
  record NamedRangeSurfaceResult(
      String requestId, GridGrindResponse.NamedRangeSurfaceReport analysis)
      implements Introspection {
    public NamedRangeSurfaceResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns formula-health findings. */
  record FormulaHealthResult(String requestId, GridGrindResponse.FormulaHealthReport analysis)
      implements Analysis {
    public FormulaHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns data-validation-health findings. */
  record DataValidationHealthResult(String requestId, DataValidationHealthReport analysis)
      implements Analysis {
    public DataValidationHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns conditional-formatting-health findings. */
  record ConditionalFormattingHealthResult(
      String requestId, ConditionalFormattingHealthReport analysis) implements Analysis {
    public ConditionalFormattingHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns autofilter-health findings. */
  record AutofilterHealthResult(String requestId, AutofilterHealthReport analysis)
      implements Analysis {
    public AutofilterHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns table-health findings. */
  record TableHealthResult(String requestId, TableHealthReport analysis) implements Analysis {
    public TableHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings. */
  record HyperlinkHealthResult(String requestId, GridGrindResponse.HyperlinkHealthReport analysis)
      implements Analysis {
    public HyperlinkHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings. */
  record NamedRangeHealthResult(String requestId, GridGrindResponse.NamedRangeHealthReport analysis)
      implements Analysis {
    public NamedRangeHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns aggregated workbook findings. */
  record WorkbookFindingsResult(String requestId, GridGrindResponse.WorkbookFindingsReport analysis)
      implements Analysis {
    public WorkbookFindingsResult {
      requestId = requireNonBlank(requestId, "requestId");
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
    List<T> copy = List.copyOf(values);
    for (T value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }
}
