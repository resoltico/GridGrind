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
      name = "WORKBOOK_SUMMARY"),
  @JsonSubTypes.Type(value = WorkbookReadResult.NamedRangesResult.class, name = "NAMED_RANGES"),
  @JsonSubTypes.Type(value = WorkbookReadResult.SheetSummaryResult.class, name = "SHEET_SUMMARY"),
  @JsonSubTypes.Type(value = WorkbookReadResult.CellsResult.class, name = "CELLS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.WindowResult.class, name = "WINDOW"),
  @JsonSubTypes.Type(value = WorkbookReadResult.MergedRegionsResult.class, name = "MERGED_REGIONS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.HyperlinksResult.class, name = "HYPERLINKS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.CommentsResult.class, name = "COMMENTS"),
  @JsonSubTypes.Type(value = WorkbookReadResult.SheetLayoutResult.class, name = "SHEET_LAYOUT"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.FormulaSurfaceResult.class,
      name = "FORMULA_SURFACE"),
  @JsonSubTypes.Type(value = WorkbookReadResult.SheetSchemaResult.class, name = "SHEET_SCHEMA"),
  @JsonSubTypes.Type(
      value = WorkbookReadResult.NamedRangeSurfaceResult.class,
      name = "NAMED_RANGE_SURFACE")
})
public sealed interface WorkbookReadResult
    permits WorkbookReadResult.Introspection, WorkbookReadResult.Insight {

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
          SheetLayoutResult {}

  /** Marker for derived workbook analysis results. */
  sealed interface Insight extends WorkbookReadResult
      permits FormulaSurfaceResult, SheetSchemaResult, NamedRangeSurfaceResult {}

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

  /** Returns grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceResult(String requestId, GridGrindResponse.FormulaSurfaceReport analysis)
      implements Insight {
    public FormulaSurfaceResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns inferred schema facts for one sheet window. */
  record SheetSchemaResult(String requestId, GridGrindResponse.SheetSchemaReport analysis)
      implements Insight {
    public SheetSchemaResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns high-level characterization of named ranges. */
  record NamedRangeSurfaceResult(
      String requestId, GridGrindResponse.NamedRangeSurfaceReport analysis) implements Insight {
    public NamedRangeSurfaceResult {
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
