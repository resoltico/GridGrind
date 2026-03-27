package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Ordered post-mutation workbook reads that introspect or analyze workbook state. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetWorkbookSummary.class,
      name = "GET_WORKBOOK_SUMMARY"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetNamedRanges.class, name = "GET_NAMED_RANGES"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetSheetSummary.class,
      name = "GET_SHEET_SUMMARY"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetCells.class, name = "GET_CELLS"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetWindow.class, name = "GET_WINDOW"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetMergedRegions.class,
      name = "GET_MERGED_REGIONS"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetHyperlinks.class, name = "GET_HYPERLINKS"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetComments.class, name = "GET_COMMENTS"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetSheetLayout.class, name = "GET_SHEET_LAYOUT"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeFormulaSurface.class,
      name = "ANALYZE_FORMULA_SURFACE"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeSheetSchema.class,
      name = "ANALYZE_SHEET_SCHEMA"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeNamedRangeSurface.class,
      name = "ANALYZE_NAMED_RANGE_SURFACE")
})
public sealed interface WorkbookReadOperation
    permits WorkbookReadOperation.Introspection, WorkbookReadOperation.Insight {

  /** Stable caller-provided identifier used to correlate this read with its result. */
  String requestId();

  /** Marker for raw workbook-fact reads with no higher-level interpretation. */
  sealed interface Introspection extends WorkbookReadOperation
      permits GetWorkbookSummary,
          GetNamedRanges,
          GetSheetSummary,
          GetCells,
          GetWindow,
          GetMergedRegions,
          GetHyperlinks,
          GetComments,
          GetSheetLayout {}

  /** Marker for derived workbook analysis built on top of introspection primitives. */
  sealed interface Insight extends WorkbookReadOperation
      permits AnalyzeFormulaSurface, AnalyzeSheetSchema, AnalyzeNamedRangeSurface {}

  /** Returns workbook-level summary facts such as sheet order and recalculation flag. */
  record GetWorkbookSummary(String requestId) implements Introspection {
    public GetWorkbookSummary {
      requestId = requireNonBlank(requestId, "requestId");
    }
  }

  /** Returns named ranges selected by exact selector or workbook-wide selection. */
  record GetNamedRanges(String requestId, NamedRangeSelection selection) implements Introspection {
    public GetNamedRanges {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns structural summary facts for one sheet. */
  record GetSheetSummary(String requestId, String sheetName) implements Introspection {
    public GetSheetSummary {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns exact cell snapshots for the provided ordered A1 addresses. */
  record GetCells(String requestId, String sheetName, List<String> addresses)
      implements Introspection {
    public GetCells {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      addresses = copyAddresses(addresses);
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at the provided top-left address. */
  record GetWindow(
      String requestId, String sheetName, String topLeftAddress, int rowCount, int columnCount)
      implements Introspection {
    public GetWindow {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      requirePositive(rowCount, "rowCount");
      requirePositive(columnCount, "columnCount");
    }
  }

  /** Returns the exact merged regions defined on one sheet. */
  record GetMergedRegions(String requestId, String sheetName) implements Introspection {
    public GetMergedRegions {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record GetHyperlinks(String requestId, String sheetName, CellSelection selection)
      implements Introspection {
    public GetHyperlinks {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record GetComments(String requestId, String sheetName, CellSelection selection)
      implements Introspection {
    public GetComments {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns layout metadata such as freeze panes and visible row and column sizing. */
  record GetSheetLayout(String requestId, String sheetName) implements Introspection {
    public GetSheetLayout {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Groups formula usage patterns across one or more sheets. */
  record AnalyzeFormulaSurface(String requestId, SheetSelection selection) implements Insight {
    public AnalyzeFormulaSurface {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Infers a simple column schema from a rectangular window on one sheet. */
  record AnalyzeSheetSchema(
      String requestId, String sheetName, String topLeftAddress, int rowCount, int columnCount)
      implements Insight {
    public AnalyzeSheetSchema {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      requirePositive(rowCount, "rowCount");
      requirePositive(columnCount, "columnCount");
    }
  }

  /** Summarizes the scope and backing kind of selected named ranges. */
  record AnalyzeNamedRangeSurface(String requestId, NamedRangeSelection selection)
      implements Insight {
    public AnalyzeNamedRangeSurface {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static void requirePositive(int value, String fieldName) {
    if (value <= 0) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
  }

  private static List<String> copyAddresses(List<String> addresses) {
    Objects.requireNonNull(addresses, "addresses must not be null");
    List<String> copy = List.copyOf(addresses);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("addresses must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String address : copy) {
      requireNonBlank(address, "addresses");
      if (!unique.add(address)) {
        throw new IllegalArgumentException("addresses must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }
}
