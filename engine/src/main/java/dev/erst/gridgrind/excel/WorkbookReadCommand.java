package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Workbook-core read commands executed after mutations and before persistence. */
public sealed interface WorkbookReadCommand
    permits WorkbookReadCommand.Introspection, WorkbookReadCommand.Insight {

  /** Stable caller-provided identifier used to correlate one read command with its result. */
  String requestId();

  /** Marker for fact-only workbook reads. */
  sealed interface Introspection extends WorkbookReadCommand
      permits GetWorkbookSummary,
          GetNamedRanges,
          GetSheetSummary,
          GetCells,
          GetWindow,
          GetMergedRegions,
          GetHyperlinks,
          GetComments,
          GetSheetLayout {}

  /** Marker for derived analysis commands. */
  sealed interface Insight extends WorkbookReadCommand
      permits AnalyzeFormulaSurface, AnalyzeSheetSchema, AnalyzeNamedRangeSurface {}

  /** Returns workbook-level summary facts. */
  record GetWorkbookSummary(String requestId) implements Introspection {
    public GetWorkbookSummary {
      requestId = requireNonBlank(requestId, "requestId");
    }
  }

  /** Returns named ranges selected by exact selector or workbook-wide selection. */
  record GetNamedRanges(String requestId, ExcelNamedRangeSelection selection)
      implements Introspection {
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
  record GetHyperlinks(String requestId, String sheetName, ExcelCellSelection selection)
      implements Introspection {
    public GetHyperlinks {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record GetComments(String requestId, String sheetName, ExcelCellSelection selection)
      implements Introspection {
    public GetComments {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns layout metadata such as freeze panes and row and column sizing. */
  record GetSheetLayout(String requestId, String sheetName) implements Introspection {
    public GetSheetLayout {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Groups formula usage patterns across one or more sheets. */
  record AnalyzeFormulaSurface(String requestId, ExcelSheetSelection selection) implements Insight {
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
  record AnalyzeNamedRangeSurface(String requestId, ExcelNamedRangeSelection selection)
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
