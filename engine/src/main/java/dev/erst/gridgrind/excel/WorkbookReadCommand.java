package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Workbook-core read commands executed after mutations and before persistence. */
public sealed interface WorkbookReadCommand
    permits WorkbookReadCommand.Introspection, WorkbookReadCommand.Analysis {

  /** Stable caller-provided identifier used to correlate one read command with its result. */
  String requestId();

  /** Marker for fact-only workbook reads. */
  sealed interface Introspection extends WorkbookReadCommand
      permits GetWorkbookSummary,
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

  /** Marker for derived analysis commands. */
  sealed interface Analysis extends WorkbookReadCommand
      permits AnalyzeFormulaHealth,
          AnalyzeDataValidationHealth,
          AnalyzeConditionalFormattingHealth,
          AnalyzeAutofilterHealth,
          AnalyzeTableHealth,
          AnalyzeHyperlinkHealth,
          AnalyzeNamedRangeHealth,
          AnalyzeWorkbookFindings {}

  /** Returns workbook-level summary facts. */
  record GetWorkbookSummary(String requestId) implements Introspection {
    public GetWorkbookSummary {
      requestId = requireNonBlank(requestId, "requestId");
    }
  }

  /** Returns workbook-level protection facts. */
  record GetWorkbookProtection(String requestId) implements Introspection {
    public GetWorkbookProtection {
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

  /**
   * Maximum number of cells ({@code rowCount * columnCount}) permitted in a single window command.
   * Enforced in both the protocol and engine layers to prevent out-of-memory failures during
   * cell-grid serialization. See docs/LIMITATIONS.md LIM-001.
   */
  int MAX_WINDOW_CELLS = 250_000; // LIM-001

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
      requireWindowSize(rowCount, columnCount);
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

  /** Returns factual drawing-object metadata for one sheet. */
  record GetDrawingObjects(String requestId, String sheetName) implements Introspection {
    public GetDrawingObjects {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  record GetCharts(String requestId, String sheetName) implements Introspection {
    public GetCharts {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object on one sheet. */
  record GetDrawingObjectPayload(String requestId, String sheetName, String objectName)
      implements Introspection {
    public GetDrawingObjectPayload {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      objectName = requireNonBlank(objectName, "objectName");
    }
  }

  /** Returns layout metadata such as pane state, zoom, and row and column sizing. */
  record GetSheetLayout(String requestId, String sheetName) implements Introspection {
    public GetSheetLayout {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns supported print-layout metadata for one sheet. */
  record GetPrintLayout(String requestId, String sheetName) implements Introspection {
    public GetPrintLayout {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record GetDataValidations(String requestId, String sheetName, ExcelRangeSelection selection)
      implements Introspection {
    public GetDataValidations {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record GetConditionalFormatting(String requestId, String sheetName, ExcelRangeSelection selection)
      implements Introspection {
    public GetConditionalFormatting {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record GetAutofilters(String requestId, String sheetName) implements Introspection {
    public GetAutofilters {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record GetTables(String requestId, ExcelTableSelection selection) implements Introspection {
    public GetTables {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Groups formula usage patterns across one or more sheets. */
  record GetFormulaSurface(String requestId, ExcelSheetSelection selection)
      implements Introspection {
    public GetFormulaSurface {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Infers a simple column schema from a rectangular window on one sheet. */
  record GetSheetSchema(
      String requestId, String sheetName, String topLeftAddress, int rowCount, int columnCount)
      implements Introspection {
    public GetSheetSchema {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      requirePositive(rowCount, "rowCount");
      requirePositive(columnCount, "columnCount");
      requireWindowSize(rowCount, columnCount);
    }
  }

  /** Summarizes the scope and backing kind of selected named ranges. */
  record GetNamedRangeSurface(String requestId, ExcelNamedRangeSelection selection)
      implements Introspection {
    public GetNamedRangeSurface {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports formula findings such as error results and volatile usage. */
  record AnalyzeFormulaHealth(String requestId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeFormulaHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports data-validation findings such as malformed or overlapping rules. */
  record AnalyzeDataValidationHealth(String requestId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeDataValidationHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports conditional-formatting findings such as broken formulas or priority collisions. */
  record AnalyzeConditionalFormattingHealth(String requestId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeConditionalFormattingHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports autofilter findings such as invalid ranges or table-ownership mismatches. */
  record AnalyzeAutofilterHealth(String requestId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeAutofilterHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports table findings such as overlaps, broken ranges, or invalid headers. */
  record AnalyzeTableHealth(String requestId, ExcelTableSelection selection) implements Analysis {
    public AnalyzeTableHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports hyperlink findings such as malformed targets and missing document destinations. */
  record AnalyzeHyperlinkHealth(String requestId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeHyperlinkHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports named-range findings such as broken references and scope shadowing. */
  record AnalyzeNamedRangeHealth(String requestId, ExcelNamedRangeSelection selection)
      implements Analysis {
    public AnalyzeNamedRangeHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Runs the first analysis family across the whole workbook and aggregates their findings. */
  record AnalyzeWorkbookFindings(String requestId) implements Analysis {
    public AnalyzeWorkbookFindings {
      requestId = requireNonBlank(requestId, "requestId");
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

  private static void requireWindowSize(int rowCount, int columnCount) {
    long cells = (long) rowCount * columnCount;
    if (cells > MAX_WINDOW_CELLS) {
      throw new IllegalArgumentException(
          "rowCount * columnCount must not exceed " + MAX_WINDOW_CELLS + " but was " + cells);
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
