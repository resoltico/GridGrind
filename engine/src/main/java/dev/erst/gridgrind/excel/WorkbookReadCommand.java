package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Workbook-core read commands executed after mutations and before persistence. */
public sealed interface WorkbookReadCommand
    permits WorkbookReadCommand.Introspection, WorkbookReadCommand.Analysis {

  /** Stable caller-provided identifier used to correlate one read command with its result. */
  String stepId();

  /** Marker for fact-only workbook reads. */
  sealed interface Introspection extends WorkbookReadCommand
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

  /** Marker for derived analysis commands. */
  sealed interface Analysis extends WorkbookReadCommand
      permits AnalyzeFormulaHealth,
          AnalyzeDataValidationHealth,
          AnalyzeConditionalFormattingHealth,
          AnalyzeAutofilterHealth,
          AnalyzeTableHealth,
          AnalyzePivotTableHealth,
          AnalyzeHyperlinkHealth,
          AnalyzeNamedRangeHealth,
          AnalyzeWorkbookFindings {}

  /** Returns workbook-level summary facts. */
  record GetWorkbookSummary(String stepId) implements Introspection {
    public GetWorkbookSummary {
      stepId = requireNonBlank(stepId, "stepId");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts for the current workbook. */
  record GetPackageSecurity(String stepId) implements Introspection {
    public GetPackageSecurity {
      stepId = requireNonBlank(stepId, "stepId");
    }
  }

  /** Returns workbook-level protection facts. */
  record GetWorkbookProtection(String stepId) implements Introspection {
    public GetWorkbookProtection {
      stepId = requireNonBlank(stepId, "stepId");
    }
  }

  /** Returns named ranges selected by exact selector or workbook-wide selection. */
  record GetNamedRanges(String stepId, ExcelNamedRangeSelection selection)
      implements Introspection {
    public GetNamedRanges {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns structural summary facts for one sheet. */
  record GetSheetSummary(String stepId, String sheetName) implements Introspection {
    public GetSheetSummary {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns exact cell snapshots for the provided ordered A1 addresses. */
  record GetCells(String stepId, String sheetName, List<String> addresses)
      implements Introspection {
    public GetCells {
      stepId = requireNonBlank(stepId, "stepId");
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
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount)
      implements Introspection {
    public GetWindow {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      requirePositive(rowCount, "rowCount");
      requirePositive(columnCount, "columnCount");
      requireWindowSize(rowCount, columnCount);
    }
  }

  /** Returns the exact merged regions defined on one sheet. */
  record GetMergedRegions(String stepId, String sheetName) implements Introspection {
    public GetMergedRegions {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record GetHyperlinks(String stepId, String sheetName, ExcelCellSelection selection)
      implements Introspection {
    public GetHyperlinks {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record GetComments(String stepId, String sheetName, ExcelCellSelection selection)
      implements Introspection {
    public GetComments {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns factual drawing-object metadata for one sheet. */
  record GetDrawingObjects(String stepId, String sheetName) implements Introspection {
    public GetDrawingObjects {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  record GetCharts(String stepId, String sheetName) implements Introspection {
    public GetCharts {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object on one sheet. */
  record GetDrawingObjectPayload(String stepId, String sheetName, String objectName)
      implements Introspection {
    public GetDrawingObjectPayload {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      objectName = requireNonBlank(objectName, "objectName");
    }
  }

  /** Returns layout metadata such as pane state, zoom, and row and column sizing. */
  record GetSheetLayout(String stepId, String sheetName) implements Introspection {
    public GetSheetLayout {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns supported print-layout metadata for one sheet. */
  record GetPrintLayout(String stepId, String sheetName) implements Introspection {
    public GetPrintLayout {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record GetDataValidations(String stepId, String sheetName, ExcelRangeSelection selection)
      implements Introspection {
    public GetDataValidations {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record GetConditionalFormatting(String stepId, String sheetName, ExcelRangeSelection selection)
      implements Introspection {
    public GetConditionalFormatting {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record GetAutofilters(String stepId, String sheetName) implements Introspection {
    public GetAutofilters {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record GetTables(String stepId, ExcelTableSelection selection) implements Introspection {
    public GetTables {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  record GetPivotTables(String stepId, ExcelPivotTableSelection selection)
      implements Introspection {
    public GetPivotTables {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Groups formula usage patterns across one or more sheets. */
  record GetFormulaSurface(String stepId, ExcelSheetSelection selection) implements Introspection {
    public GetFormulaSurface {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Infers a simple column schema from a rectangular window on one sheet. */
  record GetSheetSchema(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount)
      implements Introspection {
    public GetSheetSchema {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      requirePositive(rowCount, "rowCount");
      requirePositive(columnCount, "columnCount");
      requireWindowSize(rowCount, columnCount);
    }
  }

  /** Summarizes the scope and backing kind of selected named ranges. */
  record GetNamedRangeSurface(String stepId, ExcelNamedRangeSelection selection)
      implements Introspection {
    public GetNamedRangeSurface {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports formula findings such as error results and volatile usage. */
  record AnalyzeFormulaHealth(String stepId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeFormulaHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports data-validation findings such as malformed or overlapping rules. */
  record AnalyzeDataValidationHealth(String stepId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeDataValidationHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports conditional-formatting findings such as broken formulas or priority collisions. */
  record AnalyzeConditionalFormattingHealth(String stepId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeConditionalFormattingHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports autofilter findings such as invalid ranges or table-ownership mismatches. */
  record AnalyzeAutofilterHealth(String stepId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeAutofilterHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports table findings such as overlaps, broken ranges, or invalid headers. */
  record AnalyzeTableHealth(String stepId, ExcelTableSelection selection) implements Analysis {
    public AnalyzeTableHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports pivot-table findings such as broken caches, names, and sources. */
  record AnalyzePivotTableHealth(String stepId, ExcelPivotTableSelection selection)
      implements Analysis {
    public AnalyzePivotTableHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports hyperlink findings such as malformed targets and missing document destinations. */
  record AnalyzeHyperlinkHealth(String stepId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeHyperlinkHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports named-range findings such as broken references and scope shadowing. */
  record AnalyzeNamedRangeHealth(String stepId, ExcelNamedRangeSelection selection)
      implements Analysis {
    public AnalyzeNamedRangeHealth {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Runs the first analysis family across the whole workbook and aggregates their findings. */
  record AnalyzeWorkbookFindings(String stepId) implements Analysis {
    public AnalyzeWorkbookFindings {
      stepId = requireNonBlank(stepId, "stepId");
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
    return copy;
  }
}
