package dev.erst.gridgrind.protocol.read;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.protocol.dto.*;
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
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetWorkbookProtection.class,
      name = "GET_WORKBOOK_PROTECTION"),
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
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetPrintLayout.class, name = "GET_PRINT_LAYOUT"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetDataValidations.class,
      name = "GET_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetConditionalFormatting.class,
      name = "GET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetAutofilters.class, name = "GET_AUTOFILTERS"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetTables.class, name = "GET_TABLES"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetFormulaSurface.class,
      name = "GET_FORMULA_SURFACE"),
  @JsonSubTypes.Type(value = WorkbookReadOperation.GetSheetSchema.class, name = "GET_SHEET_SCHEMA"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.GetNamedRangeSurface.class,
      name = "GET_NAMED_RANGE_SURFACE"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeFormulaHealth.class,
      name = "ANALYZE_FORMULA_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeDataValidationHealth.class,
      name = "ANALYZE_DATA_VALIDATION_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeConditionalFormattingHealth.class,
      name = "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeAutofilterHealth.class,
      name = "ANALYZE_AUTOFILTER_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeTableHealth.class,
      name = "ANALYZE_TABLE_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeHyperlinkHealth.class,
      name = "ANALYZE_HYPERLINK_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeNamedRangeHealth.class,
      name = "ANALYZE_NAMED_RANGE_HEALTH"),
  @JsonSubTypes.Type(
      value = WorkbookReadOperation.AnalyzeWorkbookFindings.class,
      name = "ANALYZE_WORKBOOK_FINDINGS")
})
public sealed interface WorkbookReadOperation
    permits WorkbookReadOperation.Introspection, WorkbookReadOperation.Analysis {

  /** Stable caller-provided identifier used to correlate this read with its result. */
  String requestId();

  /** Marker for raw workbook-fact reads with no higher-level interpretation. */
  sealed interface Introspection extends WorkbookReadOperation
      permits GetWorkbookSummary,
          GetWorkbookProtection,
          GetNamedRanges,
          GetSheetSummary,
          GetCells,
          GetWindow,
          GetMergedRegions,
          GetHyperlinks,
          GetComments,
          GetSheetLayout,
          GetPrintLayout,
          GetDataValidations,
          GetConditionalFormatting,
          GetAutofilters,
          GetTables,
          GetFormulaSurface,
          GetSheetSchema,
          GetNamedRangeSurface {}

  /** Marker for derived workbook analysis built on top of introspection primitives. */
  sealed interface Analysis extends WorkbookReadOperation
      permits AnalyzeFormulaHealth,
          AnalyzeDataValidationHealth,
          AnalyzeConditionalFormattingHealth,
          AnalyzeAutofilterHealth,
          AnalyzeTableHealth,
          AnalyzeHyperlinkHealth,
          AnalyzeNamedRangeHealth,
          AnalyzeWorkbookFindings {}

  /** Returns workbook-level summary facts such as sheet order and recalculation flag. */
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

  /**
   * Maximum number of cells ({@code rowCount * columnCount}) permitted in a single window request.
   * Requests exceeding this limit are rejected at parse time to prevent out-of-memory failures
   * during serialization of large cell grids. See docs/LIMITATIONS.md LIM-001.
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

  /** Returns layout metadata such as pane state, zoom, and visible row and column sizing. */
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
  record GetDataValidations(String requestId, String sheetName, RangeSelection selection)
      implements Introspection {
    public GetDataValidations {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record GetConditionalFormatting(String requestId, String sheetName, RangeSelection selection)
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
  record GetTables(String requestId, TableSelection selection) implements Introspection {
    public GetTables {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Groups formula usage patterns across one or more sheets. */
  record GetFormulaSurface(String requestId, SheetSelection selection) implements Introspection {
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
  record GetNamedRangeSurface(String requestId, NamedRangeSelection selection)
      implements Introspection {
    public GetNamedRangeSurface {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports formula findings such as error results and volatile usage. */
  record AnalyzeFormulaHealth(String requestId, SheetSelection selection) implements Analysis {
    public AnalyzeFormulaHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports data-validation findings such as malformed or overlapping rules. */
  record AnalyzeDataValidationHealth(String requestId, SheetSelection selection)
      implements Analysis {
    public AnalyzeDataValidationHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports conditional-formatting findings such as broken formulas or priority collisions. */
  record AnalyzeConditionalFormattingHealth(String requestId, SheetSelection selection)
      implements Analysis {
    public AnalyzeConditionalFormattingHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports autofilter findings such as invalid ranges or table-ownership mismatches. */
  record AnalyzeAutofilterHealth(String requestId, SheetSelection selection) implements Analysis {
    public AnalyzeAutofilterHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports table findings such as overlaps, broken ranges, or invalid headers. */
  record AnalyzeTableHealth(String requestId, TableSelection selection) implements Analysis {
    public AnalyzeTableHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports hyperlink findings such as malformed targets and missing document destinations. */
  record AnalyzeHyperlinkHealth(String requestId, SheetSelection selection) implements Analysis {
    public AnalyzeHyperlinkHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports named-range findings such as broken references and scope shadowing. */
  record AnalyzeNamedRangeHealth(String requestId, NamedRangeSelection selection)
      implements Analysis {
    public AnalyzeNamedRangeHealth {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Runs the first analysis family across the workbook and aggregates their findings. */
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

  private static List<String> copyAddresses(List<String> addresses) { // LIM-007
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
