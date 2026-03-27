package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Immutable workbook-core result produced by one read command. */
public sealed interface WorkbookReadResult
    permits WorkbookReadResult.Introspection, WorkbookReadResult.Analysis {

  /** Stable caller-provided identifier copied from the originating read command. */
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
          FormulaSurfaceResult,
          SheetSchemaResult,
          NamedRangeSurfaceResult {}

  /** Marker for derived workbook analysis results. */
  sealed interface Analysis extends WorkbookReadResult
      permits FormulaHealthResult,
          HyperlinkHealthResult,
          NamedRangeHealthResult,
          WorkbookFindingsResult {}

  /** Returns workbook-level summary facts. */
  record WorkbookSummaryResult(String requestId, WorkbookSummary workbook)
      implements Introspection {
    public WorkbookSummaryResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns selected named ranges. */
  record NamedRangesResult(String requestId, List<ExcelNamedRangeSnapshot> namedRanges)
      implements Introspection {
    public NamedRangesResult {
      requestId = requireNonBlank(requestId, "requestId");
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }

  /** Returns summary facts for one sheet. */
  record SheetSummaryResult(String requestId, SheetSummary sheet) implements Introspection {
    public SheetSummaryResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  record CellsResult(String requestId, String sheetName, List<ExcelCellSnapshot> cells)
      implements Introspection {
    public CellsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      cells = copyValues(cells, "cells");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  record WindowResult(String requestId, Window window) implements Introspection {
    public WindowResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(window, "window must not be null");
    }
  }

  /** Returns every merged region defined on one sheet. */
  record MergedRegionsResult(String requestId, String sheetName, List<MergedRegion> mergedRegions)
      implements Introspection {
    public MergedRegionsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      mergedRegions = copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record HyperlinksResult(String requestId, String sheetName, List<CellHyperlink> hyperlinks)
      implements Introspection {
    public HyperlinksResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      hyperlinks = copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record CommentsResult(String requestId, String sheetName, List<CellComment> comments)
      implements Introspection {
    public CommentsResult {
      requestId = requireNonBlank(requestId, "requestId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      comments = copyValues(comments, "comments");
    }
  }

  /** Returns layout metadata such as freeze panes and visible sizing. */
  record SheetLayoutResult(String requestId, SheetLayout layout) implements Introspection {
    public SheetLayoutResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceResult(String requestId, FormulaSurface analysis) implements Introspection {
    public FormulaSurfaceResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns inferred schema facts for one rectangular sheet window. */
  record SheetSchemaResult(String requestId, SheetSchema analysis) implements Introspection {
    public SheetSchemaResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns high-level characterization of selected named ranges. */
  record NamedRangeSurfaceResult(String requestId, NamedRangeSurface analysis)
      implements Introspection {
    public NamedRangeSurfaceResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns formula-health findings for one analysis read. */
  record FormulaHealthResult(String requestId, WorkbookAnalysis.FormulaHealth analysis)
      implements Analysis {
    public FormulaHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns hyperlink-health findings for one analysis read. */
  record HyperlinkHealthResult(String requestId, WorkbookAnalysis.HyperlinkHealth analysis)
      implements Analysis {
    public HyperlinkHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns named-range-health findings for one analysis read. */
  record NamedRangeHealthResult(String requestId, WorkbookAnalysis.NamedRangeHealth analysis)
      implements Analysis {
    public NamedRangeHealthResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns the aggregated workbook findings from the first analysis family. */
  record WorkbookFindingsResult(String requestId, WorkbookAnalysis.WorkbookFindings analysis)
      implements Analysis {
    public WorkbookFindingsResult {
      requestId = requireNonBlank(requestId, "requestId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Workbook-level summary facts captured after all mutations complete. */
  record WorkbookSummary(
      int sheetCount,
      List<String> sheetNames,
      int namedRangeCount,
      boolean forceFormulaRecalculationOnOpen) {
    public WorkbookSummary {
      if (sheetCount < 0) {
        throw new IllegalArgumentException("sheetCount must not be negative");
      }
      if (namedRangeCount < 0) {
        throw new IllegalArgumentException("namedRangeCount must not be negative");
      }
      sheetNames = copyStrings(sheetNames, "sheetNames");
    }
  }

  /** Structural summary facts for one sheet. */
  record SheetSummary(
      String sheetName, int physicalRowCount, int lastRowIndex, int lastColumnIndex) {
    public SheetSummary {
      sheetName = requireNonBlank(sheetName, "sheetName");
      if (physicalRowCount < 0) {
        throw new IllegalArgumentException("physicalRowCount must not be negative");
      }
      if (lastRowIndex < -1) {
        throw new IllegalArgumentException("lastRowIndex must be greater than or equal to -1");
      }
      if (lastColumnIndex < -1) {
        throw new IllegalArgumentException("lastColumnIndex must be greater than or equal to -1");
      }
    }
  }

  /** Rectangular window of cell snapshots anchored at one top-left address. */
  record Window(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      List<WindowRow> rows) {
    public Window {
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
      rows = copyValues(rows, "rows");
    }
  }

  /** One row inside a rectangular window of cell snapshots. */
  record WindowRow(int rowIndex, List<ExcelCellSnapshot> cells) {
    public WindowRow {
      if (rowIndex < 0) {
        throw new IllegalArgumentException("rowIndex must not be negative");
      }
      cells = copyValues(cells, "cells");
    }
  }

  /** One merged region captured from a sheet. */
  record MergedRegion(String range) {
    public MergedRegion {
      range = requireNonBlank(range, "range");
    }
  }

  /** Hyperlink metadata associated with one concrete cell address. */
  record CellHyperlink(String address, ExcelHyperlink hyperlink) {
    public CellHyperlink {
      address = requireNonBlank(address, "address");
      Objects.requireNonNull(hyperlink, "hyperlink must not be null");
    }
  }

  /** Comment metadata associated with one concrete cell address. */
  record CellComment(String address, ExcelComment comment) {
    public CellComment {
      address = requireNonBlank(address, "address");
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Layout metadata such as freeze panes and visible row and column sizing for one sheet. */
  record SheetLayout(
      String sheetName, FreezePane freezePanes, List<ColumnLayout> columns, List<RowLayout> rows) {
    public SheetLayout {
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(freezePanes, "freezePanes must not be null");
      columns = copyValues(columns, "columns");
      rows = copyValues(rows, "rows");
    }
  }

  /** Freeze-pane state captured from a sheet layout read. */
  sealed interface FreezePane permits FreezePane.None, FreezePane.Frozen {

    /** Sheet has no active freeze panes. */
    record None() implements FreezePane {}

    /** Sheet is frozen at the provided split and visible-origin coordinates. */
    record Frozen(int splitColumn, int splitRow, int leftmostColumn, int topRow)
        implements FreezePane {
      public Frozen {
        if (splitColumn < 0) {
          throw new IllegalArgumentException("splitColumn must not be negative");
        }
        if (splitRow < 0) {
          throw new IllegalArgumentException("splitRow must not be negative");
        }
        if (leftmostColumn < 0) {
          throw new IllegalArgumentException("leftmostColumn must not be negative");
        }
        if (topRow < 0) {
          throw new IllegalArgumentException("topRow must not be negative");
        }
      }
    }
  }

  /** Width metadata for one sheet column. */
  record ColumnLayout(int columnIndex, double widthCharacters) {
    public ColumnLayout {
      if (columnIndex < 0) {
        throw new IllegalArgumentException("columnIndex must not be negative");
      }
      if (!Double.isFinite(widthCharacters) || widthCharacters <= 0.0d) {
        throw new IllegalArgumentException("widthCharacters must be finite and greater than 0");
      }
    }
  }

  /** Height metadata for one sheet row. */
  record RowLayout(int rowIndex, double heightPoints) {
    public RowLayout {
      if (rowIndex < 0) {
        throw new IllegalArgumentException("rowIndex must not be negative");
      }
      if (!Double.isFinite(heightPoints) || heightPoints <= 0.0d) {
        throw new IllegalArgumentException("heightPoints must be finite and greater than 0");
      }
    }
  }

  /** Grouped formula usage facts across one or more sheets. */
  record FormulaSurface(int totalFormulaCellCount, List<SheetFormulaSurface> sheets) {
    public FormulaSurface {
      if (totalFormulaCellCount < 0) {
        throw new IllegalArgumentException("totalFormulaCellCount must not be negative");
      }
      sheets = copyValues(sheets, "sheets");
    }
  }

  /** Formula usage facts for one sheet. */
  record SheetFormulaSurface(
      String sheetName,
      int formulaCellCount,
      int distinctFormulaCount,
      List<FormulaPattern> formulas) {
    public SheetFormulaSurface {
      sheetName = requireNonBlank(sheetName, "sheetName");
      if (formulaCellCount < 0) {
        throw new IllegalArgumentException("formulaCellCount must not be negative");
      }
      if (distinctFormulaCount < 0) {
        throw new IllegalArgumentException("distinctFormulaCount must not be negative");
      }
      formulas = copyValues(formulas, "formulas");
    }
  }

  /** One grouped formula pattern and the addresses where it appears. */
  record FormulaPattern(String formula, int occurrenceCount, List<String> addresses) {
    public FormulaPattern {
      formula = requireNonBlank(formula, "formula");
      if (occurrenceCount <= 0) {
        throw new IllegalArgumentException("occurrenceCount must be greater than 0");
      }
      addresses = copyStrings(addresses, "addresses");
    }
  }

  /** Inferred schema facts for one rectangular sheet window. */
  record SheetSchema(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      int dataRowCount,
      List<SchemaColumn> columns) {
    public SheetSchema {
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
      if (dataRowCount < 0) {
        throw new IllegalArgumentException("dataRowCount must not be negative");
      }
      columns = copyValues(columns, "columns");
    }
  }

  /** One inferred schema column with header text and observed value-type counts. */
  record SchemaColumn(
      int columnIndex,
      String columnAddress,
      String headerDisplayValue,
      int populatedCellCount,
      int blankCellCount,
      List<TypeCount> observedTypes,
      String dominantType) {
    public SchemaColumn {
      if (columnIndex < 0) {
        throw new IllegalArgumentException("columnIndex must not be negative");
      }
      columnAddress = requireNonBlank(columnAddress, "columnAddress");
      Objects.requireNonNull(headerDisplayValue, "headerDisplayValue must not be null");
      if (populatedCellCount < 0) {
        throw new IllegalArgumentException("populatedCellCount must not be negative");
      }
      if (blankCellCount < 0) {
        throw new IllegalArgumentException("blankCellCount must not be negative");
      }
      observedTypes = copyValues(observedTypes, "observedTypes");
    }
  }

  /** Count of one observed cell type inside a schema column. */
  record TypeCount(String type, int count) {
    public TypeCount {
      type = requireNonBlank(type, "type");
      if (count <= 0) {
        throw new IllegalArgumentException("count must be greater than 0");
      }
    }
  }

  /** High-level characterization of selected named ranges. */
  record NamedRangeSurface(
      int workbookScopedCount,
      int sheetScopedCount,
      int rangeBackedCount,
      int formulaBackedCount,
      List<NamedRangeSurfaceEntry> namedRanges) {
    public NamedRangeSurface {
      if (workbookScopedCount < 0) {
        throw new IllegalArgumentException("workbookScopedCount must not be negative");
      }
      if (sheetScopedCount < 0) {
        throw new IllegalArgumentException("sheetScopedCount must not be negative");
      }
      if (rangeBackedCount < 0) {
        throw new IllegalArgumentException("rangeBackedCount must not be negative");
      }
      if (formulaBackedCount < 0) {
        throw new IllegalArgumentException("formulaBackedCount must not be negative");
      }
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }

  /** One named-range surface entry classified by scope and backing kind. */
  record NamedRangeSurfaceEntry(
      String name, ExcelNamedRangeScope scope, String refersToFormula, NamedRangeBackingKind kind) {
    public NamedRangeSurfaceEntry {
      name = requireNonBlank(name, "name");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      Objects.requireNonNull(kind, "kind must not be null");
    }
  }

  /** Distinguishes range-backed named ranges from formula-backed named ranges. */
  enum NamedRangeBackingKind {
    RANGE,
    FORMULA
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
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
