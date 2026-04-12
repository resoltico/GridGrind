package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Structured protocol response for both successful workbook workflows and deterministic failures.
 */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
@JsonSubTypes({
  @JsonSubTypes.Type(value = GridGrindResponse.Success.class, name = "SUCCESS"),
  @JsonSubTypes.Type(value = GridGrindResponse.Failure.class, name = "ERROR")
})
public sealed interface GridGrindResponse {
  /** Protocol version negotiated for this response. */
  GridGrindProtocolVersion protocolVersion();

  /** Successful workbook execution result. */
  record Success(
      GridGrindProtocolVersion protocolVersion,
      PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<WorkbookReadResult> reads)
      implements GridGrindResponse {
    public Success {
      protocolVersion =
          protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
      persistence = persistence == null ? new PersistenceOutcome.NotSaved() : persistence;
      warnings = warnings == null ? List.of() : copyValues(warnings, "warnings");
      reads = copyValues(reads, "reads");
    }
  }

  /** Failed workbook execution with a structured problem. */
  record Failure(GridGrindProtocolVersion protocolVersion, Problem problem)
      implements GridGrindResponse {
    public Failure {
      protocolVersion =
          protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
      Objects.requireNonNull(problem, "problem must not be null");
    }
  }

  /** Reports whether the workbook was persisted during successful execution. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = PersistenceOutcome.NotSaved.class, name = "NONE"),
    @JsonSubTypes.Type(value = PersistenceOutcome.SavedAs.class, name = "SAVE_AS"),
    @JsonSubTypes.Type(value = PersistenceOutcome.Overwritten.class, name = "OVERWRITE")
  })
  sealed interface PersistenceOutcome
      permits PersistenceOutcome.NotSaved,
          PersistenceOutcome.SavedAs,
          PersistenceOutcome.Overwritten {

    /** Workbook remained in memory only and was not written to disk. */
    record NotSaved() implements PersistenceOutcome {}

    /**
     * Workbook was written to the path supplied in the SAVE_AS persistence field.
     *
     * <p>{@code requestedPath} is the literal string from the request. {@code executionPath} is the
     * absolute normalized path where the file was actually written. They differ when the request
     * supplies a relative path (e.g. {@code "report.xlsx"}) or a path with {@code ..} segments.
     */
    record SavedAs(String requestedPath, String executionPath) implements PersistenceOutcome {
      public SavedAs {
        Objects.requireNonNull(requestedPath, "requestedPath must not be null");
        Objects.requireNonNull(executionPath, "executionPath must not be null");
        if (requestedPath.isBlank()) {
          throw new IllegalArgumentException("requestedPath must not be blank");
        }
        if (executionPath.isBlank()) {
          throw new IllegalArgumentException("executionPath must not be blank");
        }
      }
    }

    /**
     * Workbook was saved by overwriting the opened source workbook path.
     *
     * <p>{@code sourcePath} is the path string from the EXISTING source as supplied in the request.
     * {@code executionPath} is the absolute normalized path where the file was written.
     */
    record Overwritten(String sourcePath, String executionPath) implements PersistenceOutcome {
      public Overwritten {
        Objects.requireNonNull(sourcePath, "sourcePath must not be null");
        Objects.requireNonNull(executionPath, "executionPath must not be null");
        if (sourcePath.isBlank()) {
          throw new IllegalArgumentException("sourcePath must not be blank");
        }
        if (executionPath.isBlank()) {
          throw new IllegalArgumentException("executionPath must not be blank");
        }
      }
    }
  }

  /** High-level workbook facts returned on success. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookSummary.Empty.class, name = "EMPTY"),
    @JsonSubTypes.Type(value = WorkbookSummary.WithSheets.class, name = "WITH_SHEETS")
  })
  sealed interface WorkbookSummary permits WorkbookSummary.Empty, WorkbookSummary.WithSheets {
    /** Total sheet count after all mutations complete. */
    int sheetCount();

    /** Ordered workbook sheet names. */
    List<String> sheetNames();

    /** Count of exposed named ranges after all mutations complete. */
    int namedRangeCount();

    /** Whether the workbook is marked to recalculate formulas on open. */
    boolean forceFormulaRecalculationOnOpen();

    /** Workbook summary for a zero-sheet workbook. */
    record Empty(
        int sheetCount,
        List<String> sheetNames,
        int namedRangeCount,
        boolean forceFormulaRecalculationOnOpen)
        implements WorkbookSummary {
      public Empty {
        sheetNames = validateCommonWorkbookSummaryFields(sheetCount, sheetNames, namedRangeCount);
        if (sheetCount != 0) {
          throw new IllegalArgumentException("sheetCount must be 0 for an empty workbook");
        }
      }
    }

    /** Workbook summary for a workbook that contains one or more sheets. */
    record WithSheets(
        int sheetCount,
        List<String> sheetNames,
        String activeSheetName,
        List<String> selectedSheetNames,
        int namedRangeCount,
        boolean forceFormulaRecalculationOnOpen)
        implements WorkbookSummary {
      public WithSheets {
        sheetNames = validateCommonWorkbookSummaryFields(sheetCount, sheetNames, namedRangeCount);
        Objects.requireNonNull(activeSheetName, "activeSheetName must not be null");
        if (activeSheetName.isBlank()) {
          throw new IllegalArgumentException("activeSheetName must not be blank");
        }
        selectedSheetNames = copyDistinctStrings(selectedSheetNames, "selectedSheetNames");
        if (sheetCount == 0) {
          throw new IllegalArgumentException("sheetCount must be greater than 0");
        }
        if (!sheetNames.contains(activeSheetName)) {
          throw new IllegalArgumentException("activeSheetName must be present in sheetNames");
        }
        if (selectedSheetNames.isEmpty()) {
          throw new IllegalArgumentException("selectedSheetNames must not be empty");
        }
        for (String selectedSheetName : selectedSheetNames) {
          if (!sheetNames.contains(selectedSheetName)) {
            throw new IllegalArgumentException(
                "selectedSheetNames must only contain values present in sheetNames");
          }
        }
      }
    }
  }

  /** Structured workbook-level analysis report for one defined name. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = NamedRangeReport.RangeReport.class, name = "RANGE"),
    @JsonSubTypes.Type(value = NamedRangeReport.FormulaReport.class, name = "FORMULA")
  })
  sealed interface NamedRangeReport {
    /** Defined-name identifier. */
    String name();

    /** Workbook or sheet scope of the defined name. */
    NamedRangeScope scope();

    /** Exact formula text stored in the workbook for this defined name. */
    String refersToFormula();

    /**
     * Resolved typed target; populated only by RangeReport, null for FormulaReport.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default NamedRangeTarget target() {
      return null;
    }

    /** Named range that resolves cleanly to a sheet-qualified cell or rectangular range target. */
    record RangeReport(
        String name, NamedRangeScope scope, String refersToFormula, NamedRangeTarget target)
        implements NamedRangeReport {
      public RangeReport {
        Objects.requireNonNull(name, "name must not be null");
        Objects.requireNonNull(scope, "scope must not be null");
        Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
        Objects.requireNonNull(target, "target must not be null");
      }
    }

    /** Defined name whose formula cannot be normalized to a typed range target. */
    record FormulaReport(String name, NamedRangeScope scope, String refersToFormula)
        implements NamedRangeReport {
      public FormulaReport {
        Objects.requireNonNull(name, "name must not be null");
        Objects.requireNonNull(scope, "scope must not be null");
        Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      }
    }
  }

  /** Structural summary facts for one sheet. */
  record SheetSummaryReport(
      String sheetName,
      ExcelSheetVisibility visibility,
      SheetProtectionReport protection,
      int physicalRowCount,
      int lastRowIndex,
      int lastColumnIndex) {
    public SheetSummaryReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      Objects.requireNonNull(visibility, "visibility must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Captures whether a sheet is protected and, if so, with which supported lock flags. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = SheetProtectionReport.Unprotected.class, name = "UNPROTECTED"),
    @JsonSubTypes.Type(value = SheetProtectionReport.Protected.class, name = "PROTECTED")
  })
  sealed interface SheetProtectionReport
      permits SheetProtectionReport.Unprotected, SheetProtectionReport.Protected {

    /**
     * Supported lock flags; populated only by Protected, null for Unprotected.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default SheetProtectionSettings settings() {
      return null;
    }

    /** Sheet protection is disabled. */
    record Unprotected() implements SheetProtectionReport {}

    /** Sheet protection is enabled with the reported supported lock flags. */
    record Protected(SheetProtectionSettings settings) implements SheetProtectionReport {
      public Protected {
        Objects.requireNonNull(settings, "settings must not be null");
      }
    }
  }

  /** Structured view of one requested or previewed cell. */
  @JsonTypeInfo(
      use = JsonTypeInfo.Id.NAME,
      include = JsonTypeInfo.As.EXISTING_PROPERTY,
      property = "effectiveType")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = CellReport.BlankReport.class, name = "BLANK"),
    @JsonSubTypes.Type(value = CellReport.TextReport.class, name = "STRING"),
    @JsonSubTypes.Type(value = CellReport.NumberReport.class, name = "NUMBER"),
    @JsonSubTypes.Type(value = CellReport.BooleanReport.class, name = "BOOLEAN"),
    @JsonSubTypes.Type(value = CellReport.ErrorReport.class, name = "ERROR"),
    @JsonSubTypes.Type(value = CellReport.FormulaReport.class, name = "FORMULA")
  })
  sealed interface CellReport {
    /** Cell address in A1 notation. */
    String address();

    /** POI cell type as declared before evaluation. */
    String declaredType();

    /** Effective cell type after formula evaluation. */
    String effectiveType();

    /** Formatted display string as shown in Excel. */
    String displayValue();

    /** Style snapshot captured for this cell. */
    CellStyleReport style();

    /** Hyperlink metadata; null when the cell has no hyperlink. */
    HyperlinkTarget hyperlink();

    /** Comment metadata; null when the cell has no comment. */
    CommentReport comment();

    /**
     * Formula text; populated only by FormulaReport, null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String formula() {
      return null;
    }

    /**
     * String value; populated only by TextReport, null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String stringValue() {
      return null;
    }

    /**
     * Structured rich-text runs; populated only by TextReport when the workbook stores run-level
     * formatting, null for all other subtypes and scalar plain-string cells.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    @SuppressWarnings("PMD.ReturnEmptyCollectionRatherThanNull")
    default List<RichTextRunReport> richText() {
      return null;
    }

    /**
     * Numeric value; populated only by NumberReport, null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default Double numberValue() {
      return null;
    }

    /**
     * Boolean value; populated only by BooleanReport, null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default Boolean booleanValue() {
      return null;
    }

    /**
     * Error code string; populated only by ErrorReport, null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String errorValue() {
      return null;
    }

    /** CellReport for a cell with no value or formula. */
    record BlankReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
        HyperlinkTarget hyperlink,
        CommentReport comment)
        implements CellReport {
      public BlankReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
      }

      @Override
      @JsonProperty
      public String effectiveType() {
        return "BLANK";
      }
    }

    /** CellReport for a cell containing a plain string value. */
    record TextReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
        HyperlinkTarget hyperlink,
        CommentReport comment,
        String stringValue,
        List<RichTextRunReport> richText)
        implements CellReport {
      public TextReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
        Objects.requireNonNull(stringValue, "stringValue must not be null");
        richText = copyRichTextRuns(richText, "richText");
        if (richText != null && !stringValue.equals(concatenateRichTextRuns(richText))) {
          throw new IllegalArgumentException(
              "richText run text must concatenate to the stringValue");
        }
      }

      @Override
      @JsonProperty
      public String effectiveType() {
        return "STRING";
      }
    }

    /** CellReport for a cell containing a numeric value. */
    record NumberReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
        HyperlinkTarget hyperlink,
        CommentReport comment,
        Double numberValue)
        implements CellReport {
      public NumberReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
      }

      @Override
      @JsonProperty
      public String effectiveType() {
        return "NUMBER";
      }
    }

    /** CellReport for a cell containing a boolean value. */
    record BooleanReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
        HyperlinkTarget hyperlink,
        CommentReport comment,
        Boolean booleanValue)
        implements CellReport {
      public BooleanReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
      }

      @Override
      @JsonProperty
      public String effectiveType() {
        return "BOOLEAN";
      }
    }

    /** CellReport for a cell in an error state (e.g., #DIV/0!, #REF!). */
    record ErrorReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
        HyperlinkTarget hyperlink,
        CommentReport comment,
        String errorValue)
        implements CellReport {
      public ErrorReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
      }

      @Override
      @JsonProperty
      public String effectiveType() {
        return "ERROR";
      }
    }

    /** CellReport for a cell containing a formula, with its evaluated result nested inside. */
    record FormulaReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
        HyperlinkTarget hyperlink,
        CommentReport comment,
        String formula,
        CellReport evaluation)
        implements CellReport {
      public FormulaReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
        Objects.requireNonNull(formula, "formula must not be null");
      }

      @Override
      @JsonProperty
      public String effectiveType() {
        return "FORMULA";
      }
    }
  }

  /** Factual comment metadata returned for analyzed cells that carry a comment. */
  record CommentReport(
      String text,
      String author,
      boolean visible,
      java.util.List<RichTextRunReport> runs,
      CommentAnchorReport anchor) {
    /** Creates a plain comment report without rich runs or anchor metadata. */
    public CommentReport(String text, String author, boolean visible) {
      this(text, author, visible, null, null);
    }

    public CommentReport {
      Objects.requireNonNull(text, "text must not be null");
      Objects.requireNonNull(author, "author must not be null");
      if (runs != null) {
        runs = copyValues(runs, "runs");
        if (!text.equals(
            runs.stream()
                .map(RichTextRunReport::text)
                .collect(java.util.stream.Collectors.joining()))) {
          throw new IllegalArgumentException("comment runs must concatenate to the plain text");
        }
      }
    }
  }

  /** Effective style facts returned with every analyzed cell. */
  record CellStyleReport(
      String numberFormat,
      CellAlignmentReport alignment,
      CellFontReport font,
      CellFillReport fill,
      CellBorderReport border,
      CellProtectionReport protection) {
    public CellStyleReport {
      Objects.requireNonNull(numberFormat, "numberFormat must not be null");
      Objects.requireNonNull(alignment, "alignment must not be null");
      Objects.requireNonNull(font, "font must not be null");
      Objects.requireNonNull(fill, "fill must not be null");
      Objects.requireNonNull(border, "border must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Rectangular window of cells anchored at one top-left address. */
  record WindowReport(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      List<WindowRowReport> rows) {
    public WindowReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
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
  record WindowRowReport(int rowIndex, List<CellReport> cells) {
    public WindowRowReport {
      cells = copyValues(cells, "cells");
    }
  }

  /** One merged region captured from a sheet. */
  record MergedRegionReport(String range) {
    public MergedRegionReport {
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Hyperlink metadata associated with one concrete cell address. */
  record CellHyperlinkReport(String address, HyperlinkTarget hyperlink) {
    public CellHyperlinkReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(hyperlink, "hyperlink must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Comment metadata associated with one concrete cell address. */
  record CellCommentReport(String address, CommentReport comment) {
    public CellCommentReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(comment, "comment must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Layout metadata such as pane state, zoom, and visible row and column sizing for one sheet. */
  record SheetLayoutReport(
      String sheetName,
      PaneReport pane,
      int zoomPercent,
      List<ColumnLayoutReport> columns,
      List<RowLayoutReport> rows) {
    public SheetLayoutReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(pane, "pane must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (zoomPercent < 10 || zoomPercent > 400) {
        throw new IllegalArgumentException(
            "zoomPercent must be between 10 and 400 inclusive: " + zoomPercent);
      }
      columns = copyValues(columns, "columns");
      rows = copyValues(rows, "rows");
    }
  }

  /** Width metadata for one sheet column. */
  record ColumnLayoutReport(
      int columnIndex,
      double widthCharacters,
      boolean hidden,
      int outlineLevel,
      boolean collapsed) {
    public ColumnLayoutReport {
      if (columnIndex < 0) {
        throw new IllegalArgumentException("columnIndex must not be negative");
      }
      if (!Double.isFinite(widthCharacters) || widthCharacters <= 0.0d) {
        throw new IllegalArgumentException("widthCharacters must be finite and greater than 0");
      }
      if (outlineLevel < 0) {
        throw new IllegalArgumentException("outlineLevel must not be negative");
      }
    }
  }

  /** Height metadata for one sheet row. */
  record RowLayoutReport(
      int rowIndex, double heightPoints, boolean hidden, int outlineLevel, boolean collapsed) {
    public RowLayoutReport {
      if (rowIndex < 0) {
        throw new IllegalArgumentException("rowIndex must not be negative");
      }
      if (!Double.isFinite(heightPoints) || heightPoints <= 0.0d) {
        throw new IllegalArgumentException("heightPoints must be finite and greater than 0");
      }
      if (outlineLevel < 0) {
        throw new IllegalArgumentException("outlineLevel must not be negative");
      }
    }
  }

  /** Grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceReport(int totalFormulaCellCount, List<SheetFormulaSurfaceReport> sheets) {
    public FormulaSurfaceReport {
      sheets = copyValues(sheets, "sheets");
      if (totalFormulaCellCount < 0) {
        throw new IllegalArgumentException("totalFormulaCellCount must not be negative");
      }
    }
  }

  /** Formula usage facts for one sheet. */
  record SheetFormulaSurfaceReport(
      String sheetName,
      int formulaCellCount,
      int distinctFormulaCount,
      List<FormulaPatternReport> formulas) {
    public SheetFormulaSurfaceReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
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
  record FormulaPatternReport(String formula, int occurrenceCount, List<String> addresses) {
    public FormulaPatternReport {
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      if (occurrenceCount <= 0) {
        throw new IllegalArgumentException("occurrenceCount must be greater than 0");
      }
      addresses = copyStrings(addresses, "addresses");
    }
  }

  /** Inferred schema facts for one rectangular sheet window. */
  record SheetSchemaReport(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      int dataRowCount,
      List<SchemaColumnReport> columns) {
    public SheetSchemaReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
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
  record SchemaColumnReport(
      int columnIndex,
      String columnAddress,
      String headerDisplayValue,
      int populatedCellCount,
      int blankCellCount,
      List<TypeCountReport> observedTypes,
      String dominantType) {
    public SchemaColumnReport {
      if (columnIndex < 0) {
        throw new IllegalArgumentException("columnIndex must not be negative");
      }
      Objects.requireNonNull(columnAddress, "columnAddress must not be null");
      Objects.requireNonNull(headerDisplayValue, "headerDisplayValue must not be null");
      if (columnAddress.isBlank()) {
        throw new IllegalArgumentException("columnAddress must not be blank");
      }
      if (populatedCellCount < 0) {
        throw new IllegalArgumentException("populatedCellCount must not be negative");
      }
      if (blankCellCount < 0) {
        throw new IllegalArgumentException("blankCellCount must not be negative");
      }
      observedTypes = copyValues(observedTypes, "observedTypes");
    }
  }

  /** Count of one observed cell value type inside a schema column. */
  record TypeCountReport(String type, int count) {
    public TypeCountReport {
      Objects.requireNonNull(type, "type must not be null");
      if (type.isBlank()) {
        throw new IllegalArgumentException("type must not be blank");
      }
      if (count <= 0) {
        throw new IllegalArgumentException("count must be greater than 0");
      }
    }
  }

  /** High-level characterization of the named ranges selected by one analysis read. */
  record NamedRangeSurfaceReport(
      int workbookScopedCount,
      int sheetScopedCount,
      int rangeBackedCount,
      int formulaBackedCount,
      List<NamedRangeSurfaceEntryReport> namedRanges) {
    public NamedRangeSurfaceReport {
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
  record NamedRangeSurfaceEntryReport(
      String name, NamedRangeScope scope, String refersToFormula, NamedRangeBackingKind kind) {
    public NamedRangeSurfaceEntryReport {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      Objects.requireNonNull(kind, "kind must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
    }
  }

  /** Distinguishes range-backed named ranges from formula-backed named ranges. */
  enum NamedRangeBackingKind {
    RANGE,
    FORMULA
  }

  /** Summary counts for one finding-bearing analysis result. */
  record AnalysisSummaryReport(int totalCount, int errorCount, int warningCount, int infoCount) {
    public AnalysisSummaryReport {
      if (totalCount < 0) {
        throw new IllegalArgumentException("totalCount must not be negative");
      }
      if (errorCount < 0) {
        throw new IllegalArgumentException("errorCount must not be negative");
      }
      if (warningCount < 0) {
        throw new IllegalArgumentException("warningCount must not be negative");
      }
      if (infoCount < 0) {
        throw new IllegalArgumentException("infoCount must not be negative");
      }
      if (totalCount != errorCount + warningCount + infoCount) {
        throw new IllegalArgumentException(
            "totalCount must equal errorCount + warningCount + infoCount");
      }
    }
  }

  /** Precise workbook location attached to one derived analysis finding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = AnalysisLocationReport.Workbook.class, name = "WORKBOOK"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.Sheet.class, name = "SHEET"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.Cell.class, name = "CELL"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.Range.class, name = "RANGE"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.NamedRange.class, name = "NAMED_RANGE")
  })
  sealed interface AnalysisLocationReport
      permits AnalysisLocationReport.Workbook,
          AnalysisLocationReport.Sheet,
          AnalysisLocationReport.Cell,
          AnalysisLocationReport.Range,
          AnalysisLocationReport.NamedRange {

    /** Workbook-level finding with no narrower location. */
    record Workbook() implements AnalysisLocationReport {}

    /** One whole-sheet finding. */
    record Sheet(String sheetName) implements AnalysisLocationReport {
      public Sheet {
        Objects.requireNonNull(sheetName, "sheetName must not be null");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetName must not be blank");
        }
      }
    }

    /** One concrete cell finding. */
    record Cell(String sheetName, String address) implements AnalysisLocationReport {
      public Cell {
        Objects.requireNonNull(sheetName, "sheetName must not be null");
        Objects.requireNonNull(address, "address must not be null");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetName must not be blank");
        }
        if (address.isBlank()) {
          throw new IllegalArgumentException("address must not be blank");
        }
      }
    }

    /** One rectangular range finding. */
    record Range(String sheetName, String range) implements AnalysisLocationReport {
      public Range {
        Objects.requireNonNull(sheetName, "sheetName must not be null");
        Objects.requireNonNull(range, "range must not be null");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetName must not be blank");
        }
        if (range.isBlank()) {
          throw new IllegalArgumentException("range must not be blank");
        }
      }
    }

    /** One named-range finding. */
    record NamedRange(String name, NamedRangeScope scope) implements AnalysisLocationReport {
      public NamedRange {
        Objects.requireNonNull(name, "name must not be null");
        Objects.requireNonNull(scope, "scope must not be null");
        if (name.isBlank()) {
          throw new IllegalArgumentException("name must not be blank");
        }
      }
    }
  }

  /** One reusable derived finding emitted by an analysis read. */
  record AnalysisFindingReport(
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String title,
      String message,
      AnalysisLocationReport location,
      List<String> evidence) {
    public AnalysisFindingReport {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(severity, "severity must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(location, "location must not be null");
      if (title.isBlank()) {
        throw new IllegalArgumentException("title must not be blank");
      }
      if (message.isBlank()) {
        throw new IllegalArgumentException("message must not be blank");
      }
      evidence = copyStrings(evidence, "evidence");
    }
  }

  /** Formula-health analysis for one selected sheet set. */
  record FormulaHealthReport(
      int checkedFormulaCellCount,
      AnalysisSummaryReport summary,
      List<AnalysisFindingReport> findings) {
    public FormulaHealthReport {
      if (checkedFormulaCellCount < 0) {
        throw new IllegalArgumentException("checkedFormulaCellCount must not be negative");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Hyperlink-health analysis for one selected sheet set. */
  record HyperlinkHealthReport(
      int checkedHyperlinkCount,
      AnalysisSummaryReport summary,
      List<AnalysisFindingReport> findings) {
    public HyperlinkHealthReport {
      if (checkedHyperlinkCount < 0) {
        throw new IllegalArgumentException("checkedHyperlinkCount must not be negative");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Named-range-health analysis for one selected named-range set. */
  record NamedRangeHealthReport(
      int checkedNamedRangeCount,
      AnalysisSummaryReport summary,
      List<AnalysisFindingReport> findings) {
    public NamedRangeHealthReport {
      if (checkedNamedRangeCount < 0) {
        throw new IllegalArgumentException("checkedNamedRangeCount must not be negative");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Aggregated workbook findings composed from the first analysis family. */
  record WorkbookFindingsReport(
      AnalysisSummaryReport summary, List<AnalysisFindingReport> findings) {
    public WorkbookFindingsReport {
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Deterministic failure payload returned for unsuccessful executions. */
  record Problem(
      GridGrindProblemCode code,
      GridGrindProblemCategory category,
      GridGrindProblemRecovery recovery,
      String title,
      String message,
      String resolution,
      ProblemContext context,
      List<ProblemCause> causes) {
    public Problem {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(category, "category must not be null");
      Objects.requireNonNull(recovery, "recovery must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(resolution, "resolution must not be null");
      Objects.requireNonNull(context, "context must not be null");
      causes = copyProblemCauses(causes);
    }

    /**
     * Constructs a Problem from a code and message, deriving category, recovery, title, and
     * resolution from the code.
     */
    static Problem of(GridGrindProblemCode code, String message, ProblemContext context) {
      return new Problem(
          code,
          code.category(),
          code.recovery(),
          code.title(),
          message,
          code.resolution(),
          context,
          List.of());
    }
  }

  /** Structured execution metadata that pinpoints where and why a failure happened. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "stage")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = ProblemContext.ParseArguments.class, name = "PARSE_ARGUMENTS"),
    @JsonSubTypes.Type(value = ProblemContext.ReadRequest.class, name = "READ_REQUEST"),
    @JsonSubTypes.Type(value = ProblemContext.ValidateRequest.class, name = "VALIDATE_REQUEST"),
    @JsonSubTypes.Type(value = ProblemContext.OpenWorkbook.class, name = "OPEN_WORKBOOK"),
    @JsonSubTypes.Type(value = ProblemContext.ApplyOperation.class, name = "APPLY_OPERATION"),
    @JsonSubTypes.Type(value = ProblemContext.ExecuteRead.class, name = "EXECUTE_READ"),
    @JsonSubTypes.Type(value = ProblemContext.PersistWorkbook.class, name = "PERSIST_WORKBOOK"),
    @JsonSubTypes.Type(value = ProblemContext.ExecuteRequest.class, name = "EXECUTE_REQUEST"),
    @JsonSubTypes.Type(value = ProblemContext.WriteResponse.class, name = "WRITE_RESPONSE")
  })
  sealed interface ProblemContext {
    /** Pipeline stage in which the failure occurred. */
    String stage();

    /**
     * Source type string; populated by ValidateRequest, OpenWorkbook, ApplyOperation,
     * PersistWorkbook, AnalyzeWorkbook, and ExecuteRequest. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String sourceType() {
      return null;
    }

    /**
     * Persistence type string; populated by ValidateRequest, OpenWorkbook, ApplyOperation,
     * PersistWorkbook, AnalyzeWorkbook, and ExecuteRequest. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String persistenceType() {
      return null;
    }

    /**
     * Path to the JSON request file; populated by ReadRequest. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String requestPath() {
      return null;
    }

    /**
     * JSON pointer path to the failing element; populated by ReadRequest. Null for all other
     * subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String jsonPath() {
      return null;
    }

    /**
     * Line number within the JSON payload where the error was detected; populated by ReadRequest.
     * Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default Integer jsonLine() {
      return null;
    }

    /**
     * Column number within the JSON payload where the error was detected; populated by ReadRequest.
     * Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default Integer jsonColumn() {
      return null;
    }

    /**
     * Path to the JSON response file; populated by WriteResponse. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String responsePath() {
      return null;
    }

    /**
     * Path to the source workbook file; populated by OpenWorkbook and PersistWorkbook. Null for all
     * other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String sourceWorkbookPath() {
      return null;
    }

    /**
     * Path to the persistence destination; populated by PersistWorkbook. Null for all other
     * subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String persistencePath() {
      return null;
    }

    /**
     * Zero-based index of the operation that failed; populated by ApplyOperation. Null for all
     * other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default Integer operationIndex() {
      return null;
    }

    /**
     * Zero-based index of the read that failed; populated by ExecuteRead. Null for all other
     * subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default Integer readIndex() {
      return null;
    }

    /**
     * SCREAMING_SNAKE_CASE type name of the failing operation; populated by ApplyOperation. Null
     * for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String operationType() {
      return null;
    }

    /**
     * SCREAMING_SNAKE_CASE type name of the failing read; populated by ExecuteRead. Null for all
     * other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String readType() {
      return null;
    }

    /**
     * Stable read request identifier; populated by ExecuteRead. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String requestId() {
      return null;
    }

    /**
     * Name of the sheet involved in the failure; populated by ApplyOperation and ExecuteRead. Null
     * for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String sheetName() {
      return null;
    }

    /**
     * Cell address in A1 notation where the failure occurred; populated by ApplyOperation and
     * ExecuteRead. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String address() {
      return null;
    }

    /**
     * Range address where the failure occurred; populated by ApplyOperation. Null for all other
     * subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String range() {
      return null;
    }

    /**
     * Formula text associated with the failure; populated by ApplyOperation and ExecuteRead. Null
     * for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String formula() {
      return null;
    }

    /**
     * Named-range identifier involved in the failure; populated by ApplyOperation and ExecuteRead.
     * Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String namedRangeName() {
      return null;
    }

    /**
     * CLI argument that triggered the failure; populated by ParseArguments. Null for all other
     * subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String argument() {
      return null;
    }

    /** Context for failures that occur while parsing CLI arguments. */
    record ParseArguments(String argument) implements ProblemContext {
      @Override
      public String stage() {
        return "PARSE_ARGUMENTS";
      }
    }

    /** Context for failures that occur while reading and parsing the JSON request. */
    record ReadRequest(String requestPath, String jsonPath, Integer jsonLine, Integer jsonColumn)
        implements ProblemContext {
      @Override
      public String stage() {
        return "READ_REQUEST";
      }

      /**
       * Returns a new ReadRequest with JSON location details merged in, keeping any existing
       * non-null values.
       */
      public ReadRequest withJson(String jsonPath, Integer jsonLine, Integer jsonColumn) {
        return new ReadRequest(
            requestPath,
            this.jsonPath != null ? this.jsonPath : jsonPath,
            this.jsonLine != null ? this.jsonLine : jsonLine,
            this.jsonColumn != null ? this.jsonColumn : jsonColumn);
      }
    }

    /** Context for failures that occur while validating request fields before execution. */
    record ValidateRequest(String sourceType, String persistenceType) implements ProblemContext {
      @Override
      public String stage() {
        return "VALIDATE_REQUEST";
      }
    }

    /** Context for failures that occur while opening the source workbook. */
    record OpenWorkbook(String sourceType, String persistenceType, String sourceWorkbookPath)
        implements ProblemContext {
      @Override
      public String stage() {
        return "OPEN_WORKBOOK";
      }
    }

    /** Context for failures that occur while applying a single workbook operation. */
    record ApplyOperation(
        String sourceType,
        String persistenceType,
        Integer operationIndex,
        String operationType,
        String sheetName,
        String address,
        String range,
        String formula,
        String namedRangeName)
        implements ProblemContext {
      @Override
      public String stage() {
        return "APPLY_OPERATION";
      }

      /**
       * Returns a new ApplyOperation with exception-derived location details merged in, keeping any
       * existing non-null values.
       */
      public ApplyOperation withExceptionData(
          String sheetName, String address, String range, String formula, String namedRangeName) {
        return new ApplyOperation(
            sourceType,
            persistenceType,
            operationIndex,
            operationType,
            this.sheetName != null ? this.sheetName : sheetName,
            this.address != null ? this.address : address,
            this.range != null ? this.range : range,
            this.formula != null ? this.formula : formula,
            this.namedRangeName != null ? this.namedRangeName : namedRangeName);
      }
    }

    /** Context for failures that occur while persisting the workbook to its destination path. */
    record PersistWorkbook(
        String sourceType,
        String persistenceType,
        String sourceWorkbookPath,
        String persistencePath)
        implements ProblemContext {
      @Override
      public String stage() {
        return "PERSIST_WORKBOOK";
      }
    }

    /** Context for failures that occur while executing one typed read after mutations finish. */
    record ExecuteRead(
        String sourceType,
        String persistenceType,
        Integer readIndex,
        String readType,
        String requestId,
        String sheetName,
        String address,
        String formula,
        String namedRangeName)
        implements ProblemContext {
      @Override
      public String stage() {
        return "EXECUTE_READ";
      }

      /**
       * Returns a new ExecuteRead with exception-derived location details merged in, keeping any
       * existing non-null values.
       */
      public ExecuteRead withExceptionData(
          String sheetName, String address, String formula, String namedRangeName) {
        return new ExecuteRead(
            sourceType,
            persistenceType,
            readIndex,
            readType,
            requestId,
            this.sheetName != null ? this.sheetName : sheetName,
            this.address != null ? this.address : address,
            this.formula != null ? this.formula : formula,
            this.namedRangeName != null ? this.namedRangeName : namedRangeName);
      }
    }

    /** Context for top-level execution failures not attributed to a specific pipeline stage. */
    record ExecuteRequest(String sourceType, String persistenceType) implements ProblemContext {
      @Override
      public String stage() {
        return "EXECUTE_REQUEST";
      }
    }

    /** Context for failures that occur while writing the JSON response to its destination. */
    record WriteResponse(String responsePath) implements ProblemContext {
      @Override
      public String stage() {
        return "WRITE_RESPONSE";
      }
    }
  }

  /**
   * Individual GridGrind-classified diagnostic entries preserved for agent-visible troubleshooting.
   */
  record ProblemCause(GridGrindProblemCode code, String message, String stage) {
    public ProblemCause {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(message, "message must not be null");
      if (message.isBlank()) {
        throw new IllegalArgumentException("message must not be blank");
      }
      if (stage != null && stage.isBlank()) {
        throw new IllegalArgumentException("stage must not be blank");
      }
    }
  }

  private static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static List<String> copyDistinctStrings(List<String> values, String fieldName) {
    List<String> copy = copyStrings(values, fieldName);
    if (copy.size() != new java.util.LinkedHashSet<>(copy).size()) {
      throw new IllegalArgumentException(fieldName + " must not contain duplicates");
    }
    return copy;
  }

  private static List<String> validateCommonWorkbookSummaryFields(
      int sheetCount, List<String> sheetNames, int namedRangeCount) {
    if (sheetCount < 0) {
      throw new IllegalArgumentException("sheetCount must not be negative");
    }
    if (namedRangeCount < 0) {
      throw new IllegalArgumentException("namedRangeCount must not be negative");
    }
    List<String> copy = copyDistinctStrings(sheetNames, "sheetNames");
    if (sheetCount != copy.size()) {
      throw new IllegalArgumentException("sheetCount must match sheetNames size");
    }
    for (String sheetName : copy) {
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetNames must not contain blank values");
      }
    }
    return copy;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  @SuppressWarnings("PMD.ReturnEmptyCollectionRatherThanNull")
  private static List<RichTextRunReport> copyRichTextRuns(
      List<RichTextRunReport> values, String fieldName) {
    if (values == null) {
      return null;
    }
    List<RichTextRunReport> copy = copyValues(values, fieldName);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copy;
  }

  private static String concatenateRichTextRuns(List<RichTextRunReport> runs) {
    StringBuilder builder = new StringBuilder();
    for (RichTextRunReport run : runs) {
      builder.append(run.text());
    }
    return builder.toString();
  }

  private static List<ProblemCause> copyProblemCauses(List<ProblemCause> causes) {
    if (causes == null) {
      return List.of();
    }
    List<ProblemCause> copy = List.copyOf(causes);
    for (ProblemCause cause : copy) {
      Objects.requireNonNull(cause, "causes must not contain nulls");
    }
    return copy;
  }
}
