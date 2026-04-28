package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

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

  /** Structured execution journal captured for this run, even when it failed. */
  ExecutionJournal journal();

  /**
   * Structured calculation report captured for this run, even when calculation was not requested.
   */
  CalculationReport calculation();

  /** Successful workbook execution result. */
  record Success(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournal journal,
      CalculationReport calculation,
      PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections)
      implements GridGrindResponse {
    public Success {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
      warnings =
          GridGrindResponseSupport.copyValues(
              Objects.requireNonNull(warnings, "warnings must not be null"), "warnings");
      assertions =
          GridGrindResponseSupport.copyValues(
              Objects.requireNonNull(assertions, "assertions must not be null"), "assertions");
      inspections = GridGrindResponseSupport.copyValues(inspections, "inspections");
    }
  }

  /** Failed workbook execution with a structured problem. */
  record Failure(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournal journal,
      CalculationReport calculation,
      Problem problem)
      implements GridGrindResponse {
    public Failure {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
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
        sheetNames =
            GridGrindResponseSupport.validateCommonWorkbookSummaryFields(
                sheetCount, sheetNames, namedRangeCount);
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
        sheetNames =
            GridGrindResponseSupport.validateCommonWorkbookSummaryFields(
                sheetCount, sheetNames, namedRangeCount);
        Objects.requireNonNull(activeSheetName, "activeSheetName must not be null");
        if (activeSheetName.isBlank()) {
          throw new IllegalArgumentException("activeSheetName must not be blank");
        }
        selectedSheetNames =
            GridGrindResponseSupport.copyDistinctStrings(selectedSheetNames, "selectedSheetNames");
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

    /** Named range that resolves cleanly to a sheet-qualified cell or rectangular range target. */
    record RangeReport(
        String name, NamedRangeScope scope, String refersToFormula, NamedRangeTarget target)
        implements NamedRangeReport {
      public RangeReport {
        Objects.requireNonNull(name, "name must not be null");
        if (name.isBlank()) {
          throw new IllegalArgumentException("name must not be blank");
        }
        Objects.requireNonNull(scope, "scope must not be null");
        Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
        if (refersToFormula.isBlank()) {
          throw new IllegalArgumentException("refersToFormula must not be blank");
        }
        Objects.requireNonNull(target, "target must not be null");
      }
    }

    /** Defined name whose formula cannot be normalized to a typed range target. */
    record FormulaReport(String name, NamedRangeScope scope, String refersToFormula)
        implements NamedRangeReport {
      public FormulaReport {
        Objects.requireNonNull(name, "name must not be null");
        if (name.isBlank()) {
          throw new IllegalArgumentException("name must not be blank");
        }
        Objects.requireNonNull(scope, "scope must not be null");
        Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
        if (refersToFormula.isBlank()) {
          throw new IllegalArgumentException("refersToFormula must not be blank");
        }
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

    /** Sheet protection is disabled. */
    record Unprotected() implements SheetProtectionReport {}

    /** Sheet protection is enabled with the reported supported lock flags. */
    record Protected(SheetProtectionSettings settings) implements SheetProtectionReport {
      public Protected {
        Objects.requireNonNull(settings, "settings must not be null");
      }
    }
  }

  /** Factual comment metadata returned for analyzed cells that carry a comment. */
  record CommentReport(
      String text,
      String author,
      boolean visible,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<java.util.List<RichTextRunReport>> runs,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentAnchorReport> anchor) {
    /** Creates a plain comment report without rich runs or anchor metadata. */
    public CommentReport(String text, String author, boolean visible) {
      this(text, author, visible, Optional.empty(), Optional.empty());
    }

    public CommentReport {
      Objects.requireNonNull(text, "text must not be null");
      Objects.requireNonNull(author, "author must not be null");
      runs = GridGrindResponseSupport.copyOptionalValues(runs, "runs");
      anchor = Objects.requireNonNullElseGet(anchor, Optional::empty);
      if (runs.isPresent()) {
        java.util.List<RichTextRunReport> copiedRuns = runs.orElseThrow();
        if (!text.equals(
            copiedRuns.stream()
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
      rows = GridGrindResponseSupport.copyValues(rows, "rows");
    }
  }

  /** One row inside a rectangular window of cell snapshots. */
  record WindowRowReport(int rowIndex, List<CellReport> cells) {
    public WindowRowReport {
      cells = GridGrindResponseSupport.copyValues(cells, "cells");
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
      SheetPresentationReport presentation,
      List<ColumnLayoutReport> columns,
      List<RowLayoutReport> rows) {
    public SheetLayoutReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(pane, "pane must not be null");
      Objects.requireNonNull(presentation, "presentation must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireZoomPercent(
          zoomPercent, "zoomPercent");
      columns = GridGrindResponseSupport.copyValues(columns, "columns");
      rows = GridGrindResponseSupport.copyValues(rows, "rows");
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
      sheets = GridGrindResponseSupport.copyValues(sheets, "sheets");
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
      formulas = GridGrindResponseSupport.copyValues(formulas, "formulas");
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
      addresses = GridGrindResponseSupport.copyStrings(addresses, "addresses");
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
      columns = GridGrindResponseSupport.copyValues(columns, "columns");
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
      observedTypes = GridGrindResponseSupport.copyValues(observedTypes, "observedTypes");
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
      namedRanges = GridGrindResponseSupport.copyValues(namedRanges, "namedRanges");
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
      evidence = GridGrindResponseSupport.copyStrings(evidence, "evidence");
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
      findings = GridGrindResponseSupport.copyValues(findings, "findings");
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
      findings = GridGrindResponseSupport.copyValues(findings, "findings");
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
      findings = GridGrindResponseSupport.copyValues(findings, "findings");
    }
  }

  /** Aggregated workbook findings composed from all shipped analysis families. */
  record WorkbookFindingsReport(
      AnalysisSummaryReport summary, List<AnalysisFindingReport> findings) {
    public WorkbookFindingsReport {
      Objects.requireNonNull(summary, "summary must not be null");
      findings = GridGrindResponseSupport.copyValues(findings, "findings");
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
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<AssertionFailure> assertionFailure,
      List<ProblemCause> causes) {
    /** Validates one deterministic problem payload and normalizes optional nested fields. */
    public Problem {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(category, "category must not be null");
      Objects.requireNonNull(recovery, "recovery must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(resolution, "resolution must not be null");
      Objects.requireNonNull(context, "context must not be null");
      assertionFailure = Objects.requireNonNullElseGet(assertionFailure, Optional::empty);
      causes = GridGrindResponseSupport.copyProblemCauses(causes);
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
          Optional.empty(),
          List.of());
    }
  }

  /**
   * Individual GridGrind-classified diagnostic entries preserved for agent-visible troubleshooting.
   */
  record ProblemCause(GridGrindProblemCode code, String message, String stage) {
    public ProblemCause {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(stage, "stage must not be null");
      if (message.isBlank()) {
        throw new IllegalArgumentException("message must not be blank");
      }
      if (stage.isBlank()) {
        throw new IllegalArgumentException("stage must not be blank");
      }
    }
  }

  /** Creates a synthetic success journal for non-step-oriented responses. */
  static ExecutionJournal syntheticSuccessJournal() {
    return GridGrindResponseSupport.syntheticSuccessJournal();
  }

  /** Creates a synthetic failed journal for non-step-oriented responses. */
  static ExecutionJournal syntheticFailureJournal(GridGrindProblemCode failureCode) {
    return GridGrindResponseSupport.syntheticFailureJournal(failureCode);
  }

  /** Creates a synthetic journal for non-step-oriented responses with explicit failure state. */
  static ExecutionJournal syntheticJournal(
      ExecutionJournal.Status status, Optional<GridGrindProblemCode> failureCode) {
    return GridGrindResponseSupport.syntheticJournal(status, failureCode);
  }
}
