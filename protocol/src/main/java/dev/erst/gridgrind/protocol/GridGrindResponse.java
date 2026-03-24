package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
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
      String savedWorkbookPath,
      WorkbookSummary workbook,
      List<SheetReport> sheets)
      implements GridGrindResponse {
    public Success {
      protocolVersion =
          protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
      Objects.requireNonNull(workbook, "workbook must not be null");
      sheets = copySheets(sheets);
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

  /** High-level workbook facts returned on success. */
  record WorkbookSummary(
      int sheetCount, List<String> sheetNames, boolean forceFormulaRecalculationOnOpen) {
    public WorkbookSummary {
      sheetNames = copyStrings(sheetNames, "sheetNames");
    }
  }

  /** Analysis report for one inspected sheet. */
  record SheetReport(
      String sheetName,
      int physicalRowCount,
      int lastRowIndex,
      int lastColumnIndex,
      List<CellReport> requestedCells,
      List<PreviewRowReport> previewRows) {
    public SheetReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      requestedCells = copyReports(requestedCells, "requestedCells");
      previewRows = copyPreviewRows(previewRows);
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
    @JsonSubTypes.Type(value = CellReport.NumberReport.class, name = "NUMERIC"),
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
        String address, String declaredType, String displayValue, CellStyleReport style)
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
        String stringValue)
        implements CellReport {
      public TextReport {
        Objects.requireNonNull(address, "address must not be null");
        Objects.requireNonNull(declaredType, "declaredType must not be null");
        Objects.requireNonNull(displayValue, "displayValue must not be null");
        Objects.requireNonNull(style, "style must not be null");
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
        return "NUMERIC";
      }
    }

    /** CellReport for a cell containing a boolean value. */
    record BooleanReport(
        String address,
        String declaredType,
        String displayValue,
        CellStyleReport style,
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

  /** Effective style facts returned with every analyzed cell. */
  record CellStyleReport(
      String numberFormat,
      boolean bold,
      boolean italic,
      boolean wrapText,
      String horizontalAlignment,
      String verticalAlignment) {
    public CellStyleReport {
      Objects.requireNonNull(numberFormat, "numberFormat must not be null");
      Objects.requireNonNull(horizontalAlignment, "horizontalAlignment must not be null");
      Objects.requireNonNull(verticalAlignment, "verticalAlignment must not be null");
    }
  }

  /** One preview row returned during sheet inspection. */
  record PreviewRowReport(int rowIndex, List<CellReport> cells) {
    public PreviewRowReport {
      cells = copyReports(cells, "cells");
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
    @JsonSubTypes.Type(value = ProblemContext.PersistWorkbook.class, name = "PERSIST_WORKBOOK"),
    @JsonSubTypes.Type(value = ProblemContext.AnalyzeWorkbook.class, name = "ANALYZE_WORKBOOK"),
    @JsonSubTypes.Type(value = ProblemContext.ExecuteRequest.class, name = "EXECUTE_REQUEST"),
    @JsonSubTypes.Type(value = ProblemContext.WriteResponse.class, name = "WRITE_RESPONSE")
  })
  sealed interface ProblemContext {
    /** Pipeline stage in which the failure occurred. */
    String stage();

    /**
     * Source mode string; populated by ValidateRequest, OpenWorkbook, ApplyOperation,
     * PersistWorkbook, AnalyzeWorkbook, and ExecuteRequest. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String sourceMode() {
      return null;
    }

    /**
     * Persistence mode string; populated by ValidateRequest, OpenWorkbook, ApplyOperation,
     * PersistWorkbook, AnalyzeWorkbook, and ExecuteRequest. Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String persistenceMode() {
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
     * Name of the sheet involved in the failure; populated by ApplyOperation and AnalyzeWorkbook.
     * Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String sheetName() {
      return null;
    }

    /**
     * Cell address in A1 notation where the failure occurred; populated by ApplyOperation and
     * AnalyzeWorkbook. Null for all other subtypes.
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
     * Formula text associated with the failure; populated by ApplyOperation and AnalyzeWorkbook.
     * Null for all other subtypes.
     *
     * <p>Null is permitted here because this is a protocol-layer default method on a sealed
     * interface used for wire serialization; internal code must use a switch expression instead.
     */
    default String formula() {
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
    record ValidateRequest(String sourceMode, String persistenceMode) implements ProblemContext {
      @Override
      public String stage() {
        return "VALIDATE_REQUEST";
      }
    }

    /** Context for failures that occur while opening the source workbook. */
    record OpenWorkbook(String sourceMode, String persistenceMode, String sourceWorkbookPath)
        implements ProblemContext {
      @Override
      public String stage() {
        return "OPEN_WORKBOOK";
      }
    }

    /** Context for failures that occur while applying a single workbook operation. */
    record ApplyOperation(
        String sourceMode,
        String persistenceMode,
        Integer operationIndex,
        String operationType,
        String sheetName,
        String address,
        String range,
        String formula)
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
          String sheetName, String address, String range, String formula) {
        return new ApplyOperation(
            sourceMode,
            persistenceMode,
            operationIndex,
            operationType,
            this.sheetName != null ? this.sheetName : sheetName,
            this.address != null ? this.address : address,
            this.range != null ? this.range : range,
            this.formula != null ? this.formula : formula);
      }
    }

    /** Context for failures that occur while persisting the workbook to its destination path. */
    record PersistWorkbook(
        String sourceMode,
        String persistenceMode,
        String sourceWorkbookPath,
        String persistencePath)
        implements ProblemContext {
      @Override
      public String stage() {
        return "PERSIST_WORKBOOK";
      }
    }

    /**
     * Context for failures that occur while analyzing the workbook after operations are applied.
     */
    record AnalyzeWorkbook(
        String sourceMode, String persistenceMode, String sheetName, String address, String formula)
        implements ProblemContext {
      @Override
      public String stage() {
        return "ANALYZE_WORKBOOK";
      }

      /**
       * Returns a new AnalyzeWorkbook with exception-derived location details merged in, keeping
       * any existing non-null values.
       */
      public AnalyzeWorkbook withExceptionData(String sheetName, String address, String formula) {
        return new AnalyzeWorkbook(
            sourceMode,
            persistenceMode,
            this.sheetName != null ? this.sheetName : sheetName,
            this.address != null ? this.address : address,
            this.formula != null ? this.formula : formula);
      }
    }

    /** Context for top-level execution failures not attributed to a specific pipeline stage. */
    record ExecuteRequest(String sourceMode, String persistenceMode) implements ProblemContext {
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

  /** Individual diagnostic cause entries preserved for agent-visible troubleshooting. */
  record ProblemCause(String type, String className, String message, String stage) {
    public ProblemCause {
      Objects.requireNonNull(type, "type must not be null");
      Objects.requireNonNull(className, "className must not be null");
      Objects.requireNonNull(message, "message must not be null");
    }
  }

  private static List<SheetReport> copySheets(List<SheetReport> sheets) {
    if (sheets == null) {
      return List.of();
    }
    List<SheetReport> copy = List.copyOf(sheets);
    for (SheetReport sheet : copy) {
      Objects.requireNonNull(sheet, "sheets must not contain nulls");
    }
    return copy;
  }

  private static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static List<CellReport> copyReports(List<CellReport> reports, String fieldName) {
    Objects.requireNonNull(reports, fieldName + " must not be null");
    List<CellReport> copy = List.copyOf(reports);
    for (CellReport report : copy) {
      Objects.requireNonNull(report, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static List<PreviewRowReport> copyPreviewRows(List<PreviewRowReport> previewRows) {
    Objects.requireNonNull(previewRows, "previewRows must not be null");
    List<PreviewRowReport> copy = List.copyOf(previewRows);
    for (PreviewRowReport previewRow : copy) {
      Objects.requireNonNull(previewRow, "previewRows must not contain nulls");
    }
    return copy;
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
