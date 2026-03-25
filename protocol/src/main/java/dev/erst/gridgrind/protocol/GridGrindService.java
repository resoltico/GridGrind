package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelPreviewRow;
import dev.erst.gridgrind.excel.ExcelSheet;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Orchestrates protocol requests against the workbook core and returns structured responses. */
public final class GridGrindService {
  private final WorkbookCommandExecutor commandExecutor;
  private final WorkbookCloser workbookCloser;

  /** Creates the production GridGrindService with the default command executor and closer. */
  public GridGrindService() {
    this(new WorkbookCommandExecutor(), ExcelWorkbook::close);
  }

  /** Constructor for testing, allowing injection of a custom executor and closer. */
  GridGrindService(WorkbookCommandExecutor commandExecutor, WorkbookCloser workbookCloser) {
    this.commandExecutor =
        Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    this.workbookCloser = Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
  }

  /** Executes one complete GridGrind request. */
  public GridGrindResponse execute(GridGrindRequest request) {
    GridGrindProtocolVersion protocolVersion =
        request == null ? GridGrindProtocolVersion.current() : request.protocolVersion();
    if (request == null) {
      return new GridGrindResponse.Failure(
          protocolVersion,
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              "request must not be null",
              new GridGrindResponse.ProblemContext.ValidateRequest(null, null),
              (Throwable) null));
    }

    Optional<GridGrindResponse> validationError = validateRequest(protocolVersion, request);
    if (validationError.isPresent()) {
      return validationError.get();
    }

    ExcelWorkbook workbook;
    try {
      workbook = openWorkbook(request.source());
    } catch (Exception exception) {
      return new GridGrindResponse.Failure(
          protocolVersion,
          GridGrindProblems.fromException(
              exception,
              new GridGrindResponse.ProblemContext.OpenWorkbook(
                  reqSourceMode(request), reqPersistenceMode(request), reqSourcePath(request))));
    }

    return guardUnexpectedRuntime(
        protocolVersion,
        request,
        workbook,
        () -> {
          GridGrindResponse response;
          for (int operationIndex = 0;
              operationIndex < request.operations().size();
              operationIndex++) {
            WorkbookOperation operation = request.operations().get(operationIndex);
            try {
              commandExecutor.apply(workbook, toCommand(operation));
            } catch (Exception exception) {
              response =
                  operationFailure(protocolVersion, request, operationIndex, operation, exception);
              return closeWorkbook(workbook, response, request);
            }
          }

          List<GridGrindResponse.SheetReport> reports =
              new ArrayList<>(request.analysis().sheets().size());
          for (GridGrindRequest.SheetInspectionRequest sheetRequest : request.analysis().sheets()) {
            try {
              reports.add(analyzeSheet(workbook, sheetRequest));
            } catch (AnalysisException exception) {
              response = analysisFailure(protocolVersion, request, exception);
              return closeWorkbook(workbook, response, request);
            }
          }

          String savedWorkbookPath;
          try {
            savedWorkbookPath =
                persistWorkbook(workbook, request.persistence(), reqSourcePath(request));
          } catch (Exception exception) {
            response =
                new GridGrindResponse.Failure(
                    protocolVersion,
                    GridGrindProblems.fromException(
                        exception,
                        new GridGrindResponse.ProblemContext.PersistWorkbook(
                            reqSourceMode(request),
                            reqPersistenceMode(request),
                            reqSourcePath(request),
                            persistencePath(request.source(), request.persistence()))));
            return closeWorkbook(workbook, response, request);
          }

          response =
              new GridGrindResponse.Success(
                  protocolVersion,
                  savedWorkbookPath,
                  new GridGrindResponse.WorkbookSummary(
                      workbook.sheetCount(),
                      workbook.sheetNames(),
                      workbook.forceFormulaRecalculationOnOpenEnabled()),
                  reports);
          return closeWorkbook(workbook, response, request);
        });
  }

  /**
   * Validates cross-field constraints that cannot be enforced at the record level. Returns a
   * failure response if the request is invalid, or empty if validation passes.
   */
  private Optional<GridGrindResponse> validateRequest(
      GridGrindProtocolVersion protocolVersion, GridGrindRequest request) {
    return switch (request.persistence()) {
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ ->
          switch (request.source()) {
            case GridGrindRequest.WorkbookSource.New _ ->
                Optional.of(
                    new GridGrindResponse.Failure(
                        protocolVersion,
                        GridGrindProblems.problem(
                            GridGrindProblemCode.INVALID_REQUEST,
                            "OVERWRITE persistence requires an EXISTING source; "
                                + "a NEW workbook has no source file to overwrite",
                            new GridGrindResponse.ProblemContext.ValidateRequest(
                                reqSourceMode(request), reqPersistenceMode(request)),
                            (Throwable) null)));
            case GridGrindRequest.WorkbookSource.ExistingFile _ -> Optional.empty();
          };
      case GridGrindRequest.WorkbookPersistence.None _ -> Optional.empty();
      case GridGrindRequest.WorkbookPersistence.SaveAs _ -> Optional.empty();
    };
  }

  private ExcelWorkbook openWorkbook(GridGrindRequest.WorkbookSource source) throws IOException {
    return switch (source) {
      case GridGrindRequest.WorkbookSource.New _ -> ExcelWorkbook.create();
      case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
          ExcelWorkbook.open(Path.of(existingFile.path()));
    };
  }

  private WorkbookCommand toCommand(WorkbookOperation operation) {
    return switch (operation) {
      case WorkbookOperation.EnsureSheet op -> new WorkbookCommand.CreateSheet(op.sheetName());
      case WorkbookOperation.SetCell op ->
          new WorkbookCommand.SetCell(op.sheetName(), op.address(), op.value().toExcelCellValue());
      case WorkbookOperation.SetRange op ->
          new WorkbookCommand.SetRange(
              op.sheetName(),
              op.range(),
              op.rows().stream()
                  .map(row -> row.stream().map(CellInput::toExcelCellValue).toList())
                  .toList());
      case WorkbookOperation.ClearRange op ->
          new WorkbookCommand.ClearRange(op.sheetName(), op.range());
      case WorkbookOperation.ApplyStyle op ->
          new WorkbookCommand.ApplyStyle(op.sheetName(), op.range(), op.style().toExcelCellStyle());
      case WorkbookOperation.AppendRow op ->
          new WorkbookCommand.AppendRow(
              op.sheetName(), op.values().stream().map(CellInput::toExcelCellValue).toList());
      case WorkbookOperation.AutoSizeColumns op ->
          new WorkbookCommand.AutoSizeColumns(op.sheetName());
      case WorkbookOperation.EvaluateFormulas _ -> new WorkbookCommand.EvaluateAllFormulas();
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ ->
          new WorkbookCommand.ForceFormulaRecalculationOnOpen();
    };
  }

  /**
   * Persists the workbook according to the persistence mode. For OverwriteSource, sourcePath is
   * guaranteed non-null by validateRequest (ExistingFile always paired with OverwriteSource).
   */
  private String persistWorkbook(
      ExcelWorkbook workbook, GridGrindRequest.WorkbookPersistence persistence, String sourcePath)
      throws IOException {
    return switch (persistence) {
      case GridGrindRequest.WorkbookPersistence.None _ -> null;
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> {
        Path path = Path.of(sourcePath);
        workbook.save(path);
        yield path.toAbsolutePath().toString();
      }
      case GridGrindRequest.WorkbookPersistence.SaveAs saveAs -> {
        Path path = Path.of(saveAs.path());
        workbook.save(path);
        yield path.toAbsolutePath().toString();
      }
    };
  }

  private GridGrindResponse.PreviewRowReport toPreviewRowReport(ExcelPreviewRow previewRow) {
    return new GridGrindResponse.PreviewRowReport(
        previewRow.rowIndex(), previewRow.cells().stream().map(this::toCellReport).toList());
  }

  private GridGrindResponse.SheetReport analyzeSheet(
      ExcelWorkbook workbook, GridGrindRequest.SheetInspectionRequest sheetRequest) {
    ExcelSheet sheet;
    try {
      sheet = workbook.sheet(sheetRequest.sheetName());
    } catch (Exception exception) {
      throw new AnalysisException(sheetRequest.sheetName(), null, exception);
    }

    List<GridGrindResponse.CellReport> requestedCells =
        new ArrayList<>(sheetRequest.cells().size());
    for (String address : sheetRequest.cells()) {
      try {
        requestedCells.add(toCellReport(sheet.snapshotCell(address)));
      } catch (Exception exception) {
        throw new AnalysisException(sheetRequest.sheetName(), address, exception);
      }
    }

    List<GridGrindResponse.PreviewRowReport> previewRows = List.of();
    if (sheetRequest.previewRowCount() != null) {
      previewRows =
          sheet.preview(sheetRequest.previewRowCount(), sheetRequest.previewColumnCount()).stream()
              .map(this::toPreviewRowReport)
              .toList();
    }

    return new GridGrindResponse.SheetReport(
        sheet.name(),
        sheet.physicalRowCount(),
        sheet.lastRowIndex(),
        sheet.lastColumnIndex(),
        requestedCells,
        previewRows);
  }

  private GridGrindResponse.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    GridGrindResponse.CellStyleReport style =
        new GridGrindResponse.CellStyleReport(
            snapshot.style().numberFormat(),
            snapshot.style().bold(),
            snapshot.style().italic(),
            snapshot.style().wrapText(),
            snapshot.style().horizontalAlignment().name(),
            snapshot.style().verticalAlignment().name());

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot s ->
          new GridGrindResponse.CellReport.BlankReport(
              s.address(), s.declaredType(), s.displayValue(), style);
      case ExcelCellSnapshot.TextSnapshot s ->
          new GridGrindResponse.CellReport.TextReport(
              s.address(), s.declaredType(), s.displayValue(), style, s.stringValue());
      case ExcelCellSnapshot.NumberSnapshot s ->
          new GridGrindResponse.CellReport.NumberReport(
              s.address(), s.declaredType(), s.displayValue(), style, s.numberValue());
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new GridGrindResponse.CellReport.BooleanReport(
              s.address(), s.declaredType(), s.displayValue(), style, s.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot s ->
          new GridGrindResponse.CellReport.ErrorReport(
              s.address(), s.declaredType(), s.displayValue(), style, s.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot s ->
          new GridGrindResponse.CellReport.FormulaReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              s.formula(),
              toCellReport(s.evaluation()));
    };
  }

  /** Delegates to GridGrindProblems to classify the exception into a problem code. */
  static GridGrindProblemCode problemCodeFor(Exception exception) {
    return GridGrindProblems.codeFor(exception);
  }

  /** Delegates to GridGrindProblems to extract a user-readable problem message. */
  static String messageFor(Exception exception) {
    return GridGrindProblems.messageFor(exception);
  }

  private String absolutePath(String path) {
    return Path.of(path).toAbsolutePath().toString();
  }

  private GridGrindResponse closeWorkbook(
      ExcelWorkbook workbook, GridGrindResponse response, GridGrindRequest request) {
    try {
      workbookCloser.close(workbook);
      return response;
    } catch (IOException exception) {
      return switch (response) {
        case GridGrindResponse.Failure f ->
            new GridGrindResponse.Failure(
                f.protocolVersion(),
                GridGrindProblems.appendCause(
                    f.problem(),
                    GridGrindProblems.supplementalCause(
                        "EXECUTE_REQUEST",
                        exception,
                        "Workbook close failed after the primary problem")));
        case GridGrindResponse.Success _ ->
            new GridGrindResponse.Failure(
                response.protocolVersion(),
                GridGrindProblems.fromException(
                    exception,
                    new GridGrindResponse.ProblemContext.ExecuteRequest(
                        reqSourceMode(request), reqPersistenceMode(request))));
      };
    }
  }

  private GridGrindResponse guardUnexpectedRuntime(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      ExcelWorkbook workbook,
      java.util.function.Supplier<GridGrindResponse> execution) {
    try {
      return execution.get();
    } catch (RuntimeException exception) {
      GridGrindResponse response =
          new GridGrindResponse.Failure(
              protocolVersion,
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ExecuteRequest(
                      reqSourceMode(request), reqPersistenceMode(request))));
      return closeWorkbook(workbook, response, request);
    }
  }

  private String reqSourceMode(GridGrindRequest request) {
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.New _ -> "NEW";
      case GridGrindRequest.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  private String reqPersistenceMode(GridGrindRequest request) {
    return switch (request.persistence()) {
      case GridGrindRequest.WorkbookPersistence.None _ -> "NONE";
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case GridGrindRequest.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  private String reqSourcePath(GridGrindRequest request) {
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
          absolutePath(existingFile.path());
      case GridGrindRequest.WorkbookSource.New _ -> null;
    };
  }

  String persistencePath(
      GridGrindRequest.WorkbookSource source, GridGrindRequest.WorkbookPersistence persistence) {
    return switch (persistence) {
      case GridGrindRequest.WorkbookPersistence.SaveAs saveAs -> absolutePath(saveAs.path());
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ ->
          switch (source) {
            case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
                absolutePath(existingFile.path());
            case GridGrindRequest.WorkbookSource.New _ -> null;
          };
      case GridGrindRequest.WorkbookPersistence.None _ -> null;
    };
  }

  /**
   * Returns the formula string for the failure context: checks the exception first (e.g.,
   * FormulaException), then falls back to the operation record for SetCell with a formula value.
   */
  static String formulaFor(WorkbookOperation operation, Exception exception) {
    String fromException = GridGrindProblems.formulaFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (operation) {
      case WorkbookOperation.SetCell op ->
          switch (op.value()) {
            case CellInput.Formula f -> f.formula();
            case CellInput.Blank _ -> null;
            case CellInput.Text _ -> null;
            case CellInput.Numeric _ -> null;
            case CellInput.BooleanValue _ -> null;
            case CellInput.Date _ -> null;
            case CellInput.DateTime _ -> null;
          };
      case WorkbookOperation.EnsureSheet _ -> null;
      case WorkbookOperation.SetRange _ -> null;
      case WorkbookOperation.ClearRange _ -> null;
      case WorkbookOperation.ApplyStyle _ -> null;
      case WorkbookOperation.AppendRow _ -> null;
      case WorkbookOperation.AutoSizeColumns _ -> null;
      case WorkbookOperation.EvaluateFormulas _ -> null;
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /**
   * Returns the sheet name for the failure context: checks the operation record first, then falls
   * back to the exception (e.g., FormulaException carrying sheet context).
   */
  static String sheetNameFor(WorkbookOperation operation, Exception exception) {
    String fromOp =
        switch (operation) {
          case WorkbookOperation.EnsureSheet op -> op.sheetName();
          case WorkbookOperation.SetCell op -> op.sheetName();
          case WorkbookOperation.SetRange op -> op.sheetName();
          case WorkbookOperation.ClearRange op -> op.sheetName();
          case WorkbookOperation.ApplyStyle op -> op.sheetName();
          case WorkbookOperation.AppendRow op -> op.sheetName();
          case WorkbookOperation.AutoSizeColumns op -> op.sheetName();
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOp != null ? fromOp : GridGrindProblems.sheetNameFor(exception);
  }

  /**
   * Returns the cell address for the failure context: checks the operation record first, then falls
   * back to the exception (e.g., FormulaException carrying address context).
   */
  static String addressFor(WorkbookOperation operation, Exception exception) {
    String fromOp =
        switch (operation) {
          case WorkbookOperation.SetCell op -> op.address();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.SetRange _ -> null;
          case WorkbookOperation.ClearRange _ -> null;
          case WorkbookOperation.ApplyStyle _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOp != null ? fromOp : GridGrindProblems.addressFor(exception);
  }

  /**
   * Returns the range string for the failure context: checks the operation record first, then falls
   * back to the exception (e.g., InvalidRangeAddressException carrying the range).
   */
  static String rangeFor(WorkbookOperation operation, Exception exception) {
    String fromOp =
        switch (operation) {
          case WorkbookOperation.SetRange op -> op.range();
          case WorkbookOperation.ClearRange op -> op.range();
          case WorkbookOperation.ApplyStyle op -> op.range();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.SetCell _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOp != null ? fromOp : GridGrindProblems.rangeFor(exception);
  }

  private GridGrindResponse operationFailure(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      int operationIndex,
      WorkbookOperation operation,
      Exception exception) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        GridGrindProblems.fromException(
            exception,
            new GridGrindResponse.ProblemContext.ApplyOperation(
                reqSourceMode(request),
                reqPersistenceMode(request),
                operationIndex,
                operation.operationType(),
                sheetNameFor(operation, exception),
                addressFor(operation, exception),
                rangeFor(operation, exception),
                formulaFor(operation, exception))));
  }

  private GridGrindResponse analysisFailure(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      AnalysisException exception) {
    Exception cause = exception.cause();
    return new GridGrindResponse.Failure(
        protocolVersion,
        GridGrindProblems.fromException(
            cause,
            new GridGrindResponse.ProblemContext.AnalyzeWorkbook(
                reqSourceMode(request),
                reqPersistenceMode(request),
                exception.sheetName(),
                exception.address(),
                GridGrindProblems.formulaFor(cause))));
  }

  /** Functional interface for closing an ExcelWorkbook after request execution. */
  @FunctionalInterface
  interface WorkbookCloser {
    /** Closes the workbook, releasing any held resources. */
    void close(ExcelWorkbook workbook) throws IOException;
  }

  /** Wraps a sheet-level analysis failure with its location for propagation out of the lambda. */
  private static final class AnalysisException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    private final String sheetName;
    private final String address;

    private AnalysisException(String sheetName, String address, Exception cause) {
      super(cause);
      this.sheetName = sheetName;
      this.address = address;
    }

    private String sheetName() {
      return sheetName;
    }

    private String address() {
      return address;
    }

    private Exception cause() {
      return (Exception) getCause();
    }
  }
}
