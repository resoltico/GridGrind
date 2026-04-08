package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.io.IOException;
import java.nio.file.Path;
import java.time.DateTimeException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Default request executor that applies the GridGrind workflow against the workbook core. */
public final class DefaultGridGrindRequestExecutor implements GridGrindRequestExecutor {
  private final WorkbookCommandExecutor commandExecutor;
  private final WorkbookReadExecutor readExecutor;
  private final WorkbookCloser workbookCloser;

  /** Creates the production request executor with the default workbook executors and closer. */
  public DefaultGridGrindRequestExecutor() {
    this(new WorkbookCommandExecutor(), new WorkbookReadExecutor(), ExcelWorkbook::close);
  }

  /** Constructor for testing, allowing injection of custom executors and closer. */
  DefaultGridGrindRequestExecutor(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      WorkbookCloser workbookCloser) {
    this.commandExecutor =
        Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    this.workbookCloser = Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
  }

  /** Executes one complete GridGrind request. */
  @Override
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
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.OpenWorkbook(
                  reqSourceType(request), reqPersistenceType(request), reqSourcePath(request))));
    }

    return guardUnexpectedRuntime(
        protocolVersion,
        request,
        workbook,
        () -> executeWorkbookWorkflow(protocolVersion, request, workbook));
  }

  private GridGrindResponse executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion, GridGrindRequest request, ExcelWorkbook workbook) {
    WorkbookLocation workbookLocation =
        workbookLocationFor(request.source(), request.persistence());
    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);
    for (int operationIndex = 0; operationIndex < request.operations().size(); operationIndex++) {
      WorkbookOperation operation = request.operations().get(operationIndex);
      try {
        commandExecutor.apply(workbook, WorkbookCommandConverter.toCommand(operation));
      } catch (Exception exception) {
        return closeWorkbook(
            workbook,
            operationFailure(protocolVersion, request, operationIndex, operation, exception),
            request);
      }
    }

    List<WorkbookReadResult> reads = new ArrayList<>(request.reads().size());
    for (int readIndex = 0; readIndex < request.reads().size(); readIndex++) {
      WorkbookReadOperation read = request.reads().get(readIndex);
      try {
        WorkbookReadCommand command = WorkbookReadCommandConverter.toReadCommand(read);
        dev.erst.gridgrind.excel.WorkbookReadResult result =
            readExecutor.apply(workbook, workbookLocation, command).getFirst();
        reads.add(WorkbookReadResultConverter.toReadResult(result));
      } catch (Exception exception) {
        return closeWorkbook(
            workbook, readFailure(protocolVersion, request, readIndex, read, exception), request);
      }
    }

    GridGrindResponse.PersistenceOutcome persistence;
    try {
      persistence = persistWorkbook(workbook, request.source(), request.persistence());
    } catch (Exception exception) {
      return closeWorkbook(
          workbook,
          new GridGrindResponse.Failure(
              protocolVersion,
              problemFor(
                  exception,
                  new GridGrindResponse.ProblemContext.PersistWorkbook(
                      reqSourceType(request),
                      reqPersistenceType(request),
                      reqSourcePath(request),
                      persistencePath(request.source(), request.persistence())))),
          request);
    }

    return closeWorkbook(
        workbook,
        new GridGrindResponse.Success(protocolVersion, persistence, warnings, List.copyOf(reads)),
        request);
  }

  /** Validates cross-field constraints that cannot be enforced at the record level. */
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
                                reqSourceType(request), reqPersistenceType(request)),
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

  GridGrindResponse.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      GridGrindRequest.WorkbookSource source,
      GridGrindRequest.WorkbookPersistence persistence)
      throws IOException {
    return switch (persistence) {
      case GridGrindRequest.WorkbookPersistence.None _ ->
          new GridGrindResponse.PersistenceOutcome.NotSaved();
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> {
        Path rawPath =
            switch (source) {
              case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
                  Path.of(existingFile.path());
              case GridGrindRequest.WorkbookSource.New _ ->
                  throw new IllegalArgumentException(
                      "OVERWRITE persistence requires an EXISTING source");
            };
        Path normalizedPath = rawPath.toAbsolutePath().normalize();
        workbook.save(normalizedPath);
        yield new GridGrindResponse.PersistenceOutcome.Overwritten(
            rawPath.toString(), normalizedPath.toString());
      }
      case GridGrindRequest.WorkbookPersistence.SaveAs saveAs -> {
        Path normalizedPath = Path.of(saveAs.path()).toAbsolutePath().normalize();
        workbook.save(normalizedPath);
        yield new GridGrindResponse.PersistenceOutcome.SavedAs(
            saveAs.path(), normalizedPath.toString());
      }
    };
  }

  /**
   * Classifies one request-execution exception into the corresponding public problem code.
   *
   * <p>Engine exceptions are owned here so protocol-level problem construction remains engine-blind
   * outside the single request-execution bridge.
   */
  static GridGrindProblemCode problemCodeFor(Exception exception) {
    return switch (exception) {
      case WorkbookNotFoundException _ -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      case SheetNotFoundException _ -> GridGrindProblemCode.SHEET_NOT_FOUND;
      case NamedRangeNotFoundException _ -> GridGrindProblemCode.NAMED_RANGE_NOT_FOUND;
      case CellNotFoundException _ -> GridGrindProblemCode.CELL_NOT_FOUND;
      case InvalidCellAddressException _ -> GridGrindProblemCode.INVALID_CELL_ADDRESS;
      case InvalidRangeAddressException _ -> GridGrindProblemCode.INVALID_RANGE_ADDRESS;
      case UnsupportedFormulaException _ -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
      case InvalidFormulaException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case IOException _ -> GridGrindProblemCode.IO_ERROR;
      case DateTimeException _ -> GridGrindProblemCode.INVALID_REQUEST;
      default -> GridGrindProblems.codeFor(exception);
    };
  }

  /** Returns the public problem message for one execution exception. */
  static String messageFor(Exception exception) {
    return GridGrindProblems.messageFor(exception);
  }

  private static GridGrindResponse.Problem problemFor(
      Exception exception, GridGrindResponse.ProblemContext context) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(context, "context must not be null");
    return GridGrindProblems.problem(
        problemCodeFor(exception),
        messageFor(exception),
        enrichContext(context, exception),
        causesFor(exception, context.stage()));
  }

  private static List<GridGrindResponse.ProblemCause> causesFor(Exception exception, String stage) {
    return List.of(
        new GridGrindResponse.ProblemCause(
            problemCodeFor(exception), messageFor(exception), stage));
  }

  /**
   * Enriches one problem context with execution-specific exception fields while preserving payload
   * coordinates supplied by protocol parsing failures.
   */
  static GridGrindResponse.ProblemContext enrichContext(
      GridGrindResponse.ProblemContext context, Exception exception) {
    GridGrindResponse.ProblemContext protocolEnriched =
        GridGrindProblems.enrichContext(context, exception);
    return switch (protocolEnriched) {
      case GridGrindResponse.ProblemContext.ApplyOperation applyOperation ->
          enrichApplyOperationContext(applyOperation, exception);
      case GridGrindResponse.ProblemContext.ExecuteRead executeRead ->
          enrichExecuteReadContext(executeRead, exception);
      case GridGrindResponse.ProblemContext.ReadRequest _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.ParseArguments _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.ValidateRequest _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.OpenWorkbook _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.PersistWorkbook _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.ExecuteRequest _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.WriteResponse _ -> protocolEnriched;
    };
  }

  private static GridGrindResponse.ProblemContext.ApplyOperation enrichApplyOperationContext(
      GridGrindResponse.ProblemContext.ApplyOperation context, Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return context.withExceptionData(
          formulaException.sheetName(),
          formulaException.address(),
          null,
          formulaException.formula(),
          namedRangeNameFor(exception));
    }
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return context.withExceptionData(
          null, null, invalidRangeAddressException.range(), null, namedRangeNameFor(exception));
    }
    if (exception instanceof NamedRangeNotFoundException) {
      return context.withExceptionData(null, null, null, null, namedRangeNameFor(exception));
    }
    return context;
  }

  private static GridGrindResponse.ProblemContext.ExecuteRead enrichExecuteReadContext(
      GridGrindResponse.ProblemContext.ExecuteRead context, Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return context.withExceptionData(
          formulaException.sheetName(),
          formulaException.address(),
          formulaException.formula(),
          namedRangeNameFor(exception));
    }
    if (exception instanceof NamedRangeNotFoundException) {
      return context.withExceptionData(null, null, null, namedRangeNameFor(exception));
    }
    return context;
  }

  static String formulaFor(Exception exception) {
    return exception instanceof FormulaException formulaException
        ? formulaException.formula()
        : null;
  }

  static String sheetNameFor(Exception exception) {
    return exception instanceof FormulaException formulaException
        ? formulaException.sheetName()
        : null;
  }

  static String addressFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.address();
      case CellNotFoundException cellNotFoundException -> cellNotFoundException.address();
      case InvalidCellAddressException invalidCellAddressException ->
          invalidCellAddressException.address();
      default -> null;
    };
  }

  static String rangeFor(Exception exception) {
    return exception instanceof InvalidRangeAddressException invalidRangeAddressException
        ? invalidRangeAddressException.range()
        : null;
  }

  static String namedRangeNameFor(Exception exception) {
    return exception instanceof NamedRangeNotFoundException namedRangeNotFoundException
        ? namedRangeNotFoundException.name()
        : null;
  }

  /**
   * Returns the workbook filesystem context implied by the source and persistence selection.
   *
   * <p>Package-private so tests can assert persisted-path resolution directly without going through
   * full request execution.
   */
  static WorkbookLocation workbookLocationFor(
      GridGrindRequest.WorkbookSource source, GridGrindRequest.WorkbookPersistence persistence) {
    String persistedPath = persistencePath(source, persistence);
    if (persistedPath != null) {
      return new WorkbookLocation.StoredWorkbook(Path.of(persistedPath));
    }
    return switch (source) {
      case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
          new WorkbookLocation.StoredWorkbook(
              Path.of(existingFile.path()).toAbsolutePath().normalize());
      case GridGrindRequest.WorkbookSource.New _ -> new WorkbookLocation.UnsavedWorkbook();
    };
  }

  private static String absolutePath(String path) {
    return Path.of(path).toAbsolutePath().normalize().toString();
  }

  private GridGrindResponse closeWorkbook(
      ExcelWorkbook workbook, GridGrindResponse response, GridGrindRequest request) {
    try {
      workbookCloser.close(workbook);
      return response;
    } catch (IOException exception) {
      return switch (response) {
        case GridGrindResponse.Failure failure ->
            new GridGrindResponse.Failure(
                failure.protocolVersion(),
                GridGrindProblems.appendCause(
                    failure.problem(),
                    GridGrindProblems.supplementalCause(
                        "EXECUTE_REQUEST",
                        exception,
                        "Workbook close failed after the primary problem")));
        case GridGrindResponse.Success _ ->
            new GridGrindResponse.Failure(
                response.protocolVersion(),
                problemFor(
                    exception,
                    new GridGrindResponse.ProblemContext.ExecuteRequest(
                        reqSourceType(request), reqPersistenceType(request))));
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
      return closeWorkbook(
          workbook,
          new GridGrindResponse.Failure(
              protocolVersion,
              problemFor(
                  exception,
                  new GridGrindResponse.ProblemContext.ExecuteRequest(
                      reqSourceType(request), reqPersistenceType(request)))),
          request);
    }
  }

  private String reqSourceType(GridGrindRequest request) {
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.New _ -> "NEW";
      case GridGrindRequest.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  private String reqPersistenceType(GridGrindRequest request) {
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

  static String persistencePath(
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

  private GridGrindResponse operationFailure(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      int operationIndex,
      WorkbookOperation operation,
      Exception exception) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        problemFor(
            exception,
            new GridGrindResponse.ProblemContext.ApplyOperation(
                reqSourceType(request),
                reqPersistenceType(request),
                operationIndex,
                operation.operationType(),
                sheetNameFor(operation, exception),
                addressFor(operation, exception),
                rangeFor(operation, exception),
                formulaFor(operation, exception),
                namedRangeNameFor(operation, exception))));
  }

  private GridGrindResponse readFailure(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      int readIndex,
      WorkbookReadOperation read,
      Exception exception) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        problemFor(
            exception,
            new GridGrindResponse.ProblemContext.ExecuteRead(
                reqSourceType(request),
                reqPersistenceType(request),
                readIndex,
                readType(read),
                read.requestId(),
                sheetNameFor(read),
                addressFor(read, exception),
                formulaFor(exception),
                namedRangeNameFor(read, exception))));
  }

  /** Returns the SCREAMING_SNAKE_CASE discriminator for one read operation. */
  static String readType(WorkbookReadOperation read) {
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadOperation.GetNamedRanges _ -> "GET_NAMED_RANGES";
      case WorkbookReadOperation.GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadOperation.GetCells _ -> "GET_CELLS";
      case WorkbookReadOperation.GetWindow _ -> "GET_WINDOW";
      case WorkbookReadOperation.GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case WorkbookReadOperation.GetHyperlinks _ -> "GET_HYPERLINKS";
      case WorkbookReadOperation.GetComments _ -> "GET_COMMENTS";
      case WorkbookReadOperation.GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadOperation.GetPrintLayout _ -> "GET_PRINT_LAYOUT";
      case WorkbookReadOperation.GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadOperation.GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadOperation.GetAutofilters _ -> "GET_AUTOFILTERS";
      case WorkbookReadOperation.GetTables _ -> "GET_TABLES";
      case WorkbookReadOperation.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadOperation.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadOperation.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> "ANALYZE_FORMULA_HEALTH";
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ ->
          "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case WorkbookReadOperation.AnalyzeTableHealth _ -> "ANALYZE_TABLE_HEALTH";
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> "ANALYZE_HYPERLINK_HEALTH";
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }

  /** Returns the sheet name associated with one read operation. */
  static String sheetNameFor(WorkbookReadOperation read) {
    return switch (read) {
      case WorkbookReadOperation.GetSheetSummary op -> op.sheetName();
      case WorkbookReadOperation.GetCells op -> op.sheetName();
      case WorkbookReadOperation.GetWindow op -> op.sheetName();
      case WorkbookReadOperation.GetMergedRegions op -> op.sheetName();
      case WorkbookReadOperation.GetHyperlinks op -> op.sheetName();
      case WorkbookReadOperation.GetComments op -> op.sheetName();
      case WorkbookReadOperation.GetSheetLayout op -> op.sheetName();
      case WorkbookReadOperation.GetPrintLayout op -> op.sheetName();
      case WorkbookReadOperation.GetDataValidations op -> op.sheetName();
      case WorkbookReadOperation.GetConditionalFormatting op -> op.sheetName();
      case WorkbookReadOperation.GetAutofilters op -> op.sheetName();
      case WorkbookReadOperation.GetSheetSchema op -> op.sheetName();
      case WorkbookReadOperation.GetFormulaSurface op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeFormulaHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeDataValidationHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth op ->
          singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeAutofilterHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeHyperlinkHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetTables _ -> null;
      case WorkbookReadOperation.GetNamedRangeSurface _ -> null;
      case WorkbookReadOperation.AnalyzeTableHealth _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> null;
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> null;
    };
  }

  /** Returns the address associated with one read operation or exception. */
  static String addressFor(WorkbookReadOperation read, Exception exception) {
    String fromException = addressFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (read) {
      case WorkbookReadOperation.GetWindow op -> op.topLeftAddress();
      case WorkbookReadOperation.GetSheetSchema op -> op.topLeftAddress();
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetSheetSummary _ -> null;
      case WorkbookReadOperation.GetCells _ -> null;
      case WorkbookReadOperation.GetMergedRegions _ -> null;
      case WorkbookReadOperation.GetHyperlinks _ -> null;
      case WorkbookReadOperation.GetComments _ -> null;
      case WorkbookReadOperation.GetSheetLayout _ -> null;
      case WorkbookReadOperation.GetPrintLayout _ -> null;
      case WorkbookReadOperation.GetDataValidations _ -> null;
      case WorkbookReadOperation.GetConditionalFormatting _ -> null;
      case WorkbookReadOperation.GetAutofilters _ -> null;
      case WorkbookReadOperation.GetTables _ -> null;
      case WorkbookReadOperation.GetFormulaSurface _ -> null;
      case WorkbookReadOperation.GetNamedRangeSurface _ -> null;
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> null;
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> null;
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ -> null;
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> null;
      case WorkbookReadOperation.AnalyzeTableHealth _ -> null;
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> null;
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> null;
    };
  }

  /** Returns the named-range name associated with one read operation or exception. */
  static String namedRangeNameFor(WorkbookReadOperation read, Exception exception) {
    String fromException = namedRangeNameFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetSheetSummary _ -> null;
      case WorkbookReadOperation.GetCells _ -> null;
      case WorkbookReadOperation.GetWindow _ -> null;
      case WorkbookReadOperation.GetMergedRegions _ -> null;
      case WorkbookReadOperation.GetHyperlinks _ -> null;
      case WorkbookReadOperation.GetComments _ -> null;
      case WorkbookReadOperation.GetSheetLayout _ -> null;
      case WorkbookReadOperation.GetPrintLayout _ -> null;
      case WorkbookReadOperation.GetDataValidations _ -> null;
      case WorkbookReadOperation.GetConditionalFormatting _ -> null;
      case WorkbookReadOperation.GetAutofilters _ -> null;
      case WorkbookReadOperation.GetTables _ -> null;
      case WorkbookReadOperation.GetFormulaSurface _ -> null;
      case WorkbookReadOperation.GetSheetSchema _ -> null;
      case WorkbookReadOperation.GetNamedRangeSurface op -> singleNamedRangeName(op.selection());
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> null;
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> null;
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ -> null;
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> null;
      case WorkbookReadOperation.AnalyzeTableHealth _ -> null;
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeHealth op -> singleNamedRangeName(op.selection());
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> null;
    };
  }

  private static String singleSheetName(SheetSelection selection) {
    return switch (selection) {
      case SheetSelection.All _ -> null;
      case SheetSelection.Selected selected ->
          selected.sheetNames().size() == 1 ? selected.sheetNames().getFirst() : null;
    };
  }

  private static String singleNamedRangeName(NamedRangeSelection selection) {
    return switch (selection) {
      case NamedRangeSelection.All _ -> null;
      case NamedRangeSelection.Selected selected ->
          selected.selectors().size() == 1
              ? switch (selected.selectors().getFirst()) {
                case NamedRangeSelector.ByName byName -> byName.name();
                case NamedRangeSelector.WorkbookScope workbookScope -> workbookScope.name();
                case NamedRangeSelector.SheetScope sheetScope -> sheetScope.name();
              }
              : null;
    };
  }

  /**
   * Returns the formula string for the failure context: checks the exception first, then falls back
   * to the write operation when it directly carries a formula.
   */
  static String formulaFor(WorkbookOperation operation, Exception exception) {
    String fromException = formulaFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (operation) {
      case WorkbookOperation.SetCell op ->
          switch (op.value()) {
            case CellInput.Formula formula -> formula.formula();
            case CellInput.Blank _ -> null;
            case CellInput.Text _ -> null;
            case CellInput.RichText _ -> null;
            case CellInput.Numeric _ -> null;
            case CellInput.BooleanValue _ -> null;
            case CellInput.Date _ -> null;
            case CellInput.DateTime _ -> null;
          };
      case WorkbookOperation.EnsureSheet _ -> null;
      case WorkbookOperation.RenameSheet _ -> null;
      case WorkbookOperation.DeleteSheet _ -> null;
      case WorkbookOperation.MoveSheet _ -> null;
      case WorkbookOperation.CopySheet _ -> null;
      case WorkbookOperation.SetActiveSheet _ -> null;
      case WorkbookOperation.SetSelectedSheets _ -> null;
      case WorkbookOperation.SetSheetVisibility _ -> null;
      case WorkbookOperation.SetSheetProtection _ -> null;
      case WorkbookOperation.ClearSheetProtection _ -> null;
      case WorkbookOperation.MergeCells _ -> null;
      case WorkbookOperation.UnmergeCells _ -> null;
      case WorkbookOperation.SetColumnWidth _ -> null;
      case WorkbookOperation.SetRowHeight _ -> null;
      case WorkbookOperation.InsertRows _ -> null;
      case WorkbookOperation.DeleteRows _ -> null;
      case WorkbookOperation.ShiftRows _ -> null;
      case WorkbookOperation.InsertColumns _ -> null;
      case WorkbookOperation.DeleteColumns _ -> null;
      case WorkbookOperation.ShiftColumns _ -> null;
      case WorkbookOperation.SetRowVisibility _ -> null;
      case WorkbookOperation.SetColumnVisibility _ -> null;
      case WorkbookOperation.GroupRows _ -> null;
      case WorkbookOperation.UngroupRows _ -> null;
      case WorkbookOperation.GroupColumns _ -> null;
      case WorkbookOperation.UngroupColumns _ -> null;
      case WorkbookOperation.SetSheetPane _ -> null;
      case WorkbookOperation.SetSheetZoom _ -> null;
      case WorkbookOperation.SetPrintLayout _ -> null;
      case WorkbookOperation.ClearPrintLayout _ -> null;
      case WorkbookOperation.SetRange _ -> null;
      case WorkbookOperation.ClearRange _ -> null;
      case WorkbookOperation.SetHyperlink _ -> null;
      case WorkbookOperation.ClearHyperlink _ -> null;
      case WorkbookOperation.SetComment _ -> null;
      case WorkbookOperation.ClearComment _ -> null;
      case WorkbookOperation.ApplyStyle _ -> null;
      case WorkbookOperation.SetDataValidation _ -> null;
      case WorkbookOperation.ClearDataValidations _ -> null;
      case WorkbookOperation.SetConditionalFormatting _ -> null;
      case WorkbookOperation.ClearConditionalFormatting _ -> null;
      case WorkbookOperation.SetAutofilter _ -> null;
      case WorkbookOperation.ClearAutofilter _ -> null;
      case WorkbookOperation.SetTable _ -> null;
      case WorkbookOperation.DeleteTable _ -> null;
      case WorkbookOperation.SetNamedRange _ -> null;
      case WorkbookOperation.DeleteNamedRange _ -> null;
      case WorkbookOperation.AppendRow _ -> null;
      case WorkbookOperation.AutoSizeColumns _ -> null;
      case WorkbookOperation.EvaluateFormulas _ -> null;
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /** Returns the sheet name associated with one write operation or exception. */
  static String sheetNameFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.EnsureSheet op -> op.sheetName();
          case WorkbookOperation.RenameSheet op -> op.sheetName();
          case WorkbookOperation.DeleteSheet op -> op.sheetName();
          case WorkbookOperation.MoveSheet op -> op.sheetName();
          case WorkbookOperation.CopySheet op -> op.sourceSheetName();
          case WorkbookOperation.SetActiveSheet op -> op.sheetName();
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility op -> op.sheetName();
          case WorkbookOperation.SetSheetProtection op -> op.sheetName();
          case WorkbookOperation.ClearSheetProtection op -> op.sheetName();
          case WorkbookOperation.MergeCells op -> op.sheetName();
          case WorkbookOperation.UnmergeCells op -> op.sheetName();
          case WorkbookOperation.SetColumnWidth op -> op.sheetName();
          case WorkbookOperation.SetRowHeight op -> op.sheetName();
          case WorkbookOperation.InsertRows op -> op.sheetName();
          case WorkbookOperation.DeleteRows op -> op.sheetName();
          case WorkbookOperation.ShiftRows op -> op.sheetName();
          case WorkbookOperation.InsertColumns op -> op.sheetName();
          case WorkbookOperation.DeleteColumns op -> op.sheetName();
          case WorkbookOperation.ShiftColumns op -> op.sheetName();
          case WorkbookOperation.SetRowVisibility op -> op.sheetName();
          case WorkbookOperation.SetColumnVisibility op -> op.sheetName();
          case WorkbookOperation.GroupRows op -> op.sheetName();
          case WorkbookOperation.UngroupRows op -> op.sheetName();
          case WorkbookOperation.GroupColumns op -> op.sheetName();
          case WorkbookOperation.UngroupColumns op -> op.sheetName();
          case WorkbookOperation.SetSheetPane op -> op.sheetName();
          case WorkbookOperation.SetSheetZoom op -> op.sheetName();
          case WorkbookOperation.SetPrintLayout op -> op.sheetName();
          case WorkbookOperation.ClearPrintLayout op -> op.sheetName();
          case WorkbookOperation.SetCell op -> op.sheetName();
          case WorkbookOperation.SetRange op -> op.sheetName();
          case WorkbookOperation.ClearRange op -> op.sheetName();
          case WorkbookOperation.SetHyperlink op -> op.sheetName();
          case WorkbookOperation.ClearHyperlink op -> op.sheetName();
          case WorkbookOperation.SetComment op -> op.sheetName();
          case WorkbookOperation.ClearComment op -> op.sheetName();
          case WorkbookOperation.ApplyStyle op -> op.sheetName();
          case WorkbookOperation.SetDataValidation op -> op.sheetName();
          case WorkbookOperation.ClearDataValidations op -> op.sheetName();
          case WorkbookOperation.SetConditionalFormatting op -> op.sheetName();
          case WorkbookOperation.ClearConditionalFormatting op -> op.sheetName();
          case WorkbookOperation.SetAutofilter op -> op.sheetName();
          case WorkbookOperation.ClearAutofilter op -> op.sheetName();
          case WorkbookOperation.SetTable op -> op.table().sheetName();
          case WorkbookOperation.DeleteTable op -> op.sheetName();
          case WorkbookOperation.SetNamedRange op -> op.target().sheetName();
          case WorkbookOperation.DeleteNamedRange op ->
              switch (op.scope()) {
                case NamedRangeScope.Workbook _ -> null;
                case NamedRangeScope.Sheet sheetScope -> sheetScope.sheetName();
              };
          case WorkbookOperation.AppendRow op -> op.sheetName();
          case WorkbookOperation.AutoSizeColumns op -> op.sheetName();
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : sheetNameFor(exception);
  }

  /** Returns the cell address associated with one write operation or exception. */
  static String addressFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.SetCell op -> op.address();
          case WorkbookOperation.SetHyperlink op -> op.address();
          case WorkbookOperation.ClearHyperlink op -> op.address();
          case WorkbookOperation.SetComment op -> op.address();
          case WorkbookOperation.ClearComment op -> op.address();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.CopySheet _ -> null;
          case WorkbookOperation.SetActiveSheet _ -> null;
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility _ -> null;
          case WorkbookOperation.SetSheetProtection _ -> null;
          case WorkbookOperation.ClearSheetProtection _ -> null;
          case WorkbookOperation.MergeCells _ -> null;
          case WorkbookOperation.UnmergeCells _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.InsertRows _ -> null;
          case WorkbookOperation.DeleteRows _ -> null;
          case WorkbookOperation.ShiftRows _ -> null;
          case WorkbookOperation.InsertColumns _ -> null;
          case WorkbookOperation.DeleteColumns _ -> null;
          case WorkbookOperation.ShiftColumns _ -> null;
          case WorkbookOperation.SetRowVisibility _ -> null;
          case WorkbookOperation.SetColumnVisibility _ -> null;
          case WorkbookOperation.GroupRows _ -> null;
          case WorkbookOperation.UngroupRows _ -> null;
          case WorkbookOperation.GroupColumns _ -> null;
          case WorkbookOperation.UngroupColumns _ -> null;
          case WorkbookOperation.SetSheetPane _ -> null;
          case WorkbookOperation.SetSheetZoom _ -> null;
          case WorkbookOperation.SetPrintLayout _ -> null;
          case WorkbookOperation.ClearPrintLayout _ -> null;
          case WorkbookOperation.SetRange _ -> null;
          case WorkbookOperation.ClearRange _ -> null;
          case WorkbookOperation.ApplyStyle _ -> null;
          case WorkbookOperation.SetDataValidation _ -> null;
          case WorkbookOperation.ClearDataValidations _ -> null;
          case WorkbookOperation.SetConditionalFormatting _ -> null;
          case WorkbookOperation.ClearConditionalFormatting _ -> null;
          case WorkbookOperation.SetAutofilter _ -> null;
          case WorkbookOperation.ClearAutofilter _ -> null;
          case WorkbookOperation.SetTable _ -> null;
          case WorkbookOperation.DeleteTable _ -> null;
          case WorkbookOperation.SetNamedRange _ -> null;
          case WorkbookOperation.DeleteNamedRange _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : addressFor(exception);
  }

  /** Returns the range string associated with one write operation or exception. */
  static String rangeFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.SetRange op -> op.range();
          case WorkbookOperation.ClearRange op -> op.range();
          case WorkbookOperation.ApplyStyle op -> op.range();
          case WorkbookOperation.SetDataValidation op -> op.range();
          case WorkbookOperation.SetConditionalFormatting op ->
              op.conditionalFormatting().ranges().size() == 1
                  ? op.conditionalFormatting().ranges().getFirst()
                  : null;
          case WorkbookOperation.SetAutofilter op -> op.range();
          case WorkbookOperation.SetTable op -> op.table().range();
          case WorkbookOperation.MergeCells op -> op.range();
          case WorkbookOperation.UnmergeCells op -> op.range();
          case WorkbookOperation.SetNamedRange op -> op.target().range();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.CopySheet _ -> null;
          case WorkbookOperation.SetActiveSheet _ -> null;
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility _ -> null;
          case WorkbookOperation.SetSheetProtection _ -> null;
          case WorkbookOperation.ClearSheetProtection _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.InsertRows _ -> null;
          case WorkbookOperation.DeleteRows _ -> null;
          case WorkbookOperation.ShiftRows _ -> null;
          case WorkbookOperation.InsertColumns _ -> null;
          case WorkbookOperation.DeleteColumns _ -> null;
          case WorkbookOperation.ShiftColumns _ -> null;
          case WorkbookOperation.SetRowVisibility _ -> null;
          case WorkbookOperation.SetColumnVisibility _ -> null;
          case WorkbookOperation.GroupRows _ -> null;
          case WorkbookOperation.UngroupRows _ -> null;
          case WorkbookOperation.GroupColumns _ -> null;
          case WorkbookOperation.UngroupColumns _ -> null;
          case WorkbookOperation.SetSheetPane _ -> null;
          case WorkbookOperation.SetSheetZoom _ -> null;
          case WorkbookOperation.SetPrintLayout _ -> null;
          case WorkbookOperation.ClearPrintLayout _ -> null;
          case WorkbookOperation.SetCell _ -> null;
          case WorkbookOperation.SetHyperlink _ -> null;
          case WorkbookOperation.ClearHyperlink _ -> null;
          case WorkbookOperation.SetComment _ -> null;
          case WorkbookOperation.ClearComment _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.ClearDataValidations _ -> null;
          case WorkbookOperation.ClearConditionalFormatting _ -> null;
          case WorkbookOperation.ClearAutofilter _ -> null;
          case WorkbookOperation.DeleteTable _ -> null;
          case WorkbookOperation.DeleteNamedRange _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : rangeFor(exception);
  }

  /** Returns the named-range identifier associated with one write operation or exception. */
  static String namedRangeNameFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.SetNamedRange op -> op.name();
          case WorkbookOperation.DeleteNamedRange op -> op.name();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.CopySheet _ -> null;
          case WorkbookOperation.SetActiveSheet _ -> null;
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility _ -> null;
          case WorkbookOperation.SetSheetProtection _ -> null;
          case WorkbookOperation.ClearSheetProtection _ -> null;
          case WorkbookOperation.MergeCells _ -> null;
          case WorkbookOperation.UnmergeCells _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.InsertRows _ -> null;
          case WorkbookOperation.DeleteRows _ -> null;
          case WorkbookOperation.ShiftRows _ -> null;
          case WorkbookOperation.InsertColumns _ -> null;
          case WorkbookOperation.DeleteColumns _ -> null;
          case WorkbookOperation.ShiftColumns _ -> null;
          case WorkbookOperation.SetRowVisibility _ -> null;
          case WorkbookOperation.SetColumnVisibility _ -> null;
          case WorkbookOperation.GroupRows _ -> null;
          case WorkbookOperation.UngroupRows _ -> null;
          case WorkbookOperation.GroupColumns _ -> null;
          case WorkbookOperation.UngroupColumns _ -> null;
          case WorkbookOperation.SetSheetPane _ -> null;
          case WorkbookOperation.SetSheetZoom _ -> null;
          case WorkbookOperation.SetPrintLayout _ -> null;
          case WorkbookOperation.ClearPrintLayout _ -> null;
          case WorkbookOperation.SetCell _ -> null;
          case WorkbookOperation.SetRange _ -> null;
          case WorkbookOperation.ClearRange _ -> null;
          case WorkbookOperation.SetHyperlink _ -> null;
          case WorkbookOperation.ClearHyperlink _ -> null;
          case WorkbookOperation.SetComment _ -> null;
          case WorkbookOperation.ClearComment _ -> null;
          case WorkbookOperation.ApplyStyle _ -> null;
          case WorkbookOperation.SetDataValidation _ -> null;
          case WorkbookOperation.ClearDataValidations _ -> null;
          case WorkbookOperation.SetConditionalFormatting _ -> null;
          case WorkbookOperation.ClearConditionalFormatting _ -> null;
          case WorkbookOperation.SetAutofilter _ -> null;
          case WorkbookOperation.ClearAutofilter _ -> null;
          case WorkbookOperation.SetTable _ -> null;
          case WorkbookOperation.DeleteTable _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : namedRangeNameFor(exception);
  }

  /** Functional interface for closing an ExcelWorkbook after request execution. */
  @FunctionalInterface
  interface WorkbookCloser {
    /** Closes the workbook, releasing any held resources. */
    void close(ExcelWorkbook workbook) throws IOException;
  }
}
