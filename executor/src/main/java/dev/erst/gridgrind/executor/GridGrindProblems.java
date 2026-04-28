package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.PayloadException;
import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import dev.erst.gridgrind.excel.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.excel.MissingExternalWorkbookException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnregisteredUserDefinedFunctionException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookPasswordRequiredException;
import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.IOException;
import java.time.DateTimeException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Centralized problem construction and exception classification for protocol and transport
 * failures.
 */
public final class GridGrindProblems {
  private GridGrindProblems() {}

  /** Builds a fully populated problem from a classified exception. */
  public static GridGrindResponse.Problem fromException(
      Throwable exception, dev.erst.gridgrind.contract.dto.ProblemContext context) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(context, "context must not be null");
    return problem(
        codeFor(exception),
        messageFor(exception),
        enrichContext(context, exception),
        assertionFailureFor(exception),
        causesFor(exception, context.stage()));
  }

  /** Builds a fully populated problem from an explicit code and message. */
  public static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      Throwable cause) {
    Objects.requireNonNull(context, "context must not be null");
    return problem(
        code, message, context, assertionFailureFor(cause), causesFor(cause, context.stage()));
  }

  /**
   * Builds a fully populated problem from an explicit code, message, and already-structured causes.
   */
  public static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      List<GridGrindResponse.ProblemCause> causes) {
    return problem(code, message, context, null, causes);
  }

  private static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      AssertionFailure assertionFailure,
      List<GridGrindResponse.ProblemCause> causes) {
    Objects.requireNonNull(code, "code must not be null");
    Objects.requireNonNull(context, "context must not be null");
    return new GridGrindResponse.Problem(
        code,
        code.category(),
        code.recovery(),
        code.title(),
        Objects.requireNonNull(message, "message must not be null"),
        code.resolution(),
        context,
        Optional.ofNullable(assertionFailure),
        causes == null ? List.of() : List.copyOf(causes));
  }

  /** Appends an extra structured cause while preserving the primary classified problem. */
  public static GridGrindResponse.Problem appendCause(
      GridGrindResponse.Problem problem, GridGrindResponse.ProblemCause cause) {
    Objects.requireNonNull(problem, "problem must not be null");
    Objects.requireNonNull(cause, "cause must not be null");
    List<GridGrindResponse.ProblemCause> causes = new ArrayList<>(problem.causes());
    causes.add(cause);
    return new GridGrindResponse.Problem(
        problem.code(),
        problem.category(),
        problem.recovery(),
        problem.title(),
        problem.message(),
        problem.resolution(),
        problem.context(),
        problem.assertionFailure(),
        List.copyOf(causes));
  }

  /** Converts an exception into one supplemental cause entry for secondary-failure reporting. */
  public static GridGrindResponse.ProblemCause supplementalCause(
      String stage, Throwable exception, String messagePrefix) {
    Objects.requireNonNull(stage, "stage must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    String message =
        messagePrefix == null || messagePrefix.isBlank()
            ? messageFor(exception)
            : messagePrefix + ": " + messageFor(exception);
    return new GridGrindResponse.ProblemCause(codeFor(exception), message, stage);
  }

  /** Converts an already-built problem into a synthetic cause entry for fallback reporting. */
  public static GridGrindResponse.ProblemCause problemCause(GridGrindResponse.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return new GridGrindResponse.ProblemCause(
        problem.code(), problem.title() + ": " + problem.message(), problem.context().stage());
  }

  static GridGrindProblemCode codeFor(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    return switch (exception) {
      case InvalidJsonException _ -> GridGrindProblemCode.INVALID_JSON;
      case InvalidRequestShapeException _ -> GridGrindProblemCode.INVALID_REQUEST_SHAPE;
      case InvalidRequestException _ -> GridGrindProblemCode.INVALID_REQUEST;
      case AssertionFailedException _ -> GridGrindProblemCode.ASSERTION_FAILED;
      case WorkbookNotFoundException _ -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      case InputSourceNotFoundException _ -> GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND;
      case InputSourceUnavailableException _ -> GridGrindProblemCode.INPUT_SOURCE_UNAVAILABLE;
      case InputSourceReadException _ -> GridGrindProblemCode.INPUT_SOURCE_IO_ERROR;
      case SheetNotFoundException _ -> GridGrindProblemCode.SHEET_NOT_FOUND;
      case NamedRangeNotFoundException _ -> GridGrindProblemCode.NAMED_RANGE_NOT_FOUND;
      case CellNotFoundException _ -> GridGrindProblemCode.CELL_NOT_FOUND;
      case InvalidCellAddressException _ -> GridGrindProblemCode.INVALID_CELL_ADDRESS;
      case InvalidRangeAddressException _ -> GridGrindProblemCode.INVALID_RANGE_ADDRESS;
      case InvalidFormulaException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case MissingExternalWorkbookException _ -> GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK;
      case UnregisteredUserDefinedFunctionException _ ->
          GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION;
      case UnsupportedFormulaException _ -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
      case WorkbookPasswordRequiredException _ -> GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED;
      case InvalidWorkbookPasswordException _ -> GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD;
      case InvalidSigningConfigurationException _ ->
          GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION;
      case WorkbookSecurityException _ -> GridGrindProblemCode.WORKBOOK_SECURITY_ERROR;
      case IOException _ -> GridGrindProblemCode.IO_ERROR;
      case IllegalArgumentException _ -> GridGrindProblemCode.INVALID_REQUEST;
      case DateTimeException _ -> GridGrindProblemCode.INVALID_REQUEST;
      default -> GridGrindProblemCode.INTERNAL_ERROR;
    };
  }

  static String messageFor(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    String message = exception.getMessage();
    return message == null || message.isBlank() ? simpleName(exception) : message;
  }

  /**
   * Returns the public diagnostic entries for one failure without exposing raw throwable internals.
   */
  static List<GridGrindResponse.ProblemCause> causesFor(Throwable exception) {
    return causesFor(exception, "EXECUTE_REQUEST");
  }

  private static List<GridGrindResponse.ProblemCause> causesFor(Throwable exception, String stage) {
    if (exception == null) {
      return List.of();
    }
    return List.of(
        new GridGrindResponse.ProblemCause(codeFor(exception), messageFor(exception), stage));
  }

  private static AssertionFailure assertionFailureFor(Throwable exception) {
    return exception instanceof AssertionFailedException assertionFailedException
        ? assertionFailedException.assertionFailure()
        : null;
  }

  private static String simpleName(Throwable exception) {
    String simpleName = exception.getClass().getSimpleName();
    return simpleName.isBlank() ? exception.getClass().getName() : simpleName;
  }

  /**
   * Enriches the problem context with exception-specific fields when the exception type and context
   * type are paired for protocol parsing failures (e.g., PayloadException in a ReadRequest
   * context).
   */
  static dev.erst.gridgrind.contract.dto.ProblemContext enrichContext(
      dev.erst.gridgrind.contract.dto.ProblemContext context, Throwable exception) {
    return switch (context) {
      case dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest rc -> {
        if (exception instanceof PayloadException pe) {
          dev.erst.gridgrind.contract.dto.ProblemContext.JsonLocation jsonLocation =
              pe.jsonLine() == null || pe.jsonColumn() == null
                  ? dev.erst.gridgrind.contract.dto.ProblemContext.JsonLocation.unavailable()
                  : pe.jsonPath() == null
                      ? dev.erst.gridgrind.contract.dto.ProblemContext.JsonLocation.lineColumn(
                          pe.jsonLine(), pe.jsonColumn())
                      : dev.erst.gridgrind.contract.dto.ProblemContext.JsonLocation.located(
                          pe.jsonPath(), pe.jsonLine(), pe.jsonColumn());
          yield rc.withJson(jsonLocation);
        }
        yield context;
      }
      case dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation executeCalculation ->
          exception instanceof Exception resolved
              ? executeCalculation.withLocation(ExecutionDiagnosticFields.locationFor(resolved))
              : context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep executeStep ->
          exception instanceof Exception resolved
              ? executeStep.withLocation(ExecutionDiagnosticFields.locationFor(resolved))
              : context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments _ -> context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest _ -> context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs _ -> context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook _ -> context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook _ -> context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest _ -> context;
      case dev.erst.gridgrind.contract.dto.ProblemContext.WriteResponse _ -> context;
    };
  }
}
