package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.PayloadException;
import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.FormulaException;
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

/**
 * Centralized problem construction and exception classification for protocol and transport
 * failures.
 */
public final class GridGrindProblems {
  private GridGrindProblems() {}

  /** Builds a fully populated problem from a classified exception. */
  public static GridGrindResponse.Problem fromException(
      Throwable exception, GridGrindResponse.ProblemContext context) {
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
      GridGrindResponse.ProblemContext context,
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
      GridGrindResponse.ProblemContext context,
      List<GridGrindResponse.ProblemCause> causes) {
    return problem(code, message, context, null, causes);
  }

  private static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      GridGrindResponse.ProblemContext context,
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
        assertionFailure,
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
    return causesFor(exception, null);
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
  static GridGrindResponse.ProblemContext enrichContext(
      GridGrindResponse.ProblemContext context, Throwable exception) {
    return switch (context) {
      case GridGrindResponse.ProblemContext.ReadRequest rc -> {
        if (exception instanceof PayloadException pe) {
          yield rc.withJson(pe.jsonPath(), pe.jsonLine(), pe.jsonColumn());
        }
        yield context;
      }
      case GridGrindResponse.ProblemContext.ExecuteCalculation executeCalculation ->
          executeCalculation.withExceptionData(
              sheetNameFor(exception), addressFor(exception), formulaFor(exception));
      case GridGrindResponse.ProblemContext.ExecuteStep executeStep ->
          executeStep.withExceptionData(
              sheetNameFor(exception),
              addressFor(exception),
              rangeFor(exception),
              formulaFor(exception),
              namedRangeNameFor(exception));
      case GridGrindResponse.ProblemContext.ParseArguments _ -> context;
      case GridGrindResponse.ProblemContext.ValidateRequest _ -> context;
      case GridGrindResponse.ProblemContext.ResolveInputs _ -> context;
      case GridGrindResponse.ProblemContext.OpenWorkbook _ -> context;
      case GridGrindResponse.ProblemContext.PersistWorkbook _ -> context;
      case GridGrindResponse.ProblemContext.ExecuteRequest _ -> context;
      case GridGrindResponse.ProblemContext.WriteResponse _ -> context;
    };
  }

  private static String sheetNameFor(Throwable exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.sheetName();
      case SheetNotFoundException sheetNotFoundException -> sheetNotFoundException.sheetName();
      case NamedRangeNotFoundException namedRangeNotFoundException ->
          switch (namedRangeNotFoundException.scope()) {
            case dev.erst.gridgrind.excel.ExcelNamedRangeScope.WorkbookScope _ -> null;
            case dev.erst.gridgrind.excel.ExcelNamedRangeScope.SheetScope sheetScope ->
                sheetScope.sheetName();
          };
      default -> null;
    };
  }

  private static String addressFor(Throwable exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.address();
      case CellNotFoundException cellNotFoundException -> cellNotFoundException.address();
      case InvalidCellAddressException invalidCellAddressException ->
          invalidCellAddressException.address();
      default -> null;
    };
  }

  private static String rangeFor(Throwable exception) {
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return invalidRangeAddressException.range();
    }
    return null;
  }

  private static String formulaFor(Throwable exception) {
    if (exception instanceof FormulaException formulaException) {
      return formulaException.formula();
    }
    return null;
  }

  private static String namedRangeNameFor(Throwable exception) {
    if (exception instanceof NamedRangeNotFoundException namedRangeNotFoundException) {
      return namedRangeNotFoundException.name();
    }
    return null;
  }
}
