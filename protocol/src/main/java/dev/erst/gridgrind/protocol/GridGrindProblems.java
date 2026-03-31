package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.FormulaException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import java.io.IOException;
import java.time.DateTimeException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Centralized problem construction and exception classification for all protocol surfaces. */
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
        causesFor(exception, context.stage()));
  }

  /** Builds a fully populated problem from an explicit code and message. */
  public static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      GridGrindResponse.ProblemContext context,
      Throwable cause) {
    Objects.requireNonNull(context, "context must not be null");
    return problem(code, message, context, causesFor(cause, context.stage()));
  }

  /**
   * Builds a fully populated problem from an explicit code, message, and already-structured causes.
   */
  public static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      GridGrindResponse.ProblemContext context,
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
      case WorkbookNotFoundException _ -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      case SheetNotFoundException _ -> GridGrindProblemCode.SHEET_NOT_FOUND;
      case NamedRangeNotFoundException _ -> GridGrindProblemCode.NAMED_RANGE_NOT_FOUND;
      case CellNotFoundException _ -> GridGrindProblemCode.CELL_NOT_FOUND;
      case InvalidCellAddressException _ -> GridGrindProblemCode.INVALID_CELL_ADDRESS;
      case InvalidRangeAddressException _ -> GridGrindProblemCode.INVALID_RANGE_ADDRESS;
      case UnsupportedFormulaException _ -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
      case InvalidFormulaException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case InvalidJsonException _ -> GridGrindProblemCode.INVALID_JSON;
      case InvalidRequestShapeException _ -> GridGrindProblemCode.INVALID_REQUEST_SHAPE;
      case InvalidRequestException _ -> GridGrindProblemCode.INVALID_REQUEST;
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

  static String formulaFor(Throwable exception) {
    return exception instanceof FormulaException fe ? fe.formula() : null;
  }

  static String sheetNameFor(Throwable exception) {
    return exception instanceof FormulaException fe ? fe.sheetName() : null;
  }

  static String addressFor(Throwable exception) {
    return switch (exception) {
      case FormulaException fe -> fe.address();
      case CellNotFoundException cnf -> cnf.address();
      case InvalidCellAddressException ica -> ica.address();
      default -> null;
    };
  }

  static String rangeFor(Throwable exception) {
    return exception instanceof InvalidRangeAddressException ire ? ire.range() : null;
  }

  static String namedRangeNameFor(Throwable exception) {
    return exception instanceof NamedRangeNotFoundException nnf ? nnf.name() : null;
  }

  private static List<GridGrindResponse.ProblemCause> causesFor(Throwable exception, String stage) {
    if (exception == null) {
      return List.of();
    }
    return List.of(
        new GridGrindResponse.ProblemCause(codeFor(exception), messageFor(exception), stage));
  }

  private static String simpleName(Throwable exception) {
    String simpleName = exception.getClass().getSimpleName();
    return simpleName.isBlank() ? exception.getClass().getName() : simpleName;
  }

  /**
   * Enriches the problem context with exception-specific fields when the exception type and context
   * type are paired (e.g., PayloadException in a ReadRequest context).
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
      case GridGrindResponse.ProblemContext.ApplyOperation ac -> {
        if (exception instanceof FormulaException fe) {
          yield ac.withExceptionData(
              fe.sheetName(), fe.address(), null, fe.formula(), namedRangeNameFor(exception));
        }
        if (exception instanceof InvalidRangeAddressException ire) {
          yield ac.withExceptionData(null, null, ire.range(), null, namedRangeNameFor(exception));
        }
        if (exception instanceof NamedRangeNotFoundException) {
          yield ac.withExceptionData(null, null, null, null, namedRangeNameFor(exception));
        }
        yield context;
      }
      case GridGrindResponse.ProblemContext.ExecuteRead executeRead -> {
        if (exception instanceof FormulaException fe) {
          yield executeRead.withExceptionData(
              fe.sheetName(), fe.address(), fe.formula(), namedRangeNameFor(exception));
        }
        if (exception instanceof NamedRangeNotFoundException) {
          yield executeRead.withExceptionData(null, null, null, namedRangeNameFor(exception));
        }
        yield context;
      }
      case GridGrindResponse.ProblemContext.ParseArguments _ -> context;
      case GridGrindResponse.ProblemContext.ValidateRequest _ -> context;
      case GridGrindResponse.ProblemContext.OpenWorkbook _ -> context;
      case GridGrindResponse.ProblemContext.PersistWorkbook _ -> context;
      case GridGrindResponse.ProblemContext.ExecuteRequest _ -> context;
      case GridGrindResponse.ProblemContext.WriteResponse _ -> context;
    };
  }
}
