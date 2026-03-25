package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.FormulaException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import java.io.IOException;
import java.time.DateTimeException;
import java.util.ArrayList;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Objects;
import java.util.Set;

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
        causesFor(exception));
  }

  /** Builds a fully populated problem from an explicit code and message. */
  public static GridGrindResponse.Problem problem(
      GridGrindProblemCode code,
      String message,
      GridGrindResponse.ProblemContext context,
      Throwable cause) {
    return problem(code, message, context, causesFor(cause));
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
    return new GridGrindResponse.ProblemCause(
        simpleName(exception), exception.getClass().getName(), message, stage);
  }

  /** Converts an already-built problem into a synthetic cause entry for fallback reporting. */
  public static GridGrindResponse.ProblemCause problemCause(GridGrindResponse.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return new GridGrindResponse.ProblemCause(
        "GridGrindProblem",
        problem.code().name(),
        problem.title() + ": " + problem.message(),
        problem.context().stage());
  }

  static GridGrindProblemCode codeFor(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    return switch (exception) {
      case WorkbookNotFoundException _ -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      case SheetNotFoundException _ -> GridGrindProblemCode.SHEET_NOT_FOUND;
      case CellNotFoundException _ -> GridGrindProblemCode.CELL_NOT_FOUND;
      case InvalidCellAddressException _ -> GridGrindProblemCode.INVALID_CELL_ADDRESS;
      case InvalidRangeAddressException _ -> GridGrindProblemCode.INVALID_RANGE_ADDRESS;
      case UnsupportedFormulaException _ -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
      case InvalidFormulaException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case InvalidJsonException _ -> GridGrindProblemCode.INVALID_JSON;
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

  static List<GridGrindResponse.ProblemCause> causesFor(Throwable exception) {
    if (exception == null) {
      return List.of();
    }

    List<GridGrindResponse.ProblemCause> causes = new ArrayList<>();
    Set<Throwable> seen = java.util.Collections.newSetFromMap(new IdentityHashMap<>());
    Throwable current = exception;
    while (current != null && seen.add(current)) {
      causes.add(
          new GridGrindResponse.ProblemCause(
              simpleName(current), current.getClass().getName(), messageFor(current), null));
      current = current.getCause();
    }
    return List.copyOf(causes);
  }

  static String formulaFor(Throwable exception) {
    return exception instanceof FormulaException fe ? fe.formula() : null;
  }

  static String sheetNameFor(Throwable exception) {
    return exception instanceof FormulaException fe ? fe.sheetName() : null;
  }

  static String addressFor(Throwable exception) {
    return exception instanceof FormulaException fe ? fe.address() : null;
  }

  static String rangeFor(Throwable exception) {
    return exception instanceof InvalidRangeAddressException ire ? ire.range() : null;
  }

  private static String simpleName(Throwable exception) {
    String simpleName = exception.getClass().getSimpleName();
    return simpleName.isBlank() ? exception.getClass().getName() : simpleName;
  }

  private static GridGrindResponse.ProblemContext enrichContext(
      GridGrindResponse.ProblemContext context, Throwable exception) {
    return switch (context) {
      case GridGrindResponse.ProblemContext.ReadRequest rc
          when exception instanceof PayloadException pe ->
          rc.withJson(pe.jsonPath(), pe.jsonLine(), pe.jsonColumn());
      case GridGrindResponse.ProblemContext.ApplyOperation ac
          when exception instanceof FormulaException fe ->
          ac.withExceptionData(fe.sheetName(), fe.address(), null, fe.formula());
      case GridGrindResponse.ProblemContext.ApplyOperation ac
          when exception instanceof InvalidRangeAddressException ire ->
          ac.withExceptionData(null, null, ire.range(), null);
      case GridGrindResponse.ProblemContext.AnalyzeWorkbook aw
          when exception instanceof FormulaException fe ->
          aw.withExceptionData(fe.sheetName(), fe.address(), fe.formula());
      default -> context;
    };
  }
}
