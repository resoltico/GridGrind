package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import dev.erst.gridgrind.excel.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.excel.WorkbookPasswordRequiredException;
import dev.erst.gridgrind.excel.WorkbookSecurityException;
import dev.erst.gridgrind.protocol.dto.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.json.InvalidJsonException;
import dev.erst.gridgrind.protocol.json.InvalidRequestException;
import dev.erst.gridgrind.protocol.json.InvalidRequestShapeException;
import dev.erst.gridgrind.protocol.json.PayloadException;
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
      case InvalidJsonException _ -> GridGrindProblemCode.INVALID_JSON;
      case InvalidRequestShapeException _ -> GridGrindProblemCode.INVALID_REQUEST_SHAPE;
      case InvalidRequestException _ -> GridGrindProblemCode.INVALID_REQUEST;
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
      case GridGrindResponse.ProblemContext.ApplyOperation _ -> context;
      case GridGrindResponse.ProblemContext.ExecuteRead _ -> context;
      case GridGrindResponse.ProblemContext.ParseArguments _ -> context;
      case GridGrindResponse.ProblemContext.ValidateRequest _ -> context;
      case GridGrindResponse.ProblemContext.OpenWorkbook _ -> context;
      case GridGrindResponse.ProblemContext.PersistWorkbook _ -> context;
      case GridGrindResponse.ProblemContext.ExecuteRequest _ -> context;
      case GridGrindResponse.ProblemContext.WriteResponse _ -> context;
    };
  }
}
