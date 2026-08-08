package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.json.PayloadException;
import dev.erst.gridgrind.contract.json.PayloadLocation;
import dev.erst.gridgrind.contract.json.RequestProblemSource;
import java.nio.file.FileAlreadyExistsException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/**
 * Centralized problem construction and exception classification for protocol and transport
 * failures.
 */
public final class GridGrindProblems {
  private GridGrindProblems() {}

  /** Builds a fully populated problem from a classified exception. */
  public static GridGrindProblemDetail.Problem fromException(
      Throwable exception, dev.erst.gridgrind.contract.dto.ProblemContext context) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(context, "context must not be null");
    GridGrindProblemCode code = codeFor(exception);
    String publicMessage = publicMessageFor(code, exception, context);
    return problem(
        code,
        publicMessage,
        enrichContext(context, exception),
        Optional.ofNullable(assertionFailureFor(exception)),
        causesFor(exception, context.stage(), publicMessage),
        exception);
  }

  /** Builds a fully populated problem from an explicit code and message. */
  public static GridGrindProblemDetail.Problem problem(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      @Nullable Throwable cause) {
    Objects.requireNonNull(context, "context must not be null");
    String publicMessage =
        code == GridGrindProblemCode.INTERNAL_ERROR
            ? code.title()
            : Objects.requireNonNull(message);
    return problem(
        code,
        publicMessage,
        context,
        Optional.ofNullable(assertionFailureFor(cause)),
        causesFor(
            cause,
            context.stage(),
            cause == null ? publicMessage : publicMessageFor(codeFor(cause), cause, context)),
        cause);
  }

  /**
   * Builds a fully populated problem from an explicit code, message, and already-structured causes.
   */
  public static GridGrindProblemDetail.Problem problem(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      List<GridGrindProblemDetail.ProblemCause> causes) {
    return problem(code, message, context, Optional.empty(), causes, null);
  }

  private static GridGrindProblemDetail.Problem problem(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      Optional<AssertionFailure> assertionFailure,
      List<GridGrindProblemDetail.ProblemCause> causes,
      @Nullable Throwable source) {
    Objects.requireNonNull(code, "code must not be null");
    Objects.requireNonNull(context, "context must not be null");
    return new GridGrindProblemDetail.Problem(
        code,
        code.category(),
        code.recovery(),
        code.title(),
        Objects.requireNonNull(message, "message must not be null"),
        resolutionFor(code, message, context, source),
        context,
        Objects.requireNonNull(assertionFailure, "assertionFailure must not be null"),
        List.copyOf(Objects.requireNonNull(causes, "causes must not be null")));
  }

  private static String resolutionFor(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      @Nullable Throwable source) {
    if (source instanceof RequestProblemSource requestProblemSource) {
      return GridGrindRequestProblemSupport.resolution(
          requestProblemSource.requestProblem(), context);
    }
    return code.resolutionFor(message, context);
  }

  /** Appends an extra structured cause while preserving the primary classified problem. */
  public static GridGrindProblemDetail.Problem appendCause(
      GridGrindProblemDetail.Problem problem, GridGrindProblemDetail.ProblemCause cause) {
    Objects.requireNonNull(problem, "problem must not be null");
    Objects.requireNonNull(cause, "cause must not be null");
    List<GridGrindProblemDetail.ProblemCause> causes = new ArrayList<>(problem.causes());
    causes.add(cause);
    return new GridGrindProblemDetail.Problem(
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
  public static GridGrindProblemDetail.ProblemCause supplementalCause(
      String stage, Throwable exception, String messagePrefix) {
    Objects.requireNonNull(stage, "stage must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    GridGrindProblemCode code = codeFor(exception);
    String publicMessage = publicMessageFor(code, exception);
    String message =
        messagePrefix == null || messagePrefix.isBlank()
            ? publicMessage
            : messagePrefix + ": " + publicMessage;
    return new GridGrindProblemDetail.ProblemCause(code, message, stage);
  }

  /** Converts an already-built problem into a synthetic cause entry for fallback reporting. */
  public static GridGrindProblemDetail.ProblemCause problemCause(
      GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return new GridGrindProblemDetail.ProblemCause(
        problem.code(), problem.title() + ": " + problem.message(), problem.context().stage());
  }

  static GridGrindProblemCode codeFor(Throwable exception) {
    return GridGrindProblemCodeClassifier.codeFor(exception);
  }

  static String messageFor(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    String message = exception.getMessage();
    return message == null || message.isBlank() ? simpleName(exception) : message;
  }

  static String messageFor(
      Throwable exception, dev.erst.gridgrind.contract.dto.ProblemContext context) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(context, "context must not be null");
    if (exception instanceof FileAlreadyExistsException) {
      if (context instanceof ProblemContext.PersistWorkbook persistWorkbook) {
        return persistWorkbook
            .persistencePath()
            .map(
                persistencePath ->
                    "Could not write workbook to "
                        + persistencePath
                        + ": already exists; SAVE_AS.ifExists=REJECT requires a new"
                        + " destination path. Use ifExists=REPLACE to allow"
                        + " create-or-replace.")
            .orElseGet(() -> messageFor(exception));
      }
      return messageFor(exception);
    }
    return messageFor(exception);
  }

  private static String publicMessageFor(
      GridGrindProblemCode code,
      Throwable exception,
      dev.erst.gridgrind.contract.dto.ProblemContext context) {
    return code == GridGrindProblemCode.INTERNAL_ERROR
        ? code.title()
        : messageFor(exception, context);
  }

  private static String publicMessageFor(GridGrindProblemCode code, Throwable exception) {
    return code == GridGrindProblemCode.INTERNAL_ERROR ? code.title() : messageFor(exception);
  }

  /**
   * Returns the public diagnostic entries for one failure without exposing raw throwable internals.
   */
  static List<GridGrindProblemDetail.ProblemCause> causesFor(Throwable exception) {
    if (exception == null) {
      return List.of();
    }
    GridGrindProblemCode code = codeFor(exception);
    return causesFor(exception, "EXECUTE_REQUEST", publicMessageFor(code, exception));
  }

  private static List<GridGrindProblemDetail.ProblemCause> causesFor(
      @Nullable Throwable exception, String stage, String publicMessage) {
    if (exception == null) {
      return List.of();
    }
    return List.of(
        new GridGrindProblemDetail.ProblemCause(codeFor(exception), publicMessage, stage));
  }

  private static @Nullable AssertionFailure assertionFailureFor(@Nullable Throwable exception) {
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
          dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation jsonLocation =
              switch (pe.jsonLocation()) {
                case PayloadLocation.PathOnly pathOnly ->
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                        .pathOnly(pathOnly.jsonPathValue());
                case PayloadLocation.LineColumn lineColumn ->
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                        .lineColumn(lineColumn.jsonLineValue(), lineColumn.jsonColumnValue());
                case PayloadLocation.Located located ->
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                        .located(
                            located.jsonPathValue(),
                            located.jsonLineValue(),
                            located.jsonColumnValue());
                case PayloadLocation.Unavailable _ ->
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                        .unavailable();
              };
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
