package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import java.util.List;
import java.util.Objects;

/** Public problem-construction helpers for execution and doctor callers. */
public final class GridGrindProblems {
  private GridGrindProblems() {}

  /** Builds a fully populated problem from a classified exception. */
  public static GridGrindProblemDetail.Problem fromException(
      Throwable exception, ProblemContext context) {
    Objects.requireNonNull(context, "context must not be null");
    return dev.erst.gridgrind.engine.runtime.GridGrindProblems.fromException(exception, context);
  }

  /** Builds a fully populated problem from an explicit code and message. */
  public static GridGrindProblemDetail.Problem problem(
      GridGrindProblemCode code, String message, ProblemContext context, Throwable cause) {
    Objects.requireNonNull(context, "context must not be null");
    return dev.erst.gridgrind.engine.runtime.GridGrindProblems.problem(
        code, message, context, cause);
  }

  /**
   * Builds a fully populated problem from an explicit code, message, and already-structured causes.
   */
  public static GridGrindProblemDetail.Problem problem(
      GridGrindProblemCode code,
      String message,
      ProblemContext context,
      List<GridGrindProblemDetail.ProblemCause> causes) {
    Objects.requireNonNull(context, "context must not be null");
    return dev.erst.gridgrind.engine.runtime.GridGrindProblems.problem(
        code, message, context, causes);
  }

  /** Appends an extra structured cause while preserving the primary classified problem. */
  public static GridGrindProblemDetail.Problem appendCause(
      GridGrindProblemDetail.Problem problem, GridGrindProblemDetail.ProblemCause cause) {
    return dev.erst.gridgrind.engine.runtime.GridGrindProblems.appendCause(problem, cause);
  }

  /** Converts an exception into one supplemental cause entry for secondary-failure reporting. */
  public static GridGrindProblemDetail.ProblemCause problemCause(
      GridGrindProblemDetail.Problem problem) {
    return dev.erst.gridgrind.engine.runtime.GridGrindProblems.problemCause(problem);
  }
}
