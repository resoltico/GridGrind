package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Rebuilds one result after close telemetry is known without discarding execution artifacts. */
final class ExecutionResponseCloseFinalizer {
  private ExecutionResponseCloseFinalizer() {}

  static WorkbookResult withCompletedClose(
      WorkbookResult response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      @Nullable GridGrindProblemCode primaryFailureCode,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId) {
    return switch (response) {
      case WorkbookResult.Success success ->
          new WorkbookResult.Success(
              success.protocolVersion(),
              success.planId(),
              journal.buildSuccess(request.steps().size()),
              success.calculation(),
              success.persistence(),
              success.evidence(),
              success.warnings(),
              success.assertions(),
              success.inspections());
      case WorkbookResult.Failure failure ->
          new WorkbookResult.Failure(
              failure.protocolVersion(),
              failure.planId(),
              journal.buildFailure(
                  request.steps().size(),
                  Objects.requireNonNullElse(primaryFailureCode, failure.problem().code()),
                  failedStepIndex,
                  failedStepId),
              failure.calculation(),
              failure.persistence(),
              failure.evidence(),
              failure.warnings(),
              failure.assertions(),
              failure.inspections(),
              failure.problem());
    };
  }

  static WorkbookResult withCloseFailure(
      WorkbookResult response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      @Nullable GridGrindProblemCode primaryFailureCode,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId,
      GridGrindProblemDetail.Problem closeProblem) {
    GridGrindProblemCode code =
        response instanceof WorkbookResult.Failure failure
            ? Objects.requireNonNullElse(primaryFailureCode, failure.problem().code())
            : closeProblem.code();
    return dev.erst.gridgrind.core.result.WorkbookResults.withSecondaryFailure(
        response,
        closeProblem,
        journal.buildFailure(request.steps().size(), code, failedStepIndex, failedStepId));
  }
}
