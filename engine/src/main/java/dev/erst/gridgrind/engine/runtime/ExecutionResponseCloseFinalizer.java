package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.List;
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
      @Nullable String failedStepId) {
    return switch (response) {
      case WorkbookResult.Success success ->
          new WorkbookResult.Success(
              success.protocolVersion(),
              success.planId(),
              journal.buildSuccess(request.steps().size()),
              success.calculation(),
              success.persistence(),
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
      @Nullable String failedStepId,
      GridGrindProblemDetail.Problem closeProblem) {
    return switch (response) {
      case WorkbookResult.Failure existingFailure ->
          new WorkbookResult.Failure(
              existingFailure.protocolVersion(),
              existingFailure.planId(),
              journal.buildFailure(
                  request.steps().size(),
                  Objects.requireNonNullElse(primaryFailureCode, existingFailure.problem().code()),
                  failedStepIndex,
                  failedStepId),
              existingFailure.calculation(),
              existingFailure.persistence(),
              existingFailure.warnings(),
              existingFailure.assertions(),
              existingFailure.inspections(),
              GridGrindProblems.appendCause(
                  existingFailure.problem(), GridGrindProblems.problemCause(closeProblem)));
      case WorkbookResult.Success success ->
          new WorkbookResult.Failure(
              request.protocolVersion(),
              request.planId(),
              journal.buildFailure(request.steps().size(), closeProblem.code(), null, null),
              success.calculation(),
              success.persistence(),
              success.warnings(),
              List.of(),
              success.inspections(),
              closeProblem);
    };
  }
}
