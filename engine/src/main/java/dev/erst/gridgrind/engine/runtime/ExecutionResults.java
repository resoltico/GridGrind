package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResults;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactIo;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.Objects;
import java.util.function.Supplier;
import org.jspecify.annotations.Nullable;

/** Response assembly, runtime guarding, and close-phase handling for request execution. */
final class ExecutionResults {
  private final WorkbookCloser workbookCloser;
  private final ReadableWorkbookCloser readableWorkbookCloser;

  ExecutionResults(WorkbookCloser workbookCloser, ReadableWorkbookCloser readableWorkbookCloser) {
    this.workbookCloser = Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
    this.readableWorkbookCloser =
        Objects.requireNonNull(readableWorkbookCloser, "readableWorkbookCloser must not be null");
  }

  WorkbookResult closeWorkbook(
      XssfWorkbook workbook,
      WorkbookResult response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      @Nullable GridGrindProblemCode primaryFailureCode,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId) {
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      workbookCloser.close(workbook);
      closePhase.succeed();
      return ExecutionResponseCloseFinalizer.withCompletedClose(
          response, request, journal, primaryFailureCode, failedStepIndex, failedStepId);
    } catch (Exception closeFailure) {
      GridGrindProblemDetail.Problem closeProblem =
          GridGrindProblems.fromException(
              closeFailure,
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      closePhase.fail(closeProblem.code());
      return ExecutionResponseCloseFinalizer.withCloseFailure(
          response,
          request,
          journal,
          primaryFailureCode,
          failedStepIndex,
          failedStepId,
          closeProblem);
    }
  }

  WorkbookResult closeReadableWorkbook(
      WorkbookArtifactIo.@Nullable MaterializedWorkbook workbook,
      WorkbookResult response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      @Nullable GridGrindProblemCode primaryFailureCode,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId) {
    if (workbook == null) {
      return response;
    }
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      readableWorkbookCloser.close(workbook);
      closePhase.succeed();
      return ExecutionResponseCloseFinalizer.withCompletedClose(
          response, request, journal, primaryFailureCode, failedStepIndex, failedStepId);
    } catch (Exception closeFailure) {
      GridGrindProblemDetail.Problem closeProblem =
          GridGrindProblems.fromException(
              closeFailure,
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      closePhase.fail(closeProblem.code());
      return ExecutionResponseCloseFinalizer.withCloseFailure(
          response,
          request,
          journal,
          primaryFailureCode,
          failedStepIndex,
          failedStepId,
          closeProblem);
    }
  }

  WorkbookResult guardUnexpectedRuntime(
      ProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      Supplier<WorkbookResult> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return failureResponse(
          protocolVersion,
          journal,
          request,
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
          problem,
          null,
          null);
    }
  }

  WorkbookResult guardUnexpectedRuntime(
      ProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      XssfWorkbook workbook,
      Supplier<WorkbookResult> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return closeWorkbook(
          workbook,
          failureResponseWithoutPlanOutcomeEvent(
              new ExecutionFailure(
                  new ExecutionFailure.Context(
                      protocolVersion,
                      journal,
                      request,
                      CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy())),
                  ExecutionFailure.Artifacts.empty(),
                  new ExecutionFailure.Detail(problem, null, null))),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
  }

  static WorkbookResult.Failure failureResponse(
      ProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId) {
    return failureResponse(
        new ExecutionFailure(
            new ExecutionFailure.Context(
                protocolVersion, journal, request, CalculationReport.notRequested()),
            ExecutionFailure.Artifacts.empty(),
            new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId)));
  }

  static WorkbookResult.Failure failureResponse(
      ProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId) {
    return failureResponse(
        new ExecutionFailure(
            new ExecutionFailure.Context(protocolVersion, journal, request, calculation),
            ExecutionFailure.Artifacts.empty(),
            new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId)));
  }

  static WorkbookResult.Failure failureResponse(ExecutionFailure failure) {
    return failureResponse(
        failure, WorkbookResults.notAttemptedPersistenceOutcome(failure.context().request()), true);
  }

  static WorkbookResult.Failure failureResponseWithoutPlanOutcomeEvent(ExecutionFailure failure) {
    return failureResponse(
        failure,
        WorkbookResults.notAttemptedPersistenceOutcome(failure.context().request()),
        false);
  }

  /** Returns a failed response retaining a publisher-proven final-destination state. */
  static WorkbookResult.Failure failureResponse(
      ExecutionFailure failure,
      dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome persistence) {
    return failureResponse(failure, persistence, true);
  }

  /** Returns a failed response retaining a publisher-proven final-destination state. */
  static WorkbookResult.Failure failureResponseWithoutPlanOutcomeEvent(
      ExecutionFailure failure,
      dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome persistence) {
    return failureResponse(failure, persistence, false);
  }

  private static WorkbookResult.Failure failureResponse(
      ExecutionFailure failure,
      dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome persistence,
      boolean emitPlanOutcomeEvent) {
    ExecutionFailure.Context context = failure.context();
    ExecutionFailure.Detail detail = failure.detail();
    return new WorkbookResult.Failure(
        context.protocolVersion(),
        context.request().planId(),
        context
            .journal()
            .buildFailure(
                context.request().steps().size(),
                detail.problem().code(),
                detail.failedStepIndex(),
                detail.failedStepId(),
                emitPlanOutcomeEvent),
        context.calculation(),
        persistence,
        ExecutionEvidence.forExecution(
            context.request(),
            context.calculation(),
            failure.artifacts().planAssertions(),
            failure.artifacts().hostAssertions(),
            failure.artifacts().preservation(),
            false),
        failure.artifacts().warnings(),
        failure.artifacts().planAssertions(),
        failure.artifacts().inspections(),
        detail.problem());
  }
}
