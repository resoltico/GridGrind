package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import java.util.List;
import java.util.Objects;
import java.util.function.Supplier;
import org.jspecify.annotations.Nullable;

/** Response assembly, runtime guarding, and close-phase handling for request execution. */
final class ExecutionResponseSupport {
  private final WorkbookCloser workbookCloser;
  private final ReadableWorkbookCloser readableWorkbookCloser;

  ExecutionResponseSupport(
      WorkbookCloser workbookCloser, ReadableWorkbookCloser readableWorkbookCloser) {
    this.workbookCloser = Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
    this.readableWorkbookCloser =
        Objects.requireNonNull(readableWorkbookCloser, "readableWorkbookCloser must not be null");
  }

  WorkbookResult closeWorkbook(
      ExcelWorkbook workbook,
      WorkbookResult response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      @Nullable GridGrindProblemCode primaryFailureCode,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      workbookCloser.close(workbook);
      closePhase.succeed();
      return ExecutionResponseCloseFinalizer.withCompletedClose(
          response, request, journal, primaryFailureCode, failedStepIndex, failedStepId);
    } catch (Exception closeFailure) {
      GridGrindProblemDetail.Problem closeProblem =
          problemFor(
              closeFailure,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      closePhase.fail("failed (" + closeProblem.code() + ")");
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
      @Nullable String failedStepId) {
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
          problemFor(
              closeFailure,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      closePhase.fail("failed (" + closeProblem.code() + ")");
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
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      Supplier<WorkbookResult> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindProblemDetail.Problem problem =
          problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
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
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      ExcelWorkbook workbook,
      Supplier<WorkbookResult> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindProblemDetail.Problem problem =
          problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return closeWorkbook(
          workbook,
          failureResponseWithoutPlanOutcomeEvent(
              protocolVersion,
              journal,
              request,
              CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
              problem,
              null,
              null),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
  }

  static GridGrindProblemCode problemCodeFor(Throwable exception) {
    return GridGrindProblems.codeFor(exception);
  }

  static dev.erst.gridgrind.contract.dto.ProblemContext enrichContext(
      dev.erst.gridgrind.contract.dto.ProblemContext context, Throwable exception) {
    return GridGrindProblems.enrichContext(context, exception);
  }

  static GridGrindProblemDetail.Problem problemFor(
      Throwable exception, dev.erst.gridgrind.contract.dto.ProblemContext context) {
    return GridGrindProblems.fromException(exception, context);
  }

  static WorkbookResult.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        CalculationReport.notRequested(),
        List.of(),
        List.of(),
        List.of(),
        problem,
        failedStepIndex,
        failedStepId,
        true);
  }

  static WorkbookResult.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        calculation,
        List.of(),
        List.of(),
        List.of(),
        problem,
        failedStepIndex,
        failedStepId,
        true);
  }

  static WorkbookResult.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      List<AssertionResult> assertions,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        calculation,
        List.of(),
        assertions,
        List.of(),
        problem,
        failedStepIndex,
        failedStepId,
        true);
  }

  static WorkbookResult.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        calculation,
        warnings,
        assertions,
        inspections,
        problem,
        failedStepIndex,
        failedStepId,
        true);
  }

  static WorkbookResult.Failure failureResponseWithoutPlanOutcomeEvent(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        calculation,
        List.of(),
        List.of(),
        List.of(),
        problem,
        failedStepIndex,
        failedStepId,
        false);
  }

  static WorkbookResult.Failure failureResponseWithoutPlanOutcomeEvent(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      List<AssertionResult> assertions,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        calculation,
        List.of(),
        assertions,
        List.of(),
        problem,
        failedStepIndex,
        failedStepId,
        false);
  }

  static WorkbookResult.Failure failureResponseWithoutPlanOutcomeEvent(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        request,
        calculation,
        warnings,
        assertions,
        inspections,
        problem,
        failedStepIndex,
        failedStepId,
        false);
  }

  private static WorkbookResult.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      WorkbookPlan request,
      CalculationReport calculation,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId,
      boolean emitPlanOutcomeEvent) {
    return new WorkbookResult.Failure(
        protocolVersion,
        request.planId(),
        journal.buildFailure(
            request.steps().size(),
            problem.code(),
            failedStepIndex,
            failedStepId,
            emitPlanOutcomeEvent),
        calculation,
        ExecutionRequestPaths.unwrittenPersistenceOutcome(request),
        warnings,
        assertions,
        inspections,
        problem);
  }
}
