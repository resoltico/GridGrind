package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.Objects;
import java.util.function.Supplier;

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

  GridGrindResponse closeWorkbook(
      ExcelWorkbook workbook,
      GridGrindResponse response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      GridGrindProblemCode primaryFailureCode,
      Integer failedStepIndex,
      String failedStepId) {
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      workbookCloser.close(workbook);
      closePhase.succeed();
      return rebuildResponse(
          response, request, journal, primaryFailureCode, failedStepIndex, failedStepId);
    } catch (Exception closeFailure) {
      GridGrindResponse.Problem closeProblem =
          problemFor(
              closeFailure,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request)));
      closePhase.fail("failed (" + closeProblem.code() + ")");
      return appendCloseFailure(
          response,
          request,
          journal,
          primaryFailureCode,
          failedStepIndex,
          failedStepId,
          closeProblem);
    }
  }

  GridGrindResponse closeReadableWorkbook(
      ExcelOoxmlPackageSecuritySupport.ReadableWorkbook workbook,
      GridGrindResponse response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      GridGrindProblemCode primaryFailureCode,
      Integer failedStepIndex,
      String failedStepId) {
    if (workbook == null) {
      return response;
    }
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      readableWorkbookCloser.close(workbook);
      closePhase.succeed();
      return rebuildResponse(
          response, request, journal, primaryFailureCode, failedStepIndex, failedStepId);
    } catch (Exception closeFailure) {
      GridGrindResponse.Problem closeProblem =
          problemFor(
              closeFailure,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request)));
      closePhase.fail("failed (" + closeProblem.code() + ")");
      return appendCloseFailure(
          response,
          request,
          journal,
          primaryFailureCode,
          failedStepIndex,
          failedStepId,
          closeProblem);
    }
  }

  GridGrindResponse guardUnexpectedRuntime(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      Supplier<GridGrindResponse> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request)));
      return failureResponse(
          protocolVersion,
          journal,
          request.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
          problem,
          null,
          null);
    }
  }

  GridGrindResponse guardUnexpectedRuntime(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      ExcelWorkbook workbook,
      Supplier<GridGrindResponse> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request)));
      return closeWorkbook(
          workbook,
          failureResponse(
              protocolVersion,
              journal,
              request.steps().size(),
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

  static GridGrindResponse.ProblemContext enrichContext(
      GridGrindResponse.ProblemContext context, Throwable exception) {
    return GridGrindProblems.enrichContext(context, exception);
  }

  static GridGrindResponse.Problem problemFor(
      Throwable exception, GridGrindResponse.ProblemContext context) {
    return GridGrindProblems.fromException(exception, context);
  }

  static GridGrindResponse.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      int plannedStepCount,
      GridGrindResponse.Problem problem,
      Integer failedStepIndex,
      String failedStepId) {
    return failureResponse(
        protocolVersion,
        journal,
        plannedStepCount,
        CalculationReport.notRequested(),
        problem,
        failedStepIndex,
        failedStepId);
  }

  static GridGrindResponse.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      int plannedStepCount,
      CalculationReport calculation,
      GridGrindResponse.Problem problem,
      Integer failedStepIndex,
      String failedStepId) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        journal.buildFailure(plannedStepCount, problem.code(), failedStepIndex, failedStepId),
        calculation,
        problem);
  }

  private static GridGrindResponse rebuildResponse(
      GridGrindResponse response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      GridGrindProblemCode primaryFailureCode,
      Integer failedStepIndex,
      String failedStepId) {
    return switch (response) {
      case GridGrindResponse.Success success ->
          new GridGrindResponse.Success(
              success.protocolVersion(),
              journal.buildSuccess(request.steps().size()),
              success.calculation(),
              success.persistence(),
              success.warnings(),
              success.assertions(),
              success.inspections());
      case GridGrindResponse.Failure failure ->
          new GridGrindResponse.Failure(
              failure.protocolVersion(),
              journal.buildFailure(
                  request.steps().size(), primaryFailureCode, failedStepIndex, failedStepId),
              failure.calculation(),
              failure.problem());
    };
  }

  private static GridGrindResponse appendCloseFailure(
      GridGrindResponse response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      GridGrindProblemCode primaryFailureCode,
      Integer failedStepIndex,
      String failedStepId,
      GridGrindResponse.Problem closeProblem) {
    if (response instanceof GridGrindResponse.Failure existingFailure) {
      return new GridGrindResponse.Failure(
          existingFailure.protocolVersion(),
          journal.buildFailure(
              request.steps().size(), primaryFailureCode, failedStepIndex, failedStepId),
          existingFailure.calculation(),
          GridGrindProblems.appendCause(
              existingFailure.problem(), GridGrindProblems.problemCause(closeProblem)));
    }
    return new GridGrindResponse.Failure(
        request.protocolVersion(),
        journal.buildFailure(request.steps().size(), closeProblem.code(), null, null),
        response.calculation(),
        closeProblem);
  }
}
