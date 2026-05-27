package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Request workflow implementations for workbook, direct-event, and streaming execution modes. */
final class ExecutionWorkflowSupport {
  private final ExecutionWorkbookSupport workbookSupport;
  private final ExecutionCalculationSupport calculationSupport;
  private final ExecutionStepSupport stepSupport;
  private final ExecutionResponseSupport responseSupport;
  private final ExecutionDirectEventReadWorkflow directEventReadWorkflow;
  private final ExecutionStreamingWorkflow streamingWorkflow;

  ExecutionWorkflowSupport(
      ExecutionWorkbookSupport workbookSupport,
      ExecutionCalculationSupport calculationSupport,
      ExecutionStepSupport stepSupport,
      ExecutionResponseSupport responseSupport,
      TempFileFactory tempFileFactory) {
    this.workbookSupport =
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null");
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
    TempFileFactory requiredTempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    this.directEventReadWorkflow =
        new ExecutionDirectEventReadWorkflow(stepSupport, responseSupport, requiredTempFileFactory);
    this.streamingWorkflow =
        new ExecutionStreamingWorkflow(
            workbookSupport, calculationSupport, stepSupport, requiredTempFileFactory);
  }

  GridGrindResponse executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      Path workingDirectory) {
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), workingDirectory);
    List<AssertionResult> assertions = new ArrayList<>();
    List<InspectionResult> inspections = new ArrayList<>();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    boolean calculationExecuted = false;
    WorkbookExecutionContext executionContext =
        new WorkbookExecutionContext(protocolVersion, request, workbook, journal);

    for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
      WorkbookStep step = request.steps().get(stepIndex);
      CalculationCheckpoint calculationCheckpoint =
          executeCalculationBeforeStepIfNeeded(
              protocolVersion, request, workbook, journal, step, calculation, calculationExecuted);
      calculation = calculationCheckpoint.report();
      calculationExecuted = calculationCheckpoint.executed();
      if (calculationCheckpoint.failureResponse() != null) {
        return calculationCheckpoint.failureResponse();
      }

      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
      try {
        executeWorkbookStep(
            workbook, workbookLocation, executionMode, assertions, inspections, step);
        stepHandle.succeed();
      } catch (Exception exception) {
        return closeFailedStepExecution(
            executionContext, calculation, assertions, stepIndex, step, stepHandle, exception);
      }
    }

    if (!calculationExecuted) {
      ExecutionCalculationSupport.CalculationExecutionOutcome calculationOutcome =
          calculationSupport.executeCalculationPolicy(workbook, request, journal);
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure().isPresent()) {
        GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
        return responseSupport.closeWorkbook(
            workbook,
            ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                protocolVersion, journal, request.steps().size(), calculation, problem, null, null),
            request,
            journal,
            problem.code(),
            null,
            null);
      }
    }

    PersistenceResult persistenceResult =
        persistWorkbook(protocolVersion, request, workbook, journal, workingDirectory, calculation);
    if (persistenceResult.failureResponse() != null) {
      return persistenceResult.failureResponse();
    }

    return responseSupport.closeWorkbook(
        workbook,
        new GridGrindResponse.Success(
            protocolVersion,
            journal.buildSuccess(request.steps().size(), false),
            calculation,
            Objects.requireNonNull(
                persistenceResult.persistence(), "persistence must exist on success"),
            warnings,
            List.copyOf(assertions),
            List.copyOf(inspections)),
        request,
        journal,
        null,
        null,
        null);
  }

  private CalculationCheckpoint executeCalculationBeforeStepIfNeeded(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionJournalRecorder journal,
      WorkbookStep step,
      CalculationReport currentCalculation,
      boolean calculationExecuted) {
    if (calculationExecuted || !shouldExecuteCalculationBeforeStep(request, step)) {
      return new CalculationCheckpoint(currentCalculation, calculationExecuted, null);
    }
    ExecutionCalculationSupport.CalculationExecutionOutcome outcome =
        calculationSupport.executeCalculationPolicy(workbook, request, journal);
    if (outcome.failure().isEmpty()) {
      return new CalculationCheckpoint(outcome.report(), true, null);
    }
    GridGrindProblemDetail.Problem problem = outcome.failure().orElseThrow();
    return new CalculationCheckpoint(
        outcome.report(),
        true,
        responseSupport.closeWorkbook(
            workbook,
            ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                protocolVersion,
                journal,
                request.steps().size(),
                outcome.report(),
                problem,
                null,
                null),
            request,
            journal,
            problem.code(),
            null,
            null));
  }

  private void executeWorkbookStep(
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections,
      WorkbookStep step)
      throws IOException, AssertionFailedException {
    switch (step) {
      case MutationStep mutationStep -> stepSupport.executeMutationStep(workbook, mutationStep);
      case AssertionStep assertionStep ->
          assertions.add(
              stepSupport.executeAssertionStep(
                  assertionStep, workbook, workbookLocation, executionMode));
      case InspectionStep inspectionStep ->
          inspections.add(
              stepSupport.executeInspectionStep(
                  inspectionStep, workbook, workbookLocation, executionMode));
    }
  }

  private GridGrindResponse closeFailedStepExecution(
      WorkbookExecutionContext executionContext,
      CalculationReport calculation,
      List<AssertionResult> assertions,
      int stepIndex,
      WorkbookStep step,
      ExecutionJournalRecorder.StepHandle stepHandle,
      Exception exception) {
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            exception,
            stepSupport.executeStepContext(executionContext.request(), stepIndex, step, exception));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().stage(), problem.message());
    if (exception instanceof AssertionFailedException assertionFailed) {
      assertions.add(
          new AssertionResult(
              dev.erst.gridgrind.contract.assertion.AssertionOutcome.FAILED,
              assertionFailed.assertionFailure().stepId(),
              assertionFailed.assertionFailure().assertionType()));
    }
    return responseSupport.closeWorkbook(
        executionContext.workbook(),
        ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
            executionContext.protocolVersion(),
            executionContext.journal(),
            executionContext.request().steps().size(),
            calculation,
            List.copyOf(assertions),
            problem,
            stepIndex,
            step.stepId()),
        executionContext.request(),
        executionContext.journal(),
        problem.code(),
        stepIndex,
        step.stepId());
  }

  private PersistenceResult persistWorkbook(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionJournalRecorder journal,
      Path workingDirectory,
      CalculationReport calculation) {
    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    try {
      GridGrindResponsePersistence.PersistenceOutcome persistence =
          workbookSupport.persistWorkbook(
              workbook, request.source(), request.persistence(), workingDirectory);
      persistencePhase.succeed();
      return new PersistenceResult(persistence, null);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(request, workingDirectory)));
      persistencePhase.fail("failed (" + problem.code() + ")");
      return new PersistenceResult(
          null,
          responseSupport.closeWorkbook(
              workbook,
              ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                  protocolVersion,
                  journal,
                  request.steps().size(),
                  calculation,
                  problem,
                  null,
                  null),
              request,
              journal,
              problem.code(),
              null,
              null));
    }
  }

  private record CalculationCheckpoint(
      CalculationReport report, boolean executed, @Nullable GridGrindResponse failureResponse) {}

  private record WorkbookExecutionContext(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionJournalRecorder journal) {}

  private record PersistenceResult(
      GridGrindResponsePersistence.@Nullable PersistenceOutcome persistence,
      @Nullable GridGrindResponse failureResponse) {}

  GridGrindResponse executeDirectEventReadWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      Path workingDirectory) {
    return directEventReadWorkflow.execute(
        protocolVersion, request, warnings, journal, workingDirectory);
  }

  GridGrindResponse executeStreamingWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      Path workingDirectory) {
    return streamingWorkflow.execute(
        protocolVersion, request, executionMode, warnings, journal, workingDirectory);
  }

  private static boolean shouldExecuteCalculationBeforeStep(
      WorkbookPlan request, WorkbookStep step) {
    return CalculationPolicyExecutor.requiresMutationPrefix(request.calculationPolicy())
        && !(step instanceof MutationStep);
  }
}
