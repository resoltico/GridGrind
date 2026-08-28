package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.io.IOException;
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

  WorkbookResult executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), bindings.workingDirectory());
    List<AssertionResult> assertions = new ArrayList<>();
    CollectedAssertionFailures collectedAssertionFailures = new CollectedAssertionFailures();
    List<InspectionResult> inspections = new ArrayList<>();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    boolean calculationExecuted = false;
    FormulaOriginTracker formulaOrigins = new FormulaOriginTracker();
    WorkbookWorkflowExecutionContext executionContext =
        new WorkbookWorkflowExecutionContext(
            protocolVersion, request, workbook, journal, warnings, assertions, inspections);

    for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
      WorkbookStep step = request.steps().get(stepIndex);
      CalculationCheckpoint calculationCheckpoint =
          executeCalculationBeforeStepIfNeeded(
              executionContext, stepIndex, step, calculation, calculationExecuted, formulaOrigins);
      calculation = calculationCheckpoint.report();
      calculationExecuted = calculationCheckpoint.executed();
      if (calculationCheckpoint.failureResponse() != null) {
        return calculationCheckpoint.failureResponse();
      }

      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
      try {
        java.util.Optional<AssertionFailedException> collectedFailure =
            executeWorkbookStep(
                workbook,
                workbookLocation,
                executionMode,
                request.assertionMode(),
                assertions,
                inspections,
                step,
                formulaOrigins,
                stepIndex);
        if (collectedFailure.isPresent()) {
          AssertionFailedException assertionFailure = collectedFailure.orElseThrow();
          GridGrindProblemDetail.Problem problem =
              ExecutionResponseSupport.problemFor(
                  assertionFailure,
                  ExecutionStepContextFactory.contextFor(
                      request, stepIndex, step, assertionFailure));
          stepHandle.fail(
              problem.code(), problem.category(), problem.context().stage(), problem.message());
          collectedAssertionFailures.add(stepIndex, step.stepId(), problem);
        } else {
          stepHandle.succeed();
        }
      } catch (Exception exception) {
        return closeFailedStepExecution(
            executionContext, calculation, stepIndex, step, stepHandle, exception, formulaOrigins);
      }
    }

    if (!calculationExecuted) {
      ExecutionCalculationSupport.CalculationExecutionOutcome calculationOutcome =
          calculationSupport.executeCalculationPolicy(
              workbook, request, journal, formulaOrigins, null);
      calculation = calculationOutcome.report();
      warnings.addAll(calculationOutcome.warnings());
      if (calculationOutcome.failure().isPresent()) {
        GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
        return responseSupport.closeWorkbook(
            workbook,
            ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                executionContext.failure(calculation, problem)),
            request,
            journal,
            problem.code(),
            null,
            null);
      }
    }

    if (!collectedAssertionFailures.isEmpty()) {
      CollectedAssertionFailures.Failure firstFailure = collectedAssertionFailures.first();
      return responseSupport.closeWorkbook(
          workbook,
          ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
              executionContext.failure(
                  calculation,
                  firstFailure.problem(),
                  firstFailure.stepIndex(),
                  firstFailure.stepId())),
          request,
          journal,
          firstFailure.problem().code(),
          firstFailure.stepIndex(),
          firstFailure.stepId());
    }

    PersistenceResult persistenceResult = persistWorkbook(executionContext, bindings, calculation);
    if (persistenceResult.failureResponse() != null) {
      return persistenceResult.failureResponse();
    }

    return responseSupport.closeWorkbook(
        workbook,
        new WorkbookResult.Success(
            protocolVersion,
            request.planId(),
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
      WorkbookWorkflowExecutionContext executionContext,
      int stepIndex,
      WorkbookStep step,
      CalculationReport currentCalculation,
      boolean calculationExecuted,
      FormulaOriginTracker formulaOrigins) {
    if (calculationExecuted
        || !shouldExecuteCalculationBeforeStep(executionContext.request(), step)) {
      return new CalculationCheckpoint(currentCalculation, calculationExecuted, null);
    }
    ExecutionCalculationSupport.CalculationExecutionOutcome outcome =
        calculationSupport.executeCalculationPolicy(
            executionContext.workbook(),
            executionContext.request(),
            executionContext.journal(),
            formulaOrigins,
            ExecutionStepContextFactory.referenceFor(stepIndex, step));
    if (outcome.failure().isEmpty()) {
      executionContext.warnings().addAll(outcome.warnings());
      return new CalculationCheckpoint(outcome.report(), true, null);
    }
    GridGrindProblemDetail.Problem problem = outcome.failure().orElseThrow();
    return new CalculationCheckpoint(
        outcome.report(),
        true,
        responseSupport.closeWorkbook(
            executionContext.workbook(),
            ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                executionContext.failure(outcome.report(), problem)),
            executionContext.request(),
            executionContext.journal(),
            problem.code(),
            null,
            null));
  }

  private java.util.Optional<AssertionFailedException> executeWorkbookStep(
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode,
      AssertionModeInput assertionMode,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections,
      WorkbookStep step,
      FormulaOriginTracker formulaOrigins,
      int stepIndex)
      throws IOException, AssertionFailedException {
    return switch (step) {
      case MutationStep mutationStep -> {
        stepSupport
            .mutationStepExecutor()
            .execute(
                workbook,
                mutationStep,
                formulaOrigins,
                ExecutionStepContextFactory.referenceFor(stepIndex, mutationStep));
        yield java.util.Optional.empty();
      }
      case AssertionStep assertionStep -> {
        if (assertionMode == AssertionModeInput.FAIL_FAST) {
          assertions.add(
              stepSupport.executeAssertionStep(
                  assertionStep, workbook, workbookLocation, executionMode));
          yield java.util.Optional.empty();
        }
        AssertionStepExecution assertionExecution =
            stepSupport.executeAssertionStepCollecting(
                assertionStep, workbook, workbookLocation, executionMode);
        assertions.add(assertionExecution.result());
        yield switch (assertionExecution) {
          case AssertionStepExecution.Passed _ -> java.util.Optional.empty();
          case AssertionStepExecution.Failed failed -> java.util.Optional.of(failed.failure());
        };
      }
      case InspectionStep inspectionStep -> {
        inspections.add(
            stepSupport.executeInspectionStep(
                inspectionStep, workbook, workbookLocation, executionMode));
        yield java.util.Optional.empty();
      }
    };
  }

  private WorkbookResult closeFailedStepExecution(
      WorkbookWorkflowExecutionContext executionContext,
      CalculationReport calculation,
      int stepIndex,
      WorkbookStep step,
      ExecutionJournalRecorder.StepHandle stepHandle,
      Exception exception,
      FormulaOriginTracker formulaOrigins) {
    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep triggeringContext =
        ExecutionStepContextFactory.contextFor(
            executionContext.request(), stepIndex, step, exception);
    java.util.Optional<dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference>
        formulaAuthor = formulaOrigins.originFor(exception);
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            exception,
            formulaAuthor
                .map(
                    author ->
                        new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                            ExecutionRequestPaths.requestShape(executionContext.request()),
                            author,
                            triggeringContext.location(),
                            java.util.Optional.of(triggeringContext.step())))
                .orElse(triggeringContext));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().stage(), problem.message());
    if (exception instanceof AssertionFailedException assertionFailed) {
      executionContext
          .assertions()
          .add(
              new AssertionResult.Failed(
                  assertionFailed.assertionFailure().stepId(),
                  assertionFailed.assertionFailure().assertionType(),
                  assertionFailed.assertionFailure()));
    }
    return responseSupport.closeWorkbook(
        executionContext.workbook(),
        ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
            executionContext.failure(calculation, problem, stepIndex, step.stepId())),
        executionContext.request(),
        executionContext.journal(),
        problem.code(),
        stepIndex,
        step.stepId());
  }

  private PersistenceResult persistWorkbook(
      WorkbookWorkflowExecutionContext executionContext,
      ExecutionInputBindings bindings,
      CalculationReport calculation) {
    ExecutionJournalRecorder.PhaseHandle persistencePhase =
        executionContext.journal().beginPersistence();
    try {
      WorkbookResultPersistence.PersistenceOutcome persistence =
          workbookSupport.persistWorkbook(
              executionContext.workbook(),
              executionContext.request().source(),
              executionContext.request().persistence(),
              bindings);
      persistencePhase.succeed();
      return new PersistenceResult(persistence, null);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(executionContext.request()),
                  ExecutionRequestPaths.persistenceReference(
                      executionContext.request(), bindings.workingDirectory())));
      persistencePhase.fail(problem.code());
      return new PersistenceResult(
          null,
          responseSupport.closeWorkbook(
              executionContext.workbook(),
              ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                  executionContext.failure(calculation, problem)),
              executionContext.request(),
              executionContext.journal(),
              problem.code(),
              null,
              null));
    }
  }

  private record CalculationCheckpoint(
      CalculationReport report, boolean executed, @Nullable WorkbookResult failureResponse) {}

  private record PersistenceResult(
      WorkbookResultPersistence.@Nullable PersistenceOutcome persistence,
      @Nullable WorkbookResult failureResponse) {}

  WorkbookResult executeDirectEventReadWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    return directEventReadWorkflow.execute(protocolVersion, request, warnings, journal, bindings);
  }

  WorkbookResult executeStreamingWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    return streamingWorkflow.execute(
        protocolVersion, request, executionMode, warnings, journal, bindings);
  }

  private static boolean shouldExecuteCalculationBeforeStep(
      WorkbookPlan request, WorkbookStep step) {
    return CalculationPolicyExecutor.requiresMutationPrefix(request.calculationPolicy())
        && !(step instanceof MutationStep);
  }
}
