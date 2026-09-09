package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.io.IOException;
import java.util.List;
import java.util.Objects;

/** Executes one full-XSSF step and shapes its truthful terminal failure response. */
final class FullXssfStepExecution {
  private final ExecutionStepSupport stepSupport;
  private final ExecutionResponseSupport responseSupport;

  FullXssfStepExecution(
      ExecutionStepSupport stepSupport, ExecutionResponseSupport responseSupport) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
  }

  java.util.Optional<AssertionFailedException> execute(
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

  dev.erst.gridgrind.contract.dto.WorkbookResult closeFailure(
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
          .planAssertions()
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
}
