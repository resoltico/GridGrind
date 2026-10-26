package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionModeSpec;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
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
      ExecutionModeSpec executionMode,
      AssertionModeSpec assertionMode,
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
        if (assertionMode == AssertionModeSpec.FAIL_FAST) {
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

  dev.erst.gridgrind.core.result.WorkbookResult closeFailure(
      WorkbookWorkflowExecutionContext executionContext,
      CalculationReport calculation,
      int stepIndex,
      WorkbookStep step,
      ExecutionJournalRecorder.StepHandle stepHandle,
      Exception exception,
      FormulaOriginTracker formulaOrigins) {
    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep triggeringContext =
        ExecutionStepContextFactory.contextFor(
            executionContext.request(), stepIndex, step, exception);
    java.util.Optional<dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference>
        formulaAuthor = formulaOrigins.originFor(exception);
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            exception,
            formulaAuthor
                .map(
                    author ->
                        new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                            ExecutionRequestPaths.requestShape(executionContext.request()),
                            author,
                            triggeringContext.location(),
                            java.util.Optional.of(triggeringContext.step())))
                .orElse(triggeringContext));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().type(), problem.message());
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
