package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Runs the required pre-observation calculation checkpoint for a full-XSSF workflow. */
final class FullXssfCalculationCheckpoint {
  private final ExecutionCalculationSupport calculationSupport;
  private final ExecutionResponseSupport responseSupport;

  FullXssfCalculationCheckpoint(
      ExecutionCalculationSupport calculationSupport, ExecutionResponseSupport responseSupport) {
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
  }

  Result beforeStep(
      WorkbookWorkflowExecutionContext executionContext,
      int stepIndex,
      WorkbookStep step,
      CalculationReport currentCalculation,
      boolean calculationExecuted,
      FormulaOriginTracker formulaOrigins) {
    if (calculationExecuted
        || !CalculationPolicyExecutor.requiresMutationPrefix(
            executionContext.request().calculationPolicy())
        || step instanceof MutationStep) {
      return new Result(currentCalculation, calculationExecuted, null);
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
      return new Result(outcome.report(), true, null);
    }
    GridGrindProblemDetail.Problem problem = outcome.failure().orElseThrow();
    return new Result(
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

  record Result(
      CalculationReport report, boolean executed, @Nullable WorkbookResult failureResponse) {}
}
