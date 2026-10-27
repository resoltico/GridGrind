package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.result.WorkbookResult;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Runs the required pre-observation calculation checkpoint for a full-XSSF workflow. */
final class FullXssfCalculationCheckpoint {
  private final CalculationWorkflow calculationSupport;
  private final ExecutionResults responseSupport;

  FullXssfCalculationCheckpoint(
      CalculationWorkflow calculationSupport, ExecutionResults responseSupport) {
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
    CalculationWorkflow.CalculationExecutionOutcome outcome =
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
            ExecutionResults.failureResponseWithoutPlanOutcomeEvent(
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
