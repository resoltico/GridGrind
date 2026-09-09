package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.List;
import java.util.Optional;

/** Runs full-XSSF calculation checkpoints and workbook steps in their declared order. */
final class FullXssfStepLoop {
  private final FullXssfStepExecution stepExecution;
  private final FullXssfCalculationCheckpoint calculationCheckpoint;

  FullXssfStepLoop(
      FullXssfStepExecution stepExecution, FullXssfCalculationCheckpoint calculationCheckpoint) {
    this.stepExecution = stepExecution;
    this.calculationCheckpoint = calculationCheckpoint;
  }

  Optional<WorkbookResult> execute(
      FullXssfWorkflowState state,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput executionMode) {
    List<WorkbookStep> steps = state.executionContext().request().steps();
    for (int stepIndex = 0; stepIndex < steps.size(); stepIndex++) {
      Optional<WorkbookResult> failure =
          executeStep(state, executionMode, stepIndex, steps.get(stepIndex));
      if (failure.isPresent()) {
        return failure;
      }
    }
    return Optional.empty();
  }

  private Optional<WorkbookResult> executeStep(
      FullXssfWorkflowState state,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput executionMode,
      int stepIndex,
      WorkbookStep step) {
    FullXssfCalculationCheckpoint.Result checkpoint =
        calculationCheckpoint.beforeStep(
            state.executionContext(),
            stepIndex,
            step,
            state.calculation(),
            state.calculationExecuted(),
            state.formulaOrigins());
    state.calculation(checkpoint.report());
    state.calculationExecuted(checkpoint.executed());
    if (checkpoint.failureResponse() != null) {
      return Optional.of(checkpoint.failureResponse());
    }
    ExecutionJournalRecorder.StepHandle stepHandle =
        state.executionContext().journal().beginStep(stepIndex, step);
    try {
      Optional<AssertionFailedException> collectedFailure =
          stepExecution.execute(
              state.executionContext().workbook(),
              state.workbookLocation(),
              executionMode,
              state.executionContext().request().assertionMode(),
              state.executionContext().planAssertions(),
              state.executionContext().inspections(),
              step,
              state.formulaOrigins(),
              stepIndex);
      recordStepResult(state, collectedFailure, stepIndex, step, stepHandle);
      return Optional.empty();
    } catch (Exception exception) {
      return Optional.of(
          stepExecution.closeFailure(
              state.executionContext(),
              state.calculation(),
              stepIndex,
              step,
              stepHandle,
              exception,
              state.formulaOrigins()));
    }
  }

  private static void recordStepResult(
      FullXssfWorkflowState state,
      Optional<AssertionFailedException> collectedFailure,
      int stepIndex,
      WorkbookStep step,
      ExecutionJournalRecorder.StepHandle stepHandle) {
    if (collectedFailure.isEmpty()) {
      stepHandle.succeed();
      return;
    }
    AssertionFailedException failure = collectedFailure.orElseThrow();
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            failure,
            ExecutionStepContextFactory.contextFor(
                state.executionContext().request(), stepIndex, step, failure));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().stage(), problem.message());
    state.collectedAssertionFailures().add(stepIndex, step.stepId(), problem);
  }
}
