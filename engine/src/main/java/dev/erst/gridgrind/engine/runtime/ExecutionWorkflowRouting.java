package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.InspectionStep;

/** Selects the executable workflow after the static request contract has accepted a plan. */
final class ExecutionWorkflowRouting {
  private ExecutionWorkflowRouting() {}

  static boolean directEventReadEligible(WorkbookPlan request, ExecutionModeInput executionMode) {
    return executionMode instanceof ExecutionModeInput.EventRead
        && request.calculationPolicy().allowsEventRead()
        && request.steps().stream().allMatch(InspectionStep.class::isInstance)
        && request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None
        && request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile;
  }

  static ExecutionModeInput executionMode(WorkbookPlan request) {
    return request.effectiveExecutionMode();
  }
}
