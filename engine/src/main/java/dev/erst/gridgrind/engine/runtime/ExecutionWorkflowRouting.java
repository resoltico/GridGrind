package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;

/** Selects the executable workflow after the static request contract has accepted a plan. */
final class ExecutionWorkflowRouting {
  private ExecutionWorkflowRouting() {}

  static boolean directEventReadEligible(WorkbookPlan request, ExecutionModeSpec executionMode) {
    return executionMode instanceof ExecutionModeSpec.EventRead
        && request.calculationPolicy().allowsEventRead()
        && request.steps().stream().allMatch(InspectionStep.class::isInstance)
        && request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None
        && request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile;
  }

  static ExecutionModeSpec executionMode(WorkbookPlan request) {
    return request.effectiveExecutionMode();
  }
}
