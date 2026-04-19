package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.Selector;

/** One ordered workbook mutation step with stable identity and a single semantic target. */
public record MutationStep(String stepId, Selector target, MutationAction action)
    implements WorkbookStep {
  public MutationStep {
    stepId = WorkbookStepValidation.requireStepId(stepId);
    target = WorkbookStepValidation.requireTarget(target);
    action = WorkbookStepValidation.requireCompatible(target, action);
  }

  @Override
  public String stepKind() {
    return "MUTATION";
  }
}
