package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;

/** One ordered workbook mutation step with stable identity and a single semantic target. */
public record MutationStep(String stepId, Selector target, MutationAction action)
    implements WorkbookStep {
  public MutationStep {
    stepId = WorkbookStepValidation.requireStepId(stepId);
    target = WorkbookStepValidation.requireTarget(target);
    action = WorkbookStepValidation.requireAction(action);
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(MutationStep.class);
  }
}
