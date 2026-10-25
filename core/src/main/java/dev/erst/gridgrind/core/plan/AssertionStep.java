package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;

/** One ordered assertion step that verifies workbook state through the canonical contract. */
public record AssertionStep(String stepId, Selector target, Assertion assertion)
    implements WorkbookStep {
  public AssertionStep {
    stepId = WorkbookStepValidation.requireStepId(stepId);
    target = WorkbookStepValidation.requireTarget(target);
    assertion = WorkbookStepValidation.requireAssertion(assertion);
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(AssertionStep.class);
  }
}
