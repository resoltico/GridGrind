package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.selector.Selector;

/** One ordered assertion step that verifies workbook state through the canonical contract. */
public record AssertionStep(String stepId, Selector target, Assertion assertion)
    implements WorkbookStep {
  public AssertionStep {
    stepId = WorkbookStepValidation.requireStepId(stepId);
    target = WorkbookStepValidation.requireTarget(target);
    assertion = WorkbookStepValidation.requireCompatible(target, assertion);
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(AssertionStep.class);
  }
}
