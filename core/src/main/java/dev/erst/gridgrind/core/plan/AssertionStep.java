package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered assertion step that verifies workbook state through the canonical contract. */
public record AssertionStep(StepId stepId, Selector target, Assertion assertion)
    implements WorkbookStep {
  public AssertionStep {
    java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    java.util.Objects.requireNonNull(target, "target must not be null");
    java.util.Objects.requireNonNull(assertion, "assertion must not be null");
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(AssertionStep.class);
  }
}
