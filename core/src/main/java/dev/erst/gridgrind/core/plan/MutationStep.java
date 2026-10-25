package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered workbook mutation step with stable identity and a single semantic target. */
public record MutationStep(StepId stepId, Selector target, MutationAction action)
    implements WorkbookStep {
  public MutationStep {
    java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    java.util.Objects.requireNonNull(target, "target must not be null");
    java.util.Objects.requireNonNull(action, "action must not be null");
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(MutationStep.class);
  }
}
