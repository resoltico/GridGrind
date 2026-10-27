package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered assertion step that verifies workbook state through the canonical contract. */
@dev.erst.gridgrind.core.metadata.Variant(
    id = "ASSERTION",
    summary = "Verify one authored expectation against the selected workbook target.")
public record AssertionStep(
    @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
        @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.STRING_PATTERN,
            text = dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues.STEP_ID_PATTERN)
        StepId stepId,
    Selector target,
    Assertion assertion)
    implements WorkbookStep, PreparedStep {
  public AssertionStep {
    java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    java.util.Objects.requireNonNull(target, "target must not be null");
    java.util.Objects.requireNonNull(assertion, "assertion must not be null");
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(AssertionStep.class);
  }

  @Override
  public WorkbookStep authoredStep() {
    return this;
  }
}
