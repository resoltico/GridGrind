package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered workbook inspection step with stable identity and a single semantic target. */
@dev.erst.gridgrind.core.metadata.Variant(
    id = "INSPECTION",
    summary =
        "Run one factual or analytical inspection query against the selected workbook target.")
public record InspectionStep(
    @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
        @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.STRING_PATTERN,
            text = dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues.STEP_ID_PATTERN)
        StepId stepId,
    Selector target,
    InspectionQuery query)
    implements WorkbookStep, PreparedStep {
  public InspectionStep {
    java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    java.util.Objects.requireNonNull(target, "target must not be null");
    java.util.Objects.requireNonNull(query, "query must not be null");
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(InspectionStep.class);
  }

  @Override
  public WorkbookStep authoredStep() {
    return this;
  }
}
