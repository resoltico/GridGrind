package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered workbook inspection step with stable identity and a single semantic target. */
public record InspectionStep(StepId stepId, Selector target, InspectionQuery query)
    implements WorkbookStep {
  public InspectionStep {
    java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    java.util.Objects.requireNonNull(target, "target must not be null");
    java.util.Objects.requireNonNull(query, "query must not be null");
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(InspectionStep.class);
  }
}
