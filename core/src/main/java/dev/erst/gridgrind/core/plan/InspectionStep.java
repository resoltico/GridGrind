package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;

/** One ordered workbook inspection step with stable identity and a single semantic target. */
public record InspectionStep(String stepId, Selector target, InspectionQuery query)
    implements WorkbookStep {
  public InspectionStep {
    stepId = WorkbookStepValidation.requireStepId(stepId);
    target = WorkbookStepValidation.requireTarget(target);
    query = WorkbookStepValidation.requireQuery(query);
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(InspectionStep.class);
  }
}
