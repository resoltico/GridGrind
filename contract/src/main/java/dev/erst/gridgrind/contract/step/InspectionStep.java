package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;

/** One ordered workbook inspection step with stable identity and a single semantic target. */
public record InspectionStep(String stepId, Selector target, InspectionQuery query)
    implements WorkbookStep {
  public InspectionStep {
    stepId = WorkbookStepValidation.requireStepId(stepId);
    target = WorkbookStepValidation.requireTarget(target);
    query = WorkbookStepValidation.requireCompatible(target, query);
  }

  @Override
  public String stepKind() {
    return GridGrindProtocolTypeNames.workbookStepTypeName(InspectionStep.class);
  }
}
