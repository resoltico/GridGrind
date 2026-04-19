package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.Objects;

/** One authored inspection step that receives a stable step id when added to a plan. */
public record PlannedInspection(String stepId, Selector target, InspectionQuery query) {
  public PlannedInspection {
    if (stepId != null && stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(query, "query must not be null");
  }

  /** Creates an unnamed authored inspection that will receive an auto-generated step id. */
  public PlannedInspection(Selector target, InspectionQuery query) {
    this(null, target, query);
  }

  /** Returns a copy that pins an explicit step id instead of relying on plan auto-generation. */
  public PlannedInspection named(String newStepId) {
    return new PlannedInspection(newStepId, target, query);
  }

  InspectionStep toStep(String generatedStepId) {
    return new InspectionStep(stepId == null ? generatedStepId : stepId, target, query);
  }
}
