package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.Objects;

/** One authored inspection step that receives a stable step id when added to a plan. */
public final class PlannedInspection {
  private final String stepId;
  private final Selector target;
  private final InspectionQuery query;

  PlannedInspection(String stepId, Selector target, InspectionQuery query) {
    if (stepId != null && stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    this.stepId = stepId;
    this.target = Objects.requireNonNull(target, "target must not be null");
    this.query = Objects.requireNonNull(query, "query must not be null");
  }

  PlannedInspection(Selector target, InspectionQuery query) {
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
