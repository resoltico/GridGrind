package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.Objects;
import java.util.Optional;

/** One authored inspection step that receives a stable step id when added to a plan. */
public final class PlannedInspection {
  private final Optional<String> stepId;
  private final Selector target;
  private final InspectionQuery query;

  PlannedInspection(Optional<String> stepId, Selector target, InspectionQuery query) {
    Objects.requireNonNull(stepId, "stepId must not be null");
    if (stepId.isPresent() && stepId.orElseThrow().isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    this.stepId = stepId;
    this.target = Objects.requireNonNull(target, "target must not be null");
    this.query = Objects.requireNonNull(query, "query must not be null");
  }

  PlannedInspection(Selector target, InspectionQuery query) {
    this(Optional.empty(), target, query);
  }

  /** Returns a copy that pins an explicit step id instead of relying on plan auto-generation. */
  public PlannedInspection named(String newStepId) {
    return new PlannedInspection(Optional.of(newStepId), target, query);
  }

  InspectionStep toStep(String generatedStepId) {
    return new InspectionStep(stepId.orElse(generatedStepId), target, query);
  }
}
