package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.Objects;

/** One authored mutation step that receives a stable step id when added to a plan. */
public final class PlannedMutation {
  private final String stepId;
  private final Selector target;
  private final MutationAction action;

  PlannedMutation(String stepId, Selector target, MutationAction action) {
    if (stepId != null && stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    this.stepId = stepId;
    this.target = Objects.requireNonNull(target, "target must not be null");
    this.action = Objects.requireNonNull(action, "action must not be null");
  }

  PlannedMutation(Selector target, MutationAction action) {
    this(null, target, action);
  }

  /** Returns a copy that pins an explicit step id instead of relying on plan auto-generation. */
  public PlannedMutation named(String newStepId) {
    return new PlannedMutation(newStepId, target, action);
  }

  MutationStep toStep(String generatedStepId) {
    return new MutationStep(stepId == null ? generatedStepId : stepId, target, action);
  }
}
