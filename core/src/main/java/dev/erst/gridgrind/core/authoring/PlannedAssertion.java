package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.plan.Assertion;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;
import java.util.Optional;

/** One authored assertion step that receives a stable step id when added to a plan. */
public final class PlannedAssertion {
  private final Optional<StepId> stepId;
  private final Selector target;
  private final Assertion assertion;

  PlannedAssertion(Optional<StepId> stepId, Selector target, Assertion assertion) {
    Objects.requireNonNull(stepId, "stepId must not be null");
    this.stepId = stepId;
    this.target = Objects.requireNonNull(target, "target must not be null");
    this.assertion = Objects.requireNonNull(assertion, "assertion must not be null");
  }

  PlannedAssertion(Selector target, Assertion assertion) {
    this(Optional.empty(), target, assertion);
  }

  /** Returns a copy that pins an explicit step id instead of relying on plan auto-generation. */
  public PlannedAssertion named(StepId newStepId) {
    return new PlannedAssertion(Optional.of(newStepId), target, assertion);
  }

  AssertionStep toStep(StepId generatedStepId) {
    return new AssertionStep(stepId.orElse(generatedStepId), target, assertion);
  }
}
