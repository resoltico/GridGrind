package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import java.util.Objects;

/** One authored assertion step that receives a stable step id when added to a plan. */
public final class PlannedAssertion {
  private final String stepId;
  private final Selector target;
  private final Assertion assertion;

  PlannedAssertion(String stepId, Selector target, Assertion assertion) {
    if (stepId != null && stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    this.stepId = stepId;
    this.target = Objects.requireNonNull(target, "target must not be null");
    this.assertion = Objects.requireNonNull(assertion, "assertion must not be null");
  }

  PlannedAssertion(Selector target, Assertion assertion) {
    this(null, target, assertion);
  }

  /** Returns a copy that pins an explicit step id instead of relying on plan auto-generation. */
  public PlannedAssertion named(String newStepId) {
    return new PlannedAssertion(newStepId, target, assertion);
  }

  AssertionStep toStep(String generatedStepId) {
    return new AssertionStep(stepId == null ? generatedStepId : stepId, target, assertion);
  }
}
