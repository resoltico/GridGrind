package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import java.util.Objects;
import java.util.Optional;

/** One authored assertion step that receives a stable step id when added to a plan. */
public final class PlannedAssertion {
  private final Optional<String> stepId;
  private final Selector target;
  private final Assertion assertion;

  PlannedAssertion(Optional<String> stepId, Selector target, Assertion assertion) {
    Objects.requireNonNull(stepId, "stepId must not be null");
    if (stepId.isPresent() && stepId.orElseThrow().isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    this.stepId = stepId;
    this.target = Objects.requireNonNull(target, "target must not be null");
    this.assertion = Objects.requireNonNull(assertion, "assertion must not be null");
  }

  PlannedAssertion(Selector target, Assertion assertion) {
    this(Optional.empty(), target, assertion);
  }

  /** Returns a copy that pins an explicit step id instead of relying on plan auto-generation. */
  public PlannedAssertion named(String newStepId) {
    return new PlannedAssertion(Optional.of(newStepId), target, assertion);
  }

  AssertionStep toStep(String generatedStepId) {
    return new AssertionStep(stepId.orElse(generatedStepId), target, assertion);
  }
}
