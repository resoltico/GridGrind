package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import java.util.Objects;

/** One authored assertion step that receives a stable step id when added to a plan. */
public record PlannedAssertion(String stepId, Selector target, Assertion assertion) {
  public PlannedAssertion {
    if (stepId != null && stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(assertion, "assertion must not be null");
  }

  /** Creates an unnamed authored assertion that will receive an auto-generated step id. */
  public PlannedAssertion(Selector target, Assertion assertion) {
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
