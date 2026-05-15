package dev.erst.gridgrind.contract.assertion;

/** Ordered outcome record for one assertion step. */
public record AssertionResult(AssertionOutcome outcome, String stepId, String assertionType) {
  public AssertionResult {
    java.util.Objects.requireNonNull(outcome, "outcome must not be null");
    stepId = AssertionSupport.requireNonBlank(stepId, "stepId");
    assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
  }
}
