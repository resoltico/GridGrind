package dev.erst.gridgrind.contract.assertion;

/** Ordered success-side acknowledgment for one passed assertion step. */
public record AssertionResult(String stepId, String assertionType) {
  public AssertionResult {
    stepId = AssertionSupport.requireNonBlank(stepId, "stepId");
    assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
  }
}
