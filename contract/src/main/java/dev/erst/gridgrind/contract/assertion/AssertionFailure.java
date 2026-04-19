package dev.erst.gridgrind.contract.assertion;

import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.List;
import java.util.Objects;

/** Structured assertion-mismatch payload attached to ASSERTION_FAILED problems. */
public record AssertionFailure(
    String stepId,
    String assertionType,
    Selector target,
    Assertion assertion,
    List<InspectionResult> observations) {
  public AssertionFailure {
    stepId = AssertionSupport.requireNonBlank(stepId, "stepId");
    assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(assertion, "assertion must not be null");
    observations = AssertionSupport.copyObservations(observations, "observations");
  }
}
