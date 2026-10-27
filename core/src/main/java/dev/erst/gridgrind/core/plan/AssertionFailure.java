package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** Structured assertion-mismatch payload attached to ASSERTION_FAILED problems. */
public record AssertionFailure(
    StepId stepId,
    String assertionType,
    Selector target,
    Assertion assertion,
    List<InspectionResult> observations) {
  public AssertionFailure {

    java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    assertionType = TextValidation.requireNonBlank(assertionType, "assertionType");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(assertion, "assertion must not be null");
    observations = AssertionSupport.copyObservations(observations, "observations");
  }
}
