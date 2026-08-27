package dev.erst.gridgrind.contract.assertion;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Ordered assertion outcome carrying complete evidence for every failed assertion step. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "outcome")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AssertionResult.Passed.class, name = "PASSED"),
  @JsonSubTypes.Type(value = AssertionResult.Failed.class, name = "FAILED")
})
public sealed interface AssertionResult permits AssertionResult.Passed, AssertionResult.Failed {
  /** Returns the stable outcome discriminator. */
  AssertionOutcome outcome();

  /** Returns the authored assertion step id. */
  String stepId();

  /** Returns the authored assertion discriminator. */
  String assertionType();

  /** A passed assertion without failure evidence. */
  record Passed(String stepId, String assertionType) implements AssertionResult {
    public Passed {
      stepId = AssertionSupport.requireNonBlank(stepId, "stepId");
      assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
    }

    @Override
    public AssertionOutcome outcome() {
      return AssertionOutcome.PASSED;
    }
  }

  /** A failed assertion with the complete target, authored assertion, and observed evidence. */
  record Failed(String stepId, String assertionType, AssertionFailure failure)
      implements AssertionResult {
    public Failed {
      stepId = AssertionSupport.requireNonBlank(stepId, "stepId");
      assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
      Objects.requireNonNull(failure, "failure must not be null");
      if (!stepId.equals(failure.stepId()) || !assertionType.equals(failure.assertionType())) {
        throw new IllegalArgumentException(
            "failed assertion result identity must match failure evidence");
      }
    }

    @Override
    public AssertionOutcome outcome() {
      return AssertionOutcome.FAILED;
    }
  }
}
