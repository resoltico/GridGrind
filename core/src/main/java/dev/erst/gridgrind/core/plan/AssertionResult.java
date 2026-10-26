package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;

/** Ordered assertion outcome carrying complete evidence for every failed assertion step. */
@Union(discriminator = "status")
public sealed interface AssertionResult permits AssertionResult.Passed, AssertionResult.Failed {
  /** Returns the stable outcome discriminator. */
  AssertionStatus status();

  /** Returns the authored assertion step id. */
  StepId stepId();

  /** Returns the authored assertion discriminator. */
  String assertionType();

  /** A passed assertion without failure evidence. */
  @Variant(id = "PASSED", summary = "A passed assertion without failure evidence.")
  record Passed(StepId stepId, String assertionType) implements AssertionResult {
    public Passed {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
    }

    @Override
    public AssertionStatus status() {
      return AssertionStatus.PASSED;
    }
  }

  /** A failed assertion with the complete target, authored assertion, and observed evidence. */
  @Variant(
      id = "FAILED",
      summary =
          "A failed assertion with the complete target, authored assertion, and observed evidence.")
  record Failed(StepId stepId, String assertionType, AssertionFailure failure)
      implements AssertionResult {
    public Failed {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      assertionType = AssertionSupport.requireNonBlank(assertionType, "assertionType");
      Objects.requireNonNull(failure, "failure must not be null");
      if (!stepId.equals(failure.stepId()) || !assertionType.equals(failure.assertionType())) {
        throw new IllegalArgumentException(
            "failed assertion result identity must match failure evidence");
      }
    }

    @Override
    public AssertionStatus status() {
      return AssertionStatus.FAILED;
    }
  }
}
