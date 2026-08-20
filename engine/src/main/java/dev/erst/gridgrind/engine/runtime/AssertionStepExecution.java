package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import java.util.Objects;

/** One assertion outcome, retaining failure detail when terminal collection must continue. */
sealed interface AssertionStepExecution
    permits AssertionStepExecution.Passed, AssertionStepExecution.Failed {
  /** Returns the assertion outcome that belongs in the response assertion matrix. */
  AssertionResult result();

  /** A passing assertion outcome. */
  record Passed(AssertionResult result) implements AssertionStepExecution {
    public Passed {
      Objects.requireNonNull(result, "result must not be null");
    }
  }

  /** A failed assertion outcome whose canonical detail is retained for the final response. */
  record Failed(AssertionResult result, AssertionFailedException failure)
      implements AssertionStepExecution {
    public Failed {
      Objects.requireNonNull(result, "result must not be null");
      Objects.requireNonNull(failure, "failure must not be null");
    }
  }
}
