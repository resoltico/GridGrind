package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Ordered assertion failures retained while a terminal COLLECT phase continues. */
final class CollectedAssertionFailures {
  private final List<Failure> failures = new ArrayList<>();

  void add(int stepIndex, String stepId, GridGrindProblemDetail.Problem problem) {
    failures.add(new Failure(stepIndex, stepId, problem));
  }

  boolean isEmpty() {
    return failures.isEmpty();
  }

  Failure first() {
    return failures.getFirst();
  }

  /** One collected assertion failure in authored step order. */
  record Failure(int stepIndex, String stepId, GridGrindProblemDetail.Problem problem) {
    Failure {
      Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(problem, "problem must not be null");
    }
  }
}
