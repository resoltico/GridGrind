package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
import java.util.Objects;

/** Immutable result of evaluating one assertion against one workbook observation set. */
record AssertionEvaluation(boolean passed, List<InspectionResult> observations, String message) {
  AssertionEvaluation {
    observations = List.copyOf(observations);
    Objects.requireNonNull(message, "message must not be null");
  }

  static AssertionEvaluation pass(List<InspectionResult> observations) {
    return new AssertionEvaluation(true, observations, "passed");
  }

  static AssertionEvaluation fail(List<InspectionResult> observations, String message) {
    return new AssertionEvaluation(false, observations, message);
  }
}
