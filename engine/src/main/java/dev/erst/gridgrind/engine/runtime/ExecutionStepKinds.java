package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import java.util.Objects;

/** Step-kind token extraction shared by journals and problem contexts. */
final class ExecutionStepKinds {
  private ExecutionStepKinds() {}

  static String stepType(WorkbookStep step) {
    Objects.requireNonNull(step, "step must not be null");
    return switch (step) {
      case MutationStep mutationStep -> mutationStep.action().actionType();
      case AssertionStep assertionStep -> assertionStep.assertion().assertionType();
      case InspectionStep inspectionStep -> inspectionStep.query().queryType();
    };
  }
}
