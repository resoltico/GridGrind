package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookStep;

/** Derives the canonical operation identifier requested by each authored workbook step. */
final class ExecutionGrantPlanOperationIds {
  private ExecutionGrantPlanOperationIds() {}

  static String forStep(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutation -> mutation.action().actionType();
      case AssertionStep assertion -> assertion.assertion().assertionType();
      case InspectionStep inspection -> inspection.query().queryType();
    };
  }
}
