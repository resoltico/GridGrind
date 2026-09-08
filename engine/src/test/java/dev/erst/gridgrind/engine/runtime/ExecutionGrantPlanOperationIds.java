package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;

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
