package dev.erst.gridgrind.contract.step;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Validates independently bound operation-target pairs through the operation contract registry. */
final class WorkbookStaticTargetValidation {
  private WorkbookStaticTargetValidation() {}

  static List<WorkbookStaticViolation> validate(List<WorkbookStaticStep> steps) {
    List<WorkbookStaticViolation> violations = new ArrayList<>();
    for (WorkbookStaticStep step : steps) {
      step.value()
          .flatMap(WorkbookStaticTargetValidation::targetViolation)
          .ifPresent(
              message ->
                  violations.add(
                      new WorkbookStaticViolation(
                          "steps[" + step.index() + "].target.type", message)));
    }
    return List.copyOf(violations);
  }

  private static Optional<String> targetViolation(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutation ->
          WorkbookOperationContracts.targetViolation(mutation.action(), mutation.target());
      case AssertionStep assertion ->
          WorkbookOperationContracts.targetViolation(assertion.assertion(), assertion.target());
      case InspectionStep inspection ->
          WorkbookOperationContracts.targetViolation(inspection.query(), inspection.target());
    };
  }
}
