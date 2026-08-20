package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import java.util.ArrayList;
import java.util.List;

/** Enforces the terminal assertion phase required by collected assertion execution. */
final class WorkbookStaticAssertionValidation {
  private WorkbookStaticAssertionValidation() {}

  static List<WorkbookStaticViolation> validate(WorkbookStaticRequest request) {
    if (request.execution().isEmpty()
        || request.execution().orElseThrow().effectiveAssertionMode()
            != AssertionModeInput.COLLECT) {
      return List.of();
    }
    List<WorkbookStaticViolation> violations = new ArrayList<>();
    boolean seenAssertion = false;
    for (WorkbookStaticStep step : request.steps()) {
      if (step.value().isEmpty()) {
        continue;
      }
      WorkbookStep value = step.value().orElseThrow();
      if (value instanceof AssertionStep) {
        seenAssertion = true;
      } else if (seenAssertion && value instanceof MutationStep) {
        violations.add(
            new WorkbookStaticViolation(
                "steps[" + step.index() + "]",
                "execution.assertionMode=COLLECT requires every MUTATION step to appear"
                    + " before the first ASSERTION step"));
      }
    }
    return List.copyOf(violations);
  }
}
