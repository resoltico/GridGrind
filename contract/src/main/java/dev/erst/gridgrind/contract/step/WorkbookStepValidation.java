package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.ProtocolConstraintValues;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.Objects;
import java.util.regex.Pattern;

/** Shared validation helpers for workbook-step compact constructors. */
final class WorkbookStepValidation {
  private static final Pattern STEP_ID_PATTERN =
      Pattern.compile(ProtocolConstraintValues.STEP_ID_PATTERN);

  private WorkbookStepValidation() {}

  static String requireStepId(String stepId) { // LIM-006A
    Objects.requireNonNull(stepId, "stepId must not be null");
    if (stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    if (!STEP_ID_PATTERN.matcher(stepId).matches()) {
      throw new IllegalArgumentException("stepId must match [A-Za-z0-9._-]+");
    }
    return stepId;
  }

  static Selector requireTarget(Selector target) {
    Objects.requireNonNull(target, "target must not be null");
    return target;
  }

  static MutationAction requireAction(MutationAction action) {
    return Objects.requireNonNull(action, "action must not be null");
  }

  static InspectionQuery requireQuery(InspectionQuery query) {
    return Objects.requireNonNull(query, "query must not be null");
  }

  static Assertion requireAssertion(Assertion assertion) {
    return Objects.requireNonNull(assertion, "assertion must not be null");
  }
}
