package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/** Shared validation helpers for workbook-step compact constructors. */
final class WorkbookStepValidation {
  private static final Pattern STEP_ID_PATTERN = Pattern.compile("[A-Za-z0-9._-]+");

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

  static MutationAction requireCompatible(Selector target, MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    SelectorTargetingSupport.requireTargetType(
        target, action.actionType(), allowedTargetTypes(action));
    return action;
  }

  static InspectionQuery requireCompatible(Selector target, InspectionQuery query) {
    Objects.requireNonNull(query, "query must not be null");
    SelectorTargetingSupport.requireTargetType(
        target, query.queryType(), allowedTargetTypes(query));
    return query;
  }

  static Assertion requireCompatible(Selector target, Assertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    SelectorTargetingSupport.requireTargetType(
        target, assertion.assertionType(), allowedTargetTypes(assertion));
    return assertion;
  }

  static Class<? extends Selector>[] allowedTargetTypes(MutationAction action) {
    return MutationAction.allowedTargetTypes(action);
  }

  static Class<? extends Selector>[] allowedTargetTypes(InspectionQuery query) {
    return InspectionQuery.allowedTargetTypes(query);
  }

  static Class<? extends Selector>[] allowedTargetTypes(Assertion assertion) {
    return Assertion.allowedTargetTypes(assertion);
  }

  static Class<? extends Selector>[] allowedTargetTypesForMutationActionType(
      Class<? extends MutationAction> actionType) {
    return MutationAction.allowedTargetTypesForType(actionType);
  }

  static Class<? extends Selector>[] allowedTargetTypesForInspectionQueryType(
      Class<? extends InspectionQuery> queryType) {
    return InspectionQuery.allowedTargetTypesForType(queryType);
  }

  static Class<? extends Selector>[] staticAllowedTargetTypesForAssertionType(
      Class<? extends Assertion> assertionType) {
    return Assertion.staticAllowedTargetTypesForType(assertionType);
  }

  static Optional<String> dynamicTargetSelectorRuleForAssertionType(
      Class<? extends Assertion> assertionType) {
    return Assertion.dynamicTargetSelectorRuleForType(assertionType);
  }

  static Optional<String> targetSelectorRuleForAssertionType(
      Class<? extends Assertion> assertionType) {
    return Assertion.targetSelectorRuleForType(assertionType);
  }
}
