package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Shared target-selector surface metadata for step operations and discovery tooling. */
public final class WorkbookStepTargeting {
  private WorkbookStepTargeting() {}

  /** Returns the static selector families accepted by one mutation action type. */
  public static TargetSurface forMutationActionType(
      Class<? extends MutationAction> mutationActionType) {
    Objects.requireNonNull(mutationActionType, "mutationActionType must not be null");
    return new TargetSurface(
        SelectorJsonSupport.familyInfosFor(
            Arrays.asList(
                WorkbookStepValidation.allowedTargetTypesForMutationActionType(
                    mutationActionType))),
        null);
  }

  /** Returns the static selector families accepted by one inspection query type. */
  public static TargetSurface forInspectionQueryType(
      Class<? extends InspectionQuery> inspectionQueryType) {
    Objects.requireNonNull(inspectionQueryType, "inspectionQueryType must not be null");
    return new TargetSurface(
        SelectorJsonSupport.familyInfosFor(
            Arrays.asList(
                WorkbookStepValidation.allowedTargetTypesForInspectionQueryType(
                    inspectionQueryType))),
        null);
  }

  /** Returns the selector-family surface accepted by one assertion type. */
  public static TargetSurface forAssertionType(Class<? extends Assertion> assertionType) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    String dynamicRule =
        WorkbookStepValidation.dynamicTargetSelectorRuleForAssertionType(assertionType);
    if (dynamicRule != null) {
      return new TargetSurface(List.of(), dynamicRule);
    }
    return new TargetSurface(
        SelectorJsonSupport.familyInfosFor(
            Arrays.asList(
                WorkbookStepValidation.staticAllowedTargetTypesForAssertionType(assertionType))),
        WorkbookStepValidation.targetSelectorRuleForAssertionType(assertionType));
  }

  /** One target-selector surface with selector families plus any derived or disambiguation rule. */
  public record TargetSurface(List<SelectorJsonSupport.FamilyInfo> selectorFamilies, String rule) {
    public TargetSurface {
      Objects.requireNonNull(selectorFamilies, "selectorFamilies must not be null");
      selectorFamilies = List.copyOf(selectorFamilies);
      if (rule != null && rule.isBlank()) {
        throw new IllegalArgumentException("rule must not be blank");
      }
      if (selectorFamilies.isEmpty() && rule == null) {
        throw new IllegalArgumentException(
            "target surface must declare selectorFamilies or a non-blank rule");
      }
    }
  }
}
