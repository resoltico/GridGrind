package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Assertion-specific target-selector rules extracted from workbook-step validation. */
final class AssertionTargetingSupport {
  private AssertionTargetingSupport() {}

  static Class<? extends Selector>[] allowedTargetTypes(
      Assertion assertion,
      Map<Class<? extends Assertion>, Class<? extends Selector>[]> staticTargetTypes) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    Objects.requireNonNull(staticTargetTypes, "staticTargetTypes must not be null");

    if (assertion instanceof Assertion.AnalysisMaxSeverity analysisMaxSeverity) {
      return WorkbookStepValidation.allowedTargetTypes(analysisMaxSeverity.query());
    }
    if (assertion instanceof Assertion.AnalysisFindingPresent analysisFindingPresent) {
      return WorkbookStepValidation.allowedTargetTypes(analysisFindingPresent.query());
    }
    if (assertion instanceof Assertion.AnalysisFindingAbsent analysisFindingAbsent) {
      return WorkbookStepValidation.allowedTargetTypes(analysisFindingAbsent.query());
    }
    if (assertion instanceof Assertion.AllOf allOf) {
      return commonTargetTypes(allOf.assertions(), assertion.assertionType(), staticTargetTypes);
    }
    if (assertion instanceof Assertion.AnyOf anyOf) {
      return commonTargetTypes(anyOf.assertions(), assertion.assertionType(), staticTargetTypes);
    }
    if (assertion instanceof Assertion.Not not) {
      return allowedTargetTypes(not.assertion(), staticTargetTypes);
    }
    return staticAllowedTargetTypesForAssertionType(assertion.getClass(), staticTargetTypes);
  }

  static Class<? extends Selector>[] staticAllowedTargetTypesForAssertionType(
      Class<? extends Assertion> assertionType,
      Map<Class<? extends Assertion>, Class<? extends Selector>[]> staticTargetTypes) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    Objects.requireNonNull(staticTargetTypes, "staticTargetTypes must not be null");

    Optional<String> dynamicRule = dynamicTargetSelectorRuleForAssertionType(assertionType);
    if (dynamicRule.isPresent()) {
      throw new IllegalArgumentException(
          "Assertion type "
              + assertionType.getName()
              + " derives target selectors dynamically: "
              + dynamicRule.orElseThrow());
    }
    Class<? extends Selector>[] configuredTargetTypes = staticTargetTypes.get(assertionType);
    if (configuredTargetTypes == null) {
      throw new IllegalArgumentException(
          "No target-type mapping configured for assertion class " + assertionType.getName());
    }
    return configuredTargetTypes;
  }

  static Optional<String> dynamicTargetSelectorRuleForAssertionType(
      Class<? extends Assertion> assertionType) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    if (assertionType.equals(Assertion.AnalysisMaxSeverity.class)
        || assertionType.equals(Assertion.AnalysisFindingPresent.class)
        || assertionType.equals(Assertion.AnalysisFindingAbsent.class)) {
      return Optional.of("Matches the nested analysis query's target selectors.");
    }
    if (assertionType.equals(Assertion.AllOf.class)
        || assertionType.equals(Assertion.AnyOf.class)) {
      return Optional.of("Matches the intersection of every nested assertion's target selectors.");
    }
    if (assertionType.equals(Assertion.Not.class)) {
      return Optional.of("Matches the nested assertion's target selectors.");
    }
    return Optional.empty();
  }

  static Optional<String> targetSelectorRuleForAssertionType(
      Class<? extends Assertion> assertionType) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    return dynamicTargetSelectorRuleForAssertionType(assertionType);
  }

  static Class<? extends Selector>[] commonTargetTypes(
      Iterable<Assertion> assertions,
      String compositeType,
      Map<Class<? extends Assertion>, Class<? extends Selector>[]> staticTargetTypes) {
    var iterator = assertions.iterator();
    if (!iterator.hasNext()) {
      throw new IllegalArgumentException(
          compositeType + " requires nested assertions with compatible target families");
    }
    Class<? extends Selector>[] intersection =
        allowedTargetTypes(iterator.next(), staticTargetTypes).clone();
    while (iterator.hasNext()) {
      intersection =
          intersect(intersection, allowedTargetTypes(iterator.next(), staticTargetTypes));
    }
    if (intersection.length == 0) {
      throw new IllegalArgumentException(
          compositeType + " requires nested assertions with compatible target families");
    }
    return intersection;
  }

  private static Class<? extends Selector>[] intersect(
      Class<? extends Selector>[] left, Class<? extends Selector>[] right) {
    List<Class<? extends Selector>> intersection = new ArrayList<>();
    for (Class<? extends Selector> leftType : left) {
      for (Class<? extends Selector> rightType : right) {
        if (leftType.equals(rightType)) {
          intersection.add(leftType);
          break;
        }
      }
    }
    @SuppressWarnings("unchecked")
    Class<? extends Selector>[] merged = intersection.toArray(new Class[0]);
    return merged;
  }
}
