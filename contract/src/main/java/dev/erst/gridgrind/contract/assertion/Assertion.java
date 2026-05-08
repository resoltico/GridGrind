package dev.erst.gridgrind.contract.assertion;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.catalog.ProtocolTargetingMode;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.Objects;
import java.util.Optional;

/** First-class verification contract evaluated against a workbook target. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
public sealed interface Assertion
    permits PresenceAssertion,
        CellAssertion,
        WorkbookFactAssertion,
        AnalysisAssertion,
        CompositeAssertion {

  /** Stable SCREAMING_SNAKE_CASE discriminator mirrored in catalog and result surfaces. */
  default String assertionType() {
    return GridGrindProtocolTypeNames.assertionTypeName(getClass().asSubclass(Assertion.class));
  }

  /** Returns the selector types accepted by one assertion instance. */
  static Class<? extends Selector>[] allowedTargetTypes(Assertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    ProtocolTargetingMode mode = ProtocolTypeMetadataSupport.targetingMode(assertion.getClass());
    return switch (mode) {
      case STATIC ->
          staticAllowedTargetTypesForType(assertion.getClass().asSubclass(Assertion.class));
      case ANALYSIS_QUERY -> AnalysisAssertion.allowedTargetTypes((AnalysisAssertion) assertion);
      case NESTED_ASSERTION ->
          CompositeAssertion.allowedTargetTypes((CompositeAssertion) assertion);
      case INTERSECTION_OF_NESTED_ASSERTIONS ->
          CompositeAssertion.allowedTargetTypes((CompositeAssertion) assertion);
    };
  }

  /** Returns the static selector types accepted by one assertion type. */
  static Class<? extends Selector>[] staticAllowedTargetTypesForType(
      Class<? extends Assertion> assertionType) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    if (!assertionType.isRecord()) {
      throw new IllegalArgumentException(
          "No target-type mapping configured for assertion class " + assertionType.getName());
    }
    Optional<String> dynamicRule = dynamicTargetSelectorRuleForType(assertionType);
    if (dynamicRule.isPresent()) {
      throw new IllegalArgumentException(
          "Assertion type "
              + assertionType.getName()
              + " derives target selectors dynamically: "
              + dynamicRule.orElseThrow());
    }
    return ProtocolTypeMetadataSupport.staticTargetSelectors(assertionType);
  }

  /** Returns the dynamic selector rule for one assertion type when selector families derive. */
  static Optional<String> dynamicTargetSelectorRuleForType(
      Class<? extends Assertion> assertionType) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    if (!assertionType.isRecord()) {
      return Optional.empty();
    }
    ProtocolTargetingMode mode = ProtocolTypeMetadataSupport.targetingMode(assertionType);
    return switch (mode) {
      case STATIC -> Optional.empty();
      case ANALYSIS_QUERY, NESTED_ASSERTION, INTERSECTION_OF_NESTED_ASSERTIONS ->
          ProtocolTypeMetadataSupport.targetSelectorRule(assertionType);
    };
  }

  /** Returns the selector rule exposed to discovery tooling for one assertion type. */
  static Optional<String> targetSelectorRuleForType(Class<? extends Assertion> assertionType) {
    Objects.requireNonNull(assertionType, "assertionType must not be null");
    if (!assertionType.isRecord()) {
      return Optional.empty();
    }
    return ProtocolTypeMetadataSupport.targetSelectorRule(assertionType);
  }
}
