package dev.erst.gridgrind.contract.assertion;

import dev.erst.gridgrind.contract.catalog.ProtocolTargetingMode;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Boolean combinator assertions layered over nested assertions. */
public sealed interface CompositeAssertion extends Assertion
    permits CompositeAssertion.AllOf, CompositeAssertion.AnyOf, CompositeAssertion.Not {

  String INTERSECTION_RULE =
      "Matches the intersection of every nested assertion's target selectors.";
  String NESTED_RULE = "Matches the nested assertion's target selectors.";

  @ProtocolTypeMetadata(
      id = "ALL_OF",
      summary = "Require every nested assertion to pass against the same step target.",
      targetingMode = ProtocolTargetingMode.INTERSECTION_OF_NESTED_ASSERTIONS,
      targetSelectorRule = INTERSECTION_RULE)
  record AllOf(List<Assertion> assertions) implements CompositeAssertion {
    public AllOf {
      assertions = AssertionSupport.copyAssertions(assertions, "assertions");
    }
  }

  @ProtocolTypeMetadata(
      id = "ANY_OF",
      summary = "Require at least one nested assertion to pass against the same step target.",
      targetingMode = ProtocolTargetingMode.INTERSECTION_OF_NESTED_ASSERTIONS,
      targetSelectorRule = INTERSECTION_RULE)
  record AnyOf(List<Assertion> assertions) implements CompositeAssertion {
    public AnyOf {
      assertions = AssertionSupport.copyAssertions(assertions, "assertions");
    }
  }

  @ProtocolTypeMetadata(
      id = "NOT",
      summary = "Invert one nested assertion against the same step target.",
      targetingMode = ProtocolTargetingMode.NESTED_ASSERTION,
      targetSelectorRule = NESTED_RULE)
  record Not(Assertion assertion) implements CompositeAssertion {
    public Not {
      Objects.requireNonNull(assertion, "assertion must not be null");
    }
  }

  /** Returns the selector types accepted by one composite assertion instance. */
  static Class<? extends Selector>[] allowedTargetTypes(CompositeAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return switch (assertion) {
      case AllOf allOf -> commonTargetTypes(allOf.assertions(), allOf.assertionType());
      case AnyOf anyOf -> commonTargetTypes(anyOf.assertions(), anyOf.assertionType());
      case Not not -> Assertion.allowedTargetTypes(not.assertion());
    };
  }

  /** Returns the shared selector-family intersection across nested assertions. */
  static Class<? extends Selector>[] commonTargetTypes(
      Iterable<Assertion> assertions, String compositeType) {
    var iterator = assertions.iterator();
    if (!iterator.hasNext()) {
      throw new IllegalArgumentException(
          compositeType + " requires nested assertions with compatible target families");
    }
    Class<? extends Selector>[] intersection =
        Assertion.allowedTargetTypes(iterator.next()).clone();
    while (iterator.hasNext()) {
      intersection = intersect(intersection, Assertion.allowedTargetTypes(iterator.next()));
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
