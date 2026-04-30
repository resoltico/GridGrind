package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import java.util.Objects;

/** Named-range fluent target. */
public final class NamedRangeTarget {
  private final NamedRangeSelector selector;

  NamedRangeTarget(NamedRangeSelector selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  NamedRangeSelector selector() {
    return selector;
  }

  /** Returns one delete-named-range mutation step. */
  public PlannedMutation delete() {
    return new PlannedMutation(selector, new StructuredMutationAction.DeleteNamedRange());
  }

  /** Returns one named-range inspection step. */
  public PlannedInspection inspect() {
    return new PlannedInspection(selector, Queries.namedRanges());
  }

  /** Returns one named-range-surface inspection step. */
  public PlannedInspection surface() {
    return new PlannedInspection(selector, Queries.namedRangeSurface());
  }

  /** Returns one named-range-health analysis step. */
  public PlannedInspection analyzeHealth() {
    return new PlannedInspection(selector, Queries.namedRangeHealth());
  }

  /** Returns one named-range presence assertion step. */
  public PlannedAssertion present() {
    return new PlannedAssertion(selector, Checks.namedRangePresent());
  }

  /** Returns one named-range absence assertion step. */
  public PlannedAssertion absent() {
    return new PlannedAssertion(selector, Checks.namedRangeAbsent());
  }
}
