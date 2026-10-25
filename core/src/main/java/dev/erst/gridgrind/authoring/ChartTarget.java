package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.selector.ChartSelector;
import java.util.Objects;

/** Chart-scoped fluent target. */
public final class ChartTarget {
  private final ChartSelector selector;

  ChartTarget(ChartSelector selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  ChartSelector selector() {
    return selector;
  }

  /** Returns one chart inspection step for the exact chart selector. */
  public PlannedInspection inspect() {
    return new PlannedInspection(selector, WorkbookAssetQueries.charts());
  }

  /** Returns one chart presence assertion step. */
  public PlannedAssertion present() {
    return new PlannedAssertion(selector, Checks.chartPresent());
  }

  /** Returns one chart absence assertion step. */
  public PlannedAssertion absent() {
    return new PlannedAssertion(selector, Checks.chartAbsent());
  }
}
