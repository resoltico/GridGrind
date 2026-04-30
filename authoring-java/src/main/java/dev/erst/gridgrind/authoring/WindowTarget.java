package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.selector.RangeSelector;
import java.util.Objects;

/** Window-scoped fluent target. */
public final class WindowTarget {
  private final RangeSelector.RectangularWindow selector;

  WindowTarget(RangeSelector.RectangularWindow selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  RangeSelector.RectangularWindow selector() {
    return selector;
  }

  /** Returns one rectangular-window inspection step. */
  public PlannedInspection read() {
    return new PlannedInspection(selector, Queries.window());
  }

  /** Returns one sheet-schema inspection step for this window. */
  public PlannedInspection schema() {
    return new PlannedInspection(selector, Queries.sheetSchema());
  }
}
