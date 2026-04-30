package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
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

  /** Returns one chart inventory inspection step on the owning sheet. */
  public PlannedInspection inspectOnSheet() {
    return switch (selector) {
      case ChartSelector.ByName byName ->
          new PlannedInspection(new SheetSelector.ByName(byName.sheetName()), Queries.charts());
      case ChartSelector.AllOnSheet allOnSheet ->
          new PlannedInspection(new SheetSelector.ByName(allOnSheet.sheetName()), Queries.charts());
    };
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
