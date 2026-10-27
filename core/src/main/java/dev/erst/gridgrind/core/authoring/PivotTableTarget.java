package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import java.util.Objects;

/** Pivot-table-scoped fluent target. */
public final class PivotTableTarget {
  private final PivotTableSelector selector;

  PivotTableTarget(PivotTableSelector selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  PivotTableSelector selector() {
    return selector;
  }

  /** Returns one delete-pivot-table mutation step. */
  public PlannedMutation delete() {
    return new PlannedMutation(selector, new StructuredMutationAction.DeletePivotTable());
  }

  /** Returns one pivot-table inspection step. */
  public PlannedInspection inspect() {
    return new PlannedInspection(selector, WorkbookAssetQueries.pivotTables());
  }

  /** Returns one pivot-table-health analysis step. */
  public PlannedInspection analyzeHealth() {
    return new PlannedInspection(selector, InspectionAnalysisQueries.pivotTableHealth());
  }

  /** Returns one pivot-table presence assertion step. */
  public PlannedAssertion present() {
    return new PlannedAssertion(selector, Checks.pivotTablePresent());
  }

  /** Returns one pivot-table absence assertion step. */
  public PlannedAssertion absent() {
    return new PlannedAssertion(selector, Checks.pivotTableAbsent());
  }
}
