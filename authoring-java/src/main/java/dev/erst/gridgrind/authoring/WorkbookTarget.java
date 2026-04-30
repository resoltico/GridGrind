package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.Objects;

/** Workbook-scoped fluent target. */
public final class WorkbookTarget {
  private final WorkbookSelector selector;

  WorkbookTarget(WorkbookSelector selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  WorkbookSelector selector() {
    return selector;
  }

  /** Returns one workbook-summary inspection step. */
  public PlannedInspection summary() {
    return new PlannedInspection(selector, Queries.workbookSummary());
  }

  /** Returns one package-security inspection step. */
  public PlannedInspection packageSecurity() {
    return new PlannedInspection(selector, Queries.packageSecurity());
  }

  /** Returns one workbook-protection inspection step. */
  public PlannedInspection protection() {
    return new PlannedInspection(selector, Queries.workbookProtection());
  }

  /** Returns one aggregate workbook-findings analysis step. */
  public PlannedInspection findings() {
    return new PlannedInspection(selector, Queries.workbookFindings());
  }
}
