package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.selector.TableSelector;
import java.util.Objects;

/** Table-scoped fluent target. */
public final class TableTarget {
  private final TableSelector selector;

  TableTarget(TableSelector selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  TableSelector selector() {
    return selector;
  }

  /** Returns one exact table-row target by zero-based table row index. */
  public TableRowTarget row(int rowIndex) {
    return new TableRowTarget(
        new dev.erst.gridgrind.contract.selector.TableRowSelector.ByIndex(selector, rowIndex));
  }

  /** Returns one exact table-row target by logical key-column match. */
  public TableRowTarget rowByKey(String columnName, Values.CellValue expectedValue) {
    return new TableRowTarget(
        new dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell(
            selector, columnName, Values.toCellInput(expectedValue)));
  }

  /** Returns one set-table mutation step. */
  public PlannedMutation define(Tables.Definition table) {
    return new PlannedMutation(
        selector, new StructuredMutationAction.SetTable(Tables.toTableInput(table)));
  }

  /** Returns one delete-table mutation step. */
  public PlannedMutation delete() {
    return new PlannedMutation(selector, new StructuredMutationAction.DeleteTable());
  }

  /** Returns one table inspection step. */
  public PlannedInspection inspect() {
    return new PlannedInspection(selector, Queries.tables());
  }

  /** Returns one table-health analysis step. */
  public PlannedInspection analyzeHealth() {
    return new PlannedInspection(selector, Queries.tableHealth());
  }

  /** Returns one table presence assertion step. */
  public PlannedAssertion present() {
    return new PlannedAssertion(selector, Checks.tablePresent());
  }

  /** Returns one table absence assertion step. */
  public PlannedAssertion absent() {
    return new PlannedAssertion(selector, Checks.tableAbsent());
  }
}
