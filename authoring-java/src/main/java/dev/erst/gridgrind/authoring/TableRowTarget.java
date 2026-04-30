package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import java.util.Objects;

/** Table-row fluent target used to address one logical cell by column name. */
public final class TableRowTarget {
  private final TableRowSelector selector;

  TableRowTarget(TableRowSelector selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  TableRowSelector selector() {
    return selector;
  }

  /** Returns one exact logical table-cell target by column name. */
  public TableCellTarget cell(String columnName) {
    return new TableCellTarget(new TableCellSelector.ByColumnName(selector, columnName));
  }
}
