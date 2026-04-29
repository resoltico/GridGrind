package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import java.io.IOException;

/** Resolves selectors whose nested authored values can come from bound external input sources. */
final class SourceBackedSelectorResolver {
  private SourceBackedSelectorResolver() {}

  static Selector resolve(Selector selector, ExecutionInputBindings bindings) throws IOException {
    if (selector instanceof TableCellSelector.ByColumnName tableCell) {
      TableRowSelector resolvedRow = (TableRowSelector) resolve(tableCell.row(), bindings);
      return sameReference(resolvedRow, tableCell.row())
          ? tableCell
          : new TableCellSelector.ByColumnName(resolvedRow, tableCell.columnName());
    }
    if (selector instanceof TableRowSelector.ByKeyCell byKeyCell) {
      CellInput resolvedValue =
          SourceBackedPlanResolver.resolveCellInput(byKeyCell.expectedValue(), bindings);
      return sameReference(resolvedValue, byKeyCell.expectedValue())
          ? byKeyCell
          : new TableRowSelector.ByKeyCell(
              byKeyCell.table(), byKeyCell.columnName(), resolvedValue);
    }
    return selector;
  }
}
