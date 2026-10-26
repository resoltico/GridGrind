package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Selects one logical cell within one selected table row. */
@Union(discriminator = "type", catalog = true, displayOrder = 33)
public sealed interface TableCellSelector extends Selector
    permits dev.erst.gridgrind.core.plan.TableCellSelector.ByColumnName {

  /** Selects one logical cell by table-row selector plus logical column name. */
  @Variant(
      id = "TABLE_CELL_BY_COLUMN_NAME",
      summary = "Select one logical cell within one selected table row by column name.")
  record ByColumnName(TableRowSelector row, String columnName) implements TableCellSelector {
    public ByColumnName {
      Objects.requireNonNull(row, "row must not be null");
      requireSingleRowSelector(row, "row");
      columnName = SelectorValueValidation.requireNonBlank(columnName, "columnName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ZERO_OR_ONE;
    }
  }

  private static TableRowSelector requireSingleRowSelector(TableRowSelector row, String fieldName) {
    if (row instanceof TableRowSelector.AllRows) {
      throw new IllegalArgumentException(
          fieldName + " must resolve at most one table row for table-cell selection");
    }
    return row;
  }
}
