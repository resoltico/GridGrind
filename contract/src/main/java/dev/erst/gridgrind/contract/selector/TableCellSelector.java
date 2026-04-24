package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Selects one logical cell within one selected table row. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes(
    @JsonSubTypes.Type(
        value = TableCellSelector.ByColumnName.class,
        name = "TABLE_CELL_BY_COLUMN_NAME"))
public sealed interface TableCellSelector extends Selector permits TableCellSelector.ByColumnName {

  /** Selects one logical cell by table-row selector plus logical column name. */
  record ByColumnName(TableRowSelector row, String columnName) implements TableCellSelector {
    public ByColumnName {
      Objects.requireNonNull(row, "row must not be null");
      requireSingleRowSelector(row, "row");
      columnName = SelectorSupport.requireNonBlank(columnName, "columnName");
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
