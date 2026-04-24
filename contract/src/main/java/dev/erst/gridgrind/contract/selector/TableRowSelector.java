package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.CellInput;
import java.util.Objects;

/** Selects one or more logical rows within one selected table. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableRowSelector.AllRows.class, name = "TABLE_ROW_ALL_ROWS"),
  @JsonSubTypes.Type(value = TableRowSelector.ByIndex.class, name = "TABLE_ROW_BY_INDEX"),
  @JsonSubTypes.Type(value = TableRowSelector.ByKeyCell.class, name = "TABLE_ROW_BY_KEY_CELL")
})
public sealed interface TableRowSelector extends Selector
    permits TableRowSelector.AllRows, TableRowSelector.ByIndex, TableRowSelector.ByKeyCell {

  /** Selects every data row in one selected table. */
  record AllRows(TableSelector table) implements TableRowSelector {
    public AllRows {
      Objects.requireNonNull(table, "table must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one zero-based data row by index within one selected table. */
  record ByIndex(TableSelector table, int rowIndex) implements TableRowSelector {
    public ByIndex {
      Objects.requireNonNull(table, "table must not be null");
      requireExactTableSelector(table, "table");
      rowIndex = SelectorSupport.requireNonNegative(rowIndex, "rowIndex");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one logical data row by matching one key-column cell within one selected table. */
  record ByKeyCell(TableSelector table, String columnName, CellInput expectedValue)
      implements TableRowSelector {
    public ByKeyCell {
      Objects.requireNonNull(table, "table must not be null");
      requireExactTableSelector(table, "table");
      columnName = SelectorSupport.requireNonBlank(columnName, "columnName");
      expectedValue = requireSupportedKeyValue(expectedValue);
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ZERO_OR_ONE;
    }
  }

  private static TableSelector requireExactTableSelector(TableSelector table, String fieldName) {
    if (table instanceof TableSelector.All || table instanceof TableSelector.ByNames) {
      throw new IllegalArgumentException(
          fieldName + " must resolve exactly one table for row selection");
    }
    return table;
  }

  private static CellInput requireSupportedKeyValue(CellInput expectedValue) {
    Objects.requireNonNull(expectedValue, "expectedValue must not be null");
    if (expectedValue instanceof CellInput.Blank
        || expectedValue instanceof CellInput.Text
        || expectedValue instanceof CellInput.Numeric
        || expectedValue instanceof CellInput.BooleanValue
        || expectedValue instanceof CellInput.Formula) {
      return expectedValue;
    }
    throw new IllegalArgumentException(
        "expectedValue must be BLANK, TEXT, NUMBER, BOOLEAN, or FORMULA for row key selection");
  }
}
