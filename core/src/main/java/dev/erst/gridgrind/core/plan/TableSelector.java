package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;

/** Selects one or more tables by workbook-global identity. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableSelector.All.class, name = "TABLE_ALL"),
  @JsonSubTypes.Type(value = TableSelector.ByName.class, name = "TABLE_BY_NAME"),
  @JsonSubTypes.Type(value = TableSelector.ByNames.class, name = "TABLE_BY_NAMES"),
  @JsonSubTypes.Type(value = TableSelector.ByNameOnSheet.class, name = "TABLE_BY_NAME_ON_SHEET")
})
public sealed interface TableSelector extends Selector
    permits TableSelector.All,
        TableSelector.ByName,
        TableSelector.ByNames,
        TableSelector.ByNameOnSheet {

  /** Selects every table in workbook order. */
  record All() implements TableSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one table by workbook-global name. */
  record ByName(TableName name) implements TableSelector {
    public ByName {
      java.util.Objects.requireNonNull(name, "name must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more tables by workbook-global names. */
  record ByNames(List<TableName> names) implements TableSelector {
    public ByNames {
      names = SelectorListValidation.copyDistinctTableNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one workbook-global table and asserts its expected owning sheet. */
  record ByNameOnSheet(TableName name, SheetName sheetName) implements TableSelector {
    public ByNameOnSheet {
      java.util.Objects.requireNonNull(name, "name must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
