package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more tables by workbook-global identity. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableSelector.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = TableSelector.ByName.class, name = "BY_NAME"),
  @JsonSubTypes.Type(value = TableSelector.ByNames.class, name = "BY_NAMES"),
  @JsonSubTypes.Type(value = TableSelector.ByNameOnSheet.class, name = "BY_NAME_ON_SHEET")
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
  record ByName(String name) implements TableSelector {
    public ByName {
      name = SelectorSupport.requireDefinedName(name, "name");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more tables by workbook-global names. */
  record ByNames(List<String> names) implements TableSelector {
    public ByNames {
      names = SelectorSupport.copyDistinctDefinedNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one workbook-global table and asserts its expected owning sheet. */
  record ByNameOnSheet(String name, String sheetName) implements TableSelector {
    public ByNameOnSheet {
      name = SelectorSupport.requireDefinedName(name, "name");
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
