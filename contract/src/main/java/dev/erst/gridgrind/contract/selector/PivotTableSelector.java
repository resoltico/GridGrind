package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more pivot tables by workbook-global identity. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PivotTableSelector.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = PivotTableSelector.ByName.class, name = "BY_NAME"),
  @JsonSubTypes.Type(value = PivotTableSelector.ByNames.class, name = "BY_NAMES"),
  @JsonSubTypes.Type(value = PivotTableSelector.ByNameOnSheet.class, name = "BY_NAME_ON_SHEET")
})
public sealed interface PivotTableSelector extends Selector
    permits PivotTableSelector.All,
        PivotTableSelector.ByName,
        PivotTableSelector.ByNames,
        PivotTableSelector.ByNameOnSheet {

  /** Selects every pivot table in workbook order. */
  record All() implements PivotTableSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one pivot table by workbook-global pivot name. */
  record ByName(String name) implements PivotTableSelector {
    public ByName {
      name = SelectorSupport.requirePivotTableName(name, "name");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more pivot tables by workbook-global pivot names. */
  record ByNames(List<String> names) implements PivotTableSelector {
    public ByNames {
      names = SelectorSupport.copyDistinctPivotTableNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one pivot table by workbook-global name and asserts its owning sheet. */
  record ByNameOnSheet(String name, String sheetName) implements PivotTableSelector {
    public ByNameOnSheet {
      name = SelectorSupport.requirePivotTableName(name, "name");
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
