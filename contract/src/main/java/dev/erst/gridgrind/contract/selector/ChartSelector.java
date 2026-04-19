package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Selects one or more sheet-local charts. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartSelector.AllOnSheet.class, name = "ALL_ON_SHEET"),
  @JsonSubTypes.Type(value = ChartSelector.ByName.class, name = "BY_NAME")
})
public sealed interface ChartSelector extends Selector
    permits ChartSelector.AllOnSheet, ChartSelector.ByName {

  /** Selects every chart on one sheet. */
  record AllOnSheet(String sheetName) implements ChartSelector {
    public AllOnSheet {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one chart by sheet-local chart name. */
  record ByName(String sheetName, String chartName) implements ChartSelector {
    public ByName {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      chartName = SelectorSupport.requireNonBlank(chartName, "chartName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
