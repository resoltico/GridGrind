package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects one or more sheet-local charts. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartSelector.AllOnSheet.class, name = "CHART_ALL_ON_SHEET"),
  @JsonSubTypes.Type(value = ChartSelector.ByName.class, name = "CHART_BY_NAME")
})
public sealed interface ChartSelector extends Selector
    permits ChartSelector.AllOnSheet, ChartSelector.ByName {

  /** Selects every chart on one sheet. */
  record AllOnSheet(SheetName sheetName) implements ChartSelector {
    public AllOnSheet {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one chart by sheet-local chart name. */
  record ByName(SheetName sheetName, String chartName) implements ChartSelector {
    public ByName {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      chartName = SelectorValueValidation.requireNonBlank(chartName, "chartName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
