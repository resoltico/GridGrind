package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selects one or more sheet-local charts. */
@Union(discriminator = "type", catalog = true, displayOrder = 38)
public sealed interface ChartSelector extends Selector
    permits dev.erst.gridgrind.core.plan.ChartSelector.AllOnSheet,
        dev.erst.gridgrind.core.plan.ChartSelector.ByName {

  /** Selects every chart on one sheet. */
  @Variant(id = "CHART_ALL_ON_SHEET", summary = "Select every chart on one sheet.")
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
  @Variant(id = "CHART_BY_NAME", summary = "Select one chart by exact sheet-local chart name.")
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
