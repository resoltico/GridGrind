package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Selects the single workbook currently being executed. */
@Union(discriminator = "type", catalog = true, displayOrder = 25)
public sealed interface WorkbookSelector extends Selector
    permits dev.erst.gridgrind.core.plan.WorkbookSelector.Current {

  /** Targets the only workbook currently in execution. */
  @Variant(id = "WORKBOOK_CURRENT", summary = "Target the workbook currently being executed.")
  record Current() implements WorkbookSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
