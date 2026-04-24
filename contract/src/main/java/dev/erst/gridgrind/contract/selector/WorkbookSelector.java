package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Selects the single workbook currently being executed. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes(@JsonSubTypes.Type(value = WorkbookSelector.Current.class, name = "WORKBOOK_CURRENT"))
public sealed interface WorkbookSelector extends Selector permits WorkbookSelector.Current {

  /** Targets the only workbook currently in execution. */
  record Current() implements WorkbookSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
