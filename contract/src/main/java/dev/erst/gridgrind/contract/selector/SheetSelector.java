package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more sheets. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = SheetSelector.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = SheetSelector.ByName.class, name = "BY_NAME"),
  @JsonSubTypes.Type(value = SheetSelector.ByNames.class, name = "BY_NAMES")
})
public sealed interface SheetSelector extends Selector
    permits SheetSelector.All, SheetSelector.ByName, SheetSelector.ByNames {

  /** Selects every sheet in workbook order. */
  record All() implements SheetSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one sheet by exact sheet name. */
  record ByName(String name) implements SheetSelector {
    public ByName {
      name = SelectorSupport.requireSheetName(name, "name");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more sheets by exact sheet names. */
  record ByNames(List<String> names) implements SheetSelector {
    public ByNames {
      names = SelectorSupport.copyDistinctSheetNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }
}
