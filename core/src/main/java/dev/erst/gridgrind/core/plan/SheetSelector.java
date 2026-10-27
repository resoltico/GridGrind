package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;

/** Selects one or more sheets. */
@Union(discriminator = "type", catalog = true, displayOrder = 26)
public sealed interface SheetSelector extends Selector
    permits dev.erst.gridgrind.core.plan.SheetSelector.All,
        dev.erst.gridgrind.core.plan.SheetSelector.ByName,
        dev.erst.gridgrind.core.plan.SheetSelector.ByNames {

  /** Selects every sheet in workbook order. */
  @Variant(id = "SHEET_ALL", summary = "Select every sheet in workbook order.")
  record All() implements SheetSelector, ResolvedTarget {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one sheet by exact sheet name. */
  @Variant(id = "SHEET_BY_NAME", summary = "Select one exact sheet by name.")
  record ByName(SheetName name) implements SheetSelector, ResolvedTarget {
    public ByName {
      java.util.Objects.requireNonNull(name, "name must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more sheets by exact sheet names. */
  @Variant(id = "SHEET_BY_NAMES", summary = "Select one or more exact sheets by ordered name list.")
  record ByNames(List<SheetName> names) implements SheetSelector, ResolvedTarget {
    public ByNames {
      names = SelectorListValidation.copyDistinctSheetNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }
}
