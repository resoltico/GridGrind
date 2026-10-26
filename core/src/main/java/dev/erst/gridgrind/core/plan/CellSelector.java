package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;

/** Selects one or more cells on one sheet. */
@Union(discriminator = "type", catalog = true, displayOrder = 27)
public sealed interface CellSelector extends Selector
    permits dev.erst.gridgrind.core.plan.CellSelector.AllUsedInSheet,
        dev.erst.gridgrind.core.plan.CellSelector.ByAddress,
        dev.erst.gridgrind.core.plan.CellSelector.ByAddresses {

  /** Selects every physically present cell on one sheet. */
  @Variant(
      id = "CELL_ALL_USED_IN_SHEET",
      summary = "Select every physically present cell on one sheet.")
  record AllUsedInSheet(SheetName sheetName) implements CellSelector {
    public AllUsedInSheet {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one exact cell on one sheet. */
  @Variant(id = "CELL_BY_ADDRESS", summary = "Select one exact cell on one sheet.")
  record ByAddress(SheetName sheetName, CellAddress address) implements CellSelector {
    public ByAddress {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      java.util.Objects.requireNonNull(address, "address must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more exact cells on one sheet. */
  @Variant(id = "CELL_BY_ADDRESSES", summary = "Select one or more exact cells on one sheet.")
  record ByAddresses(SheetName sheetName, List<CellAddress> addresses) implements CellSelector {
    public ByAddresses {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      addresses = SelectorListValidation.copyDistinctAddresses(addresses, "addresses");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }
}
