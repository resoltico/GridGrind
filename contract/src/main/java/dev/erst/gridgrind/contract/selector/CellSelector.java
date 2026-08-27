package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more cells on one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellSelector.AllUsedInSheet.class, name = "CELL_ALL_USED_IN_SHEET"),
  @JsonSubTypes.Type(value = CellSelector.ByAddress.class, name = "CELL_BY_ADDRESS"),
  @JsonSubTypes.Type(value = CellSelector.ByAddresses.class, name = "CELL_BY_ADDRESSES")
})
public sealed interface CellSelector extends Selector
    permits CellSelector.AllUsedInSheet, CellSelector.ByAddress, CellSelector.ByAddresses {

  /** Selects every physically present cell on one sheet. */
  record AllUsedInSheet(String sheetName) implements CellSelector {
    public AllUsedInSheet {
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one exact cell on one sheet. */
  record ByAddress(String sheetName, String address) implements CellSelector {
    public ByAddress {
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
      address = SelectorValueValidation.requireAddress(address, "address");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more exact cells on one sheet. */
  record ByAddresses(String sheetName, List<String> addresses) implements CellSelector {
    public ByAddresses {
      sheetName = SelectorValueValidation.requireSheetName(sheetName, "sheetName");
      addresses = SelectorListValidation.copyDistinctAddresses(addresses, "addresses");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }
}
