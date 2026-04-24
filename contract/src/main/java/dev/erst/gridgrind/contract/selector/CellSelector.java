package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Selects one or more cells, either on one sheet or by exact workbook-qualified addresses. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellSelector.AllUsedInSheet.class, name = "CELL_ALL_USED_IN_SHEET"),
  @JsonSubTypes.Type(value = CellSelector.ByAddress.class, name = "CELL_BY_ADDRESS"),
  @JsonSubTypes.Type(value = CellSelector.ByAddresses.class, name = "CELL_BY_ADDRESSES"),
  @JsonSubTypes.Type(
      value = CellSelector.ByQualifiedAddresses.class,
      name = "CELL_BY_QUALIFIED_ADDRESSES")
})
public sealed interface CellSelector extends Selector
    permits CellSelector.AllUsedInSheet,
        CellSelector.ByAddress,
        CellSelector.ByAddresses,
        CellSelector.ByQualifiedAddresses {

  /** Selects every physically present cell on one sheet. */
  record AllUsedInSheet(String sheetName) implements CellSelector {
    public AllUsedInSheet {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one exact cell on one sheet. */
  record ByAddress(String sheetName, String address) implements CellSelector {
    public ByAddress {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      address = SelectorSupport.requireAddress(address, "address");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more exact cells on one sheet. */
  record ByAddresses(String sheetName, List<String> addresses) implements CellSelector {
    public ByAddresses {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      addresses = SelectorSupport.copyDistinctAddresses(addresses, "addresses");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects exact cells across one or more sheets. */
  record ByQualifiedAddresses(List<QualifiedAddress> cells) implements CellSelector {
    public ByQualifiedAddresses {
      cells = SelectorSupport.copyDistinctValues(cells, "cells");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** One workbook-qualified cell address. */
  record QualifiedAddress(String sheetName, String address) {
    public QualifiedAddress {
      sheetName = SelectorSupport.requireSheetName(sheetName, "sheetName");
      address = SelectorSupport.requireAddress(address, "address");
    }

    @Override
    public String toString() {
      return sheetName + "!" + address;
    }
  }
}
