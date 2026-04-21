package dev.erst.gridgrind.excel;

import java.util.List;

/** Selects which cells a metadata read command should inspect on one sheet. */
public sealed interface ExcelCellSelection
    permits ExcelCellSelection.AllUsedCells, ExcelCellSelection.Selected {

  /** Selects every physically present cell on the sheet. */
  record AllUsedCells() implements ExcelCellSelection {}

  /** Selects only the exact A1 addresses named in the provided ordered list. */
  record Selected(List<String> addresses) implements ExcelCellSelection {
    public Selected {
      addresses = copyAddresses(addresses);
    }
  }

  private static List<String> copyAddresses(List<String> addresses) {
    return ExcelAddressLists.copyNonEmptyDistinctAddresses(addresses);
  }
}
