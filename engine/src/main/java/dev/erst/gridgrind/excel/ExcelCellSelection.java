package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

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
    Objects.requireNonNull(addresses, "addresses must not be null");
    List<String> copy = List.copyOf(addresses);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("addresses must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String address : copy) {
      requireNonBlank(address, "addresses");
      if (!unique.add(address)) {
        throw new IllegalArgumentException("addresses must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not contain nulls");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not contain blank values");
    }
  }
}
