package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which cells a metadata read operation should inspect. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "mode")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellSelection.AllUsedCells.class, name = "ALL_USED_CELLS"),
  @JsonSubTypes.Type(value = CellSelection.Selected.class, name = "SELECTED")
})
public sealed interface CellSelection permits CellSelection.AllUsedCells, CellSelection.Selected {

  /** Selects every physically present cell on the sheet. */
  record AllUsedCells() implements CellSelection {}

  /** Selects only the exact A1 addresses named in the provided ordered list. */
  record Selected(List<String> addresses) implements CellSelection {
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
