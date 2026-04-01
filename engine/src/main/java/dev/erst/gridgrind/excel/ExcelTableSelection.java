package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Workbook-level selector for one or more tables. */
public sealed interface ExcelTableSelection
    permits ExcelTableSelection.All, ExcelTableSelection.ByNames {

  /** Select every table in workbook sheet order. */
  record All() implements ExcelTableSelection {}

  /** Select only the supplied workbook-global table names. */
  record ByNames(List<String> names) implements ExcelTableSelection {
    public ByNames {
      Objects.requireNonNull(names, "names must not be null");
      names = List.copyOf(names);
      if (names.isEmpty()) {
        throw new IllegalArgumentException("names must not be empty");
      }
      Set<String> unique = new LinkedHashSet<>();
      for (String name : names) {
        String normalized = ExcelTableDefinition.validateName(name);
        String key = normalized.toUpperCase(java.util.Locale.ROOT);
        if (!unique.add(key)) {
          throw new IllegalArgumentException("names must not contain duplicates");
        }
      }
    }
  }
}
