package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.vocabulary.TableName;
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
    // Each selected identifier requires its own canonical value validation.
    @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
    public ByNames {
      Objects.requireNonNull(names, "names must not be null");
      names = List.copyOf(names);
      if (names.isEmpty()) {
        throw new IllegalArgumentException("names must not be empty");
      }
      Set<String> unique = new LinkedHashSet<>();
      for (String name : names) {
        String normalized = new TableName(name).value();
        String key = normalized.toUpperCase(java.util.Locale.ROOT);
        if (!unique.add(key)) {
          throw new IllegalArgumentException("names must not contain duplicates");
        }
      }
    }
  }
}
