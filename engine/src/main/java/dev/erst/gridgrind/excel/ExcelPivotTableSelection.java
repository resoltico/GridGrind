package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Workbook-level selector for one or more pivot tables. */
public sealed interface ExcelPivotTableSelection
    permits ExcelPivotTableSelection.All, ExcelPivotTableSelection.ByNames {

  /** Select every pivot table in workbook sheet order. */
  record All() implements ExcelPivotTableSelection {}

  /** Select only the supplied workbook-global pivot table names. */
  record ByNames(List<String> names) implements ExcelPivotTableSelection {
    public ByNames {
      Objects.requireNonNull(names, "names must not be null");
      names = List.copyOf(names);
      if (names.isEmpty()) {
        throw new IllegalArgumentException("names must not be empty");
      }
      Set<String> unique = new LinkedHashSet<>();
      for (String name : names) {
        String normalized = ExcelPivotTableNaming.validateName(name);
        String key = normalized.toUpperCase(java.util.Locale.ROOT);
        if (!unique.add(key)) {
          throw new IllegalArgumentException("names must not contain duplicates");
        }
      }
    }
  }
}
