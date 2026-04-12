package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelPivotTableNaming;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Workbook-level selector for one or more pivot tables. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PivotTableSelection.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = PivotTableSelection.ByNames.class, name = "BY_NAMES")
})
public sealed interface PivotTableSelection
    permits PivotTableSelection.All, PivotTableSelection.ByNames {

  /** Select every pivot table in workbook order. */
  record All() implements PivotTableSelection {}

  /** Select only the supplied workbook-global pivot table names. */
  record ByNames(List<String> names) implements PivotTableSelection {
    public ByNames {
      Objects.requireNonNull(names, "names must not be null");
      names = List.copyOf(names);
      if (names.isEmpty()) {
        throw new IllegalArgumentException("names must not be empty");
      }
      Set<String> unique = new LinkedHashSet<>();
      for (String name : names) {
        String validated = ExcelPivotTableNaming.validateName(name);
        if (!unique.add(validated.toUpperCase(Locale.ROOT))) {
          throw new IllegalArgumentException("names must not contain duplicates");
        }
      }
    }
  }
}
