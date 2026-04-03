package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Workbook-level selector for one or more tables. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableSelection.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = TableSelection.ByNames.class, name = "BY_NAMES")
})
public sealed interface TableSelection permits TableSelection.All, TableSelection.ByNames {

  /** Select every table in workbook order. */
  record All() implements TableSelection {}

  /** Select only the supplied workbook-global table names. */
  record ByNames(List<String> names) implements TableSelection {
    public ByNames {
      Objects.requireNonNull(names, "names must not be null");
      names = List.copyOf(names);
      if (names.isEmpty()) {
        throw new IllegalArgumentException("names must not be empty");
      }
      Set<String> unique = new LinkedHashSet<>();
      for (String name : names) {
        String validated = ProtocolDefinedNameValidation.validateName(name);
        if (!unique.add(validated.toUpperCase(Locale.ROOT))) {
          throw new IllegalArgumentException("names must not contain duplicates");
        }
      }
    }
  }
}
