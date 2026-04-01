package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelTableSelection;
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

  /** Converts this selector into the workbook-core table selection. */
  ExcelTableSelection toExcelTableSelection();

  /** Select every table in workbook order. */
  record All() implements TableSelection {
    @Override
    public ExcelTableSelection toExcelTableSelection() {
      return new ExcelTableSelection.All();
    }
  }

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
        String validated = dev.erst.gridgrind.excel.ExcelTableDefinition.validateName(name);
        if (!unique.add(validated.toUpperCase(Locale.ROOT))) {
          throw new IllegalArgumentException("names must not contain duplicates");
        }
      }
    }

    @Override
    public ExcelTableSelection toExcelTableSelection() {
      return new ExcelTableSelection.ByNames(names);
    }
  }
}
