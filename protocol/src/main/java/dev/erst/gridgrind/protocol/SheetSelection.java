package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which sheets a read operation should inspect or analyze. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = SheetSelection.All.class, name = "ALL"),
  @JsonSubTypes.Type(value = SheetSelection.Selected.class, name = "SELECTED")
})
public sealed interface SheetSelection permits SheetSelection.All, SheetSelection.Selected {

  /** Selects every sheet in workbook order. */
  record All() implements SheetSelection {}

  /** Selects only the exact sheets named in the provided ordered list. */
  record Selected(List<String> sheetNames) implements SheetSelection {
    public Selected {
      sheetNames = copySheetNames(sheetNames);
    }
  }

  private static List<String> copySheetNames(List<String> sheetNames) {
    Objects.requireNonNull(sheetNames, "sheetNames must not be null");
    List<String> copy = List.copyOf(sheetNames);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("sheetNames must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String sheetName : copy) {
      requireNonBlank(sheetName, "sheetNames");
      if (!unique.add(sheetName)) {
        throw new IllegalArgumentException("sheetNames must not contain duplicates");
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
