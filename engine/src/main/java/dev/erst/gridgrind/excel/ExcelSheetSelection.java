package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which sheets a workbook-core read command should inspect or analyze. */
public sealed interface ExcelSheetSelection
    permits ExcelSheetSelection.All, ExcelSheetSelection.Selected {

  /** Selects every sheet in workbook order. */
  record All() implements ExcelSheetSelection {}

  /** Selects only the exact sheets named in the provided ordered list. */
  record Selected(List<String> sheetNames) implements ExcelSheetSelection {
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
    return copy;
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not contain nulls");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not contain blank values");
    }
  }
}
