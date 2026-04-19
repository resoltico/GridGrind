package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which A1-style ranges on one sheet a document-intelligence read should inspect. */
public sealed interface ExcelRangeSelection
    permits ExcelRangeSelection.All, ExcelRangeSelection.Selected {

  /** Selects every matching structure on the sheet. */
  record All() implements ExcelRangeSelection {}

  /** Selects only structures whose ranges intersect the provided A1-style ranges. */
  record Selected(List<String> ranges) implements ExcelRangeSelection {
    public Selected {
      ranges = copyRanges(ranges);
    }
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (String range : copy) {
      requireNonBlank(range, "ranges");
      if (!unique.add(range)) {
        throw new IllegalArgumentException("ranges must not contain duplicates");
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
