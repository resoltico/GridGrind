package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Selects which named ranges a workbook-core read command should inspect or analyze. */
public sealed interface ExcelNamedRangeSelection
    permits ExcelNamedRangeSelection.All, ExcelNamedRangeSelection.Selected {

  /** Selects every user-facing named range in the workbook. */
  record All() implements ExcelNamedRangeSelection {}

  /** Selects only the named ranges matched by exact selectors. */
  record Selected(List<ExcelNamedRangeSelector> selectors) implements ExcelNamedRangeSelection {
    public Selected {
      selectors = copySelectors(selectors);
    }
  }

  private static List<ExcelNamedRangeSelector> copySelectors(
      List<ExcelNamedRangeSelector> selectors) {
    Objects.requireNonNull(selectors, "selectors must not be null");
    List<ExcelNamedRangeSelector> copy = List.copyOf(selectors);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("selectors must not be empty");
    }
    Set<ExcelNamedRangeSelector> unique = new LinkedHashSet<>();
    for (ExcelNamedRangeSelector selector : copy) {
      Objects.requireNonNull(selector, "selectors must not contain nulls");
      if (!unique.add(selector)) {
        throw new IllegalArgumentException("selectors must not contain duplicates");
      }
    }
    return copy;
  }
}
