package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;

/** Protocol-facing explicit cell or rectangular range target for named-range authoring. */
public record NamedRangeTarget(String sheetName, String range, String formula) {
  /** Creates a sheet-local cell or rectangular range target. */
  public NamedRangeTarget(String sheetName, String range) {
    this(sheetName, range, null);
  }

  /** Creates a formula-defined target that is stored exactly as authored. */
  public NamedRangeTarget(String formula) {
    this(null, null, formula);
  }

  public NamedRangeTarget {
    if (formula != null) {
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      if (sheetName != null || range != null) {
        throw new IllegalArgumentException(
            "formula-defined named-range targets must not also set sheetName or range");
      }
    } else {
      ExcelSheetNames.requireValid(sheetName, "sheetName");
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }
}
