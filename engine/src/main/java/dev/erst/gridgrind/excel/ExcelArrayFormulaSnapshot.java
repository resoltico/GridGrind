package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual snapshot of one array-formula group. */
public record ExcelArrayFormulaSnapshot(
    String sheetName, String range, String topLeftAddress, String formula, boolean singleCell) {
  public ExcelArrayFormulaSnapshot {
    sheetName = requireNonBlank(sheetName, "sheetName");
    range = requireNonBlank(range, "range");
    topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
    formula = requireNonBlank(formula, "formula");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
