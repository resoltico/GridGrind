package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual array-formula group metadata returned by workbook inspection. */
public record ArrayFormulaReport(
    String sheetName, String range, String topLeftAddress, String formula, boolean singleCell) {
  public ArrayFormulaReport {
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
