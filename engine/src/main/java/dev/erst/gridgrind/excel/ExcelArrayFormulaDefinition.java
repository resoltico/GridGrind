package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core array-formula definition. */
public record ExcelArrayFormulaDefinition(String formula) {
  public ExcelArrayFormulaDefinition {
    Objects.requireNonNull(formula, "formula must not be null");
    if (formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
  }
}
