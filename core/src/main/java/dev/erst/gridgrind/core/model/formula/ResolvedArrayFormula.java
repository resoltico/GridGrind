package dev.erst.gridgrind.core.model.formula;

import java.util.Objects;

/** Materialized array-formula text ready for workbook authoring. */
public record ResolvedArrayFormula(String formula) {
  public ResolvedArrayFormula {
    Objects.requireNonNull(formula, "formula must not be null");
    if (formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
  }
}
