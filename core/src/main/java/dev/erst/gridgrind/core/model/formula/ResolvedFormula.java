package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;

/** Materialized formula text authored through the formula parser. */
public record ResolvedFormula(String value) implements ResolvedCellWrite {
  public ResolvedFormula {
    value = FormulaTextValidation.requireNormalFormulaBody(value, "value");
  }
}
