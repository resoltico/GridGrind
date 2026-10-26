package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;

/** Materialized opaque OOXML formula character data. */
public record ResolvedRawFormula(String value) implements ResolvedCellWrite {
  public ResolvedRawFormula {
    value = FormulaTextValidation.requireRawFormulaBody(value, "value");
  }
}
