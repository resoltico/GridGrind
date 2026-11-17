package dev.erst.gridgrind.core.model.formula;

/** Materialized array-formula text ready for workbook authoring. */
public record ResolvedArrayFormula(String formula) {
  public ResolvedArrayFormula {
    formula = ArrayFormulaSpec.normalizeFormula(formula);
  }
}
