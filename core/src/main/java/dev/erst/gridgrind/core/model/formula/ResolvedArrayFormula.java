package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Materialized array-formula text ready for workbook authoring. */
public record ResolvedArrayFormula(String formula) {
  public ResolvedArrayFormula {
    formula = TextValidation.requireNonBlank(formula, "formula");
  }
}
