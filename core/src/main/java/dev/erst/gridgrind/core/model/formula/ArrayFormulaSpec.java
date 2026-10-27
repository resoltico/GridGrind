package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Objects;

/** Dedicated array-formula input used for single-cell or multi-cell array authoring. */
@CatalogType(
    id = "ArrayFormulaSpec",
    summary =
        "One authored array formula bound to a contiguous single-cell or multi-cell range. Leading = or {=...} wrappers normalize away for inline sources.",
    displayOrder = 30)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_DDE_REJECTION_IN_ALL_FORMULA_INPUTS
})
public record ArrayFormulaSpec(TextSourceSpec source) {
  public ArrayFormulaSpec {
    TextSourceSpec normalized = Objects.requireNonNull(source, "source must not be null");
    if (normalized instanceof TextSourceSpec.Inline inline) {
      String formula = inline.text();
      FormulaInputSecurity.rejectDde(formula); // check before stripping
      if (formula.startsWith("{=") && formula.endsWith("}")) {
        formula = formula.substring(2, formula.length() - 1);
      } else if (formula.startsWith("=")) {
        formula = formula.substring(1);
      }
      if (formula.isBlank()) {
        throw new IllegalArgumentException(
            "source must not be blank after stripping formula syntax");
      }
      normalized = new TextSourceSpec.Inline(formula);
    }
    source = normalized;
  }
}
