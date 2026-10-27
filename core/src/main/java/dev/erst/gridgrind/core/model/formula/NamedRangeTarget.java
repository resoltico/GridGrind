package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Protocol-facing explicit sheet range or formula target for named-range authoring. */
@Union(discriminator = "type", catalog = true, displayOrder = 23)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_NAMED_RANGE_FORMULA_INJECTION
})
public sealed interface NamedRangeTarget
    permits dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Range,
        dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Formula {
  /** Creates a sheet-local cell or rectangular range target. */
  static Range range(SheetName sheetName, RangeRef range) {
    return new Range(sheetName, range);
  }

  /** Creates a formula-defined target that is stored exactly as authored. */
  static Formula formula(String formula) {
    return new Formula(formula);
  }

  /** Sheet-qualified explicit cell or rectangular range target. */
  @Variant(
      id = "RANGE",
      summary =
          "Write one named-range target backed by an explicit sheet-qualified cell or rectangular range.")
  record Range(SheetName sheetName, RangeRef range) implements NamedRangeTarget {
    public Range {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
    }
  }

  /** Formula-defined named-range target stored exactly as authored. */
  @Variant(
      id = "FORMULA",
      summary = "Write one named-range target backed by an exact stored formula.")
  record Formula(String formula) implements NamedRangeTarget {
    public Formula {
      formula = TextValidation.requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula);
    }
  }
}
