package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.DefinedNameLabel;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Structured workbook-level analysis report for one defined name. */
@Union(discriminator = "type", catalog = true, displayOrder = 6)
public sealed interface NamedRangeReport
    permits dev.erst.gridgrind.core.model.formula.NamedRangeReport.RangeReport,
        dev.erst.gridgrind.core.model.formula.NamedRangeReport.FormulaReport {
  /** Defined-name identifier. */
  DefinedNameLabel name();

  /** Workbook or sheet scope of the defined name. */
  NamedRangeScopeReport scope();

  /** Exact formula text stored in the workbook for this defined name. */
  String refersToFormula();

  /** Named range that resolves cleanly to a sheet-qualified cell or rectangular range target. */
  @Variant(
      id = "RANGE",
      summary = "Exact named-range report that resolves to one typed workbook target.")
  record RangeReport(
      DefinedNameLabel name,
      NamedRangeScopeReport scope,
      String refersToFormula,
      NamedRangeTarget target)
      implements NamedRangeReport {
    public RangeReport {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      refersToFormula = TextValidation.requireNonBlank(refersToFormula, "refersToFormula");
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Defined name whose formula cannot be normalized to a typed range target. */
  @Variant(id = "FORMULA", summary = "Exact named-range report that remains formula-backed.")
  record FormulaReport(DefinedNameLabel name, NamedRangeScopeReport scope, String refersToFormula)
      implements NamedRangeReport {
    public FormulaReport {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
    }
  }
}
