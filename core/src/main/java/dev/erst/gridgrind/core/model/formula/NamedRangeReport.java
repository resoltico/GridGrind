package dev.erst.gridgrind.core.model.formula;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.DefinedNameLabel;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import java.util.Objects;

/** Structured workbook-level analysis report for one defined name. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeReport.RangeReport.class, name = "RANGE"),
  @JsonSubTypes.Type(value = NamedRangeReport.FormulaReport.class, name = "FORMULA")
})
public sealed interface NamedRangeReport
    permits NamedRangeReport.RangeReport, NamedRangeReport.FormulaReport {
  /** Defined-name identifier. */
  DefinedNameLabel name();

  /** Workbook or sheet scope of the defined name. */
  NamedRangeScope scope();

  /** Exact formula text stored in the workbook for this defined name. */
  String refersToFormula();

  /** Named range that resolves cleanly to a sheet-qualified cell or rectangular range target. */
  record RangeReport(
      DefinedNameLabel name, NamedRangeScope scope, String refersToFormula, NamedRangeTarget target)
      implements NamedRangeReport {
    public RangeReport {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      if (refersToFormula.isBlank()) {
        throw new IllegalArgumentException("refersToFormula must not be blank");
      }
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Defined name whose formula cannot be normalized to a typed range target. */
  record FormulaReport(DefinedNameLabel name, NamedRangeScope scope, String refersToFormula)
      implements NamedRangeReport {
    public FormulaReport {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      if (refersToFormula.isBlank()) {
        throw new IllegalArgumentException("refersToFormula must not be blank");
      }
    }
  }
}
