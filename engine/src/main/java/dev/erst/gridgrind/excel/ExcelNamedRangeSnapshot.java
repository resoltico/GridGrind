package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable analyzed defined-name snapshot captured from an `.xlsx` workbook. */
public sealed interface ExcelNamedRangeSnapshot
    permits ExcelNamedRangeSnapshot.RangeSnapshot, ExcelNamedRangeSnapshot.FormulaSnapshot {

  /** Returns the defined-name identifier. */
  String name();

  /** Returns the workbook or sheet scope of the defined name. */
  ExcelNamedRangeScope scope();

  /** Returns the exact underlying Excel formula text stored by the workbook. */
  String refersToFormula();

  /** Snapshot of a defined name that resolves cleanly to a typed sheet-and-range target. */
  record RangeSnapshot(
      String name, ExcelNamedRangeScope scope, String refersToFormula, ExcelNamedRangeTarget target)
      implements ExcelNamedRangeSnapshot {
    public RangeSnapshot {
      name = ExcelNamedRangeDefinition.validateName(name);
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Snapshot of a defined name whose formula cannot be normalized to a typed range target. */
  record FormulaSnapshot(String name, ExcelNamedRangeScope scope, String refersToFormula)
      implements ExcelNamedRangeSnapshot {
    public FormulaSnapshot {
      name = ExcelNamedRangeDefinition.validateName(name);
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
    }
  }
}
