package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.List;
import java.util.Objects;

/** Formula usage facts for one sheet. */
public record SheetFormulaSurfaceReport(
    SheetLabel sheetName,
    int formulaCellCount,
    int distinctFormulaCount,
    List<FormulaPatternReport> formulas) {
  public SheetFormulaSurfaceReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (formulaCellCount < 0) {
      throw new IllegalArgumentException("formulaCellCount must not be negative");
    }
    if (distinctFormulaCount < 0) {
      throw new IllegalArgumentException("distinctFormulaCount must not be negative");
    }
    formulas = WorkbookResultSupport.copyValues(formulas, "formulas");
  }
}
