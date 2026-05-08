package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Formula usage facts for one sheet. */
public record SheetFormulaSurfaceReport(
    String sheetName,
    int formulaCellCount,
    int distinctFormulaCount,
    List<FormulaPatternReport> formulas) {
  public SheetFormulaSurfaceReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    if (formulaCellCount < 0) {
      throw new IllegalArgumentException("formulaCellCount must not be negative");
    }
    if (distinctFormulaCount < 0) {
      throw new IllegalArgumentException("distinctFormulaCount must not be negative");
    }
    formulas = GridGrindResponseSupport.copyValues(formulas, "formulas");
  }
}
