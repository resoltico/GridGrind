package dev.erst.gridgrind.engine.poi;

import dev.erst.gridgrind.core.model.formula.FormulaTextValidation;
import dev.erst.gridgrind.engine.formulas.FormulaExceptions;
import dev.erst.gridgrind.engine.formulas.FormulaLimits;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;

/** Centralizes authored, rewritten, and scratch formula writes behind consistent errors. */
@SuppressWarnings("PMD.CommentRequired")
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_FORMULA_LENGTH,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_NESTED_FUNCTION_LEVELS,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_FUNCTION_ARGUMENTS
})
public final class FormulaWrites {
  private FormulaWrites() {}

  public static void setAuthoredFormula(
      Cell cell, String formula, FormulaRuntime formulaRuntime, String sheetName, String address) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    try {
      FormulaLimits.requireSupportedFormula(cellContext(cell), formula);
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(formulaRuntime, sheetName, address, formula, exception);
    }
  }

  public static void setAuthoredFormula(Cell cell, String formula) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    String sheetName = cell.getSheet().getSheetName();
    String address = cell.getAddress().formatAsString();
    try {
      FormulaLimits.requireSupportedFormula(cellContext(cell), formula);
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(sheetName, address, formula, exception);
    }
  }

  /** Writes one opaque formula body directly to OOXML without POI formula parsing. */
  public static void setOpaqueFormula(Cell cell, String formula) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    FormulaTextValidation.requireRawFormulaBody(formula, "formula");
    switch (cell) {
      case XSSFCell xssfCell -> setOpaqueXssfFormula(xssfCell, formula);
      case SXSSFCell sxssfCell -> sxssfCell.setCellFormula(formula);
      default ->
          throw new IllegalArgumentException("Opaque formulas require an XSSF or SXSSF cell");
    }
  }

  private static void setOpaqueXssfFormula(XSSFCell xssfCell, String formula) {
    boolean validationEnabled = xssfCell.getSheet().getWorkbook().getCellFormulaValidation();
    xssfCell.getSheet().getWorkbook().setCellFormulaValidation(false);
    try {
      xssfCell.setCellFormula(formula);
    } finally {
      xssfCell.getSheet().getWorkbook().setCellFormulaValidation(validationEnabled);
    }
    var ctCell = xssfCell.getCTCell();
    if (ctCell.isSetV()) {
      ctCell.unsetV();
    }
    if (ctCell.isSetT()) {
      ctCell.unsetT();
    }
  }

  public static void setRewrittenFormula(Cell cell, String formula, String operation) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(operation, "operation must not be null");
    try {
      FormulaLimits.requireSupportedFormula(cellContext(cell), formula);
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw new IllegalStateException(
          operation
              + " produced an invalid formula at "
              + cell.getSheet().getSheetName()
              + "!"
              + cell.getAddress().formatAsString()
              + ": "
              + formula,
          exception);
    }
  }

  public static void setScratchFormula(Cell cell, String formula, String operation) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(operation, "operation must not be null");
    try {
      FormulaLimits.requireSupportedFormula(cellContext(cell), formula);
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "Invalid scratch formula for " + operation + ": " + formula, exception);
    }
  }

  private static FormulaLimits.CellContext cellContext(Cell cell) {
    return new FormulaLimits.CellContext(
        cell.getSheet().getWorkbook(),
        cell.getSheet().getWorkbook().getSheetIndex(cell.getSheet()));
  }
}
