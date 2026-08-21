package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.contract.dto.FormulaTextValidation;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;

/** Centralizes authored, rewritten, and scratch formula writes behind consistent errors. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelFormulaWriteSupport {
  private ExcelFormulaWriteSupport() {}

  public static void setAuthoredFormula(
      Cell cell,
      String formula,
      ExcelFormulaRuntime formulaRuntime,
      String sheetName,
      String address) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    try {
      ExcelFormulaLimits.requireSupportedFormula(
          cellContext(cell), formula); // LIM-013, LIM-014, LIM-015
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
      ExcelFormulaLimits.requireSupportedFormula(
          cellContext(cell), formula); // LIM-013, LIM-014, LIM-015
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

  static void setRewrittenFormula(Cell cell, String formula, String operation) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(operation, "operation must not be null");
    try {
      ExcelFormulaLimits.requireSupportedFormula(
          cellContext(cell), formula); // LIM-013, LIM-014, LIM-015
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

  static void setScratchFormula(Cell cell, String formula, String operation) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(operation, "operation must not be null");
    try {
      ExcelFormulaLimits.requireSupportedFormula(
          cellContext(cell), formula); // LIM-013, LIM-014, LIM-015
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "Invalid scratch formula for " + operation + ": " + formula, exception);
    }
  }

  private static ExcelFormulaLimits.CellContext cellContext(Cell cell) {
    return new ExcelFormulaLimits.CellContext(
        cell.getSheet().getWorkbook(),
        cell.getSheet().getWorkbook().getSheetIndex(cell.getSheet()));
  }
}
