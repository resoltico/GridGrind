package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;

/** Centralizes authored, rewritten, and scratch formula writes behind consistent errors. */
final class ExcelFormulaWriteSupport {
  private ExcelFormulaWriteSupport() {}

  static void setAuthoredFormula(
      Cell cell,
      String formula,
      ExcelFormulaRuntime formulaRuntime,
      String sheetName,
      String address) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    try {
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(formulaRuntime, sheetName, address, formula, exception);
    }
  }

  static void setAuthoredFormula(Cell cell, String formula) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    String sheetName = cell.getSheet().getSheetName();
    String address = cell.getAddress().formatAsString();
    try {
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(sheetName, address, formula, exception);
    }
  }

  static void setRewrittenFormula(Cell cell, String formula, String operation) {
    Objects.requireNonNull(cell, "cell must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(operation, "operation must not be null");
    try {
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
      cell.setCellFormula(formula);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "Invalid scratch formula for " + operation + ": " + formula, exception);
    }
  }
}
