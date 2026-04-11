package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.Pxg;
import org.apache.poi.ss.formula.ptg.Pxg3D;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Rewrites explicit sheet-qualified formula tokens when a copied sheet receives a new name. */
final class ExcelFormulaSheetRenameSupport {
  private ExcelFormulaSheetRenameSupport() {}

  /** Rewrites explicit references from one sheet name to another inside one Excel formula. */
  static String renameSheet(
      XSSFWorkbook workbook,
      String formula,
      FormulaType formulaType,
      int sheetIndex,
      String oldSheetName,
      String newSheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(formulaType, "formulaType must not be null");
    Objects.requireNonNull(oldSheetName, "oldSheetName must not be null");
    Objects.requireNonNull(newSheetName, "newSheetName must not be null");
    if (formula.isBlank() || oldSheetName.equals(newSheetName)) {
      return formula;
    }

    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(workbook);
    Ptg[] ptgs =
        FormulaParser.parse(
            formula, evaluationWorkbook, parserFormulaType(formulaType), sheetIndex, -1);
    for (Ptg ptg : ptgs) {
      if (ptg instanceof Pxg pxg && pxg.getExternalWorkbookNumber() < 1) {
        if (oldSheetName.equals(pxg.getSheetName())) {
          pxg.setSheetName(newSheetName);
        }
        if (ptg instanceof Pxg3D pxg3D && oldSheetName.equals(pxg3D.getLastSheetName())) {
          pxg3D.setLastSheetName(newSheetName);
        }
      }
    }
    return FormulaRenderer.toFormulaString(evaluationWorkbook, ptgs);
  }

  private static FormulaType parserFormulaType(FormulaType formulaType) {
    return formulaType == FormulaType.CONDFORMAT ? FormulaType.CELL : formulaType;
  }
}
