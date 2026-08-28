package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Reads authored formula facts without invoking a formula evaluator. */
final class ExcelFormulaFactReader {
  private ExcelFormulaFactReader() {}

  static List<ExcelFormulaCell> formulaCells(Sheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    List<ExcelFormulaCell> formulas = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          formulas.add(
              new ExcelFormulaCell(
                  new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                  cell.getCellFormula()));
        }
      }
    }
    return List.copyOf(formulas);
  }
}
