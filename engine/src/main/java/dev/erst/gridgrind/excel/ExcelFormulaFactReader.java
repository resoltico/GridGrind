package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.FormulaPatternReport;
import dev.erst.gridgrind.core.model.formula.SheetFormulaSurfaceReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Groups stored formula text without invoking a formula evaluator. */
final class ExcelFormulaFactReader {
  private ExcelFormulaFactReader() {}

  // Request-local insertion order preserves physical formula/address order; report values are
  // allocated per cell/pattern.
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  static SheetFormulaSurfaceReport surface(Sheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Map<String, List<CellAddress>> formulas = new LinkedHashMap<>();
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          formulas
              .computeIfAbsent(cell.getCellFormula(), _ -> new ArrayList<>())
              .add(
                  new CellAddress(
                      new CellReference(cell.getRowIndex(), cell.getColumnIndex())
                          .formatAsString()));
          count++;
        }
      }
    }
    List<FormulaPatternReport> patterns = new ArrayList<>(formulas.size());
    for (Map.Entry<String, List<CellAddress>> entry : formulas.entrySet()) {
      patterns.add(
          new FormulaPatternReport(entry.getKey(), entry.getValue().size(), entry.getValue()));
    }
    return new SheetFormulaSurfaceReport(
        new SheetLabel(sheet.getSheetName()), count, patterns.size(), patterns);
  }
}
