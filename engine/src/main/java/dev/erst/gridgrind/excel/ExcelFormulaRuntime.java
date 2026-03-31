package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/** Abstracts formula evaluation and display rendering behind a narrow GridGrind-owned seam. */
interface ExcelFormulaRuntime {
  /** Evaluates one formula cell and returns the computed value, or null when POI reports none. */
  CellValue evaluate(Cell cell);

  /** Formats one cell exactly as Excel would display it, evaluating formulas when needed. */
  String displayValue(DataFormatter formatter, Cell cell);

  /** Creates the production runtime backed by Apache POI's formula evaluator. */
  static ExcelFormulaRuntime poi(FormulaEvaluator formulaEvaluator) {
    return new PoiExcelFormulaRuntime(formulaEvaluator);
  }

  /** Production runtime that delegates evaluation and display formatting to Apache POI. */
  record PoiExcelFormulaRuntime(FormulaEvaluator formulaEvaluator) implements ExcelFormulaRuntime {
    public PoiExcelFormulaRuntime {
      Objects.requireNonNull(formulaEvaluator, "formulaEvaluator must not be null");
    }

    @Override
    public CellValue evaluate(Cell cell) {
      Objects.requireNonNull(cell, "cell must not be null");
      return formulaEvaluator.evaluate(cell);
    }

    @Override
    public String displayValue(DataFormatter formatter, Cell cell) {
      Objects.requireNonNull(formatter, "formatter must not be null");
      Objects.requireNonNull(cell, "cell must not be null");
      return formatter.formatCellValue(cell, formulaEvaluator);
    }
  }
}
