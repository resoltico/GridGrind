package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;
import org.apache.poi.ss.util.CellReference;

/** Immutable workbook-core target of a defined name. */
public sealed interface ExcelNamedRangeTarget
    permits ExcelNamedRangeTarget.Range, ExcelNamedRangeTarget.Formula {

  /** Sheet-local cell or rectangular range target. */
  record Range(String sheetName, String range) implements ExcelNamedRangeTarget {
    public Range {
      ExcelSheetNames.requireValid(sheetName, "sheetName");
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
      range = normalizeRange(range);
    }

    @Override
    public String refersToFormula() {
      ExcelRange excelRange = ExcelRange.parse(range);
      CellReference first =
          new CellReference(sheetName, excelRange.firstRow(), excelRange.firstColumn(), true, true);
      if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
        return first.formatAsString();
      }
      return first.formatAsString()
          + ":"
          + new CellReference(excelRange.lastRow(), excelRange.lastColumn(), true, true)
              .formatAsString();
    }
  }

  /** Formula-defined target stored exactly as authored. */
  record Formula(String formula) implements ExcelNamedRangeTarget {
    public Formula {
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
    }

    @Override
    public String refersToFormula() {
      return formula;
    }
  }

  /** Returns this target as a sheet-qualified absolute Excel formula string. */
  String refersToFormula();

  /** Creates a typed sheet-local range target. */
  static ExcelNamedRangeTarget range(String sheetName, String range) {
    return new Range(sheetName, range);
  }

  /** Creates an exact formula-defined target. */
  static ExcelNamedRangeTarget formula(String formula) {
    return new Formula(formula);
  }

  private static String normalizeRange(String range) {
    ExcelRange excelRange = ExcelRange.parse(range);
    CellReference first =
        new CellReference(excelRange.firstRow(), excelRange.firstColumn(), false, false);
    if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
      return first.formatAsString();
    }
    CellReference last =
        new CellReference(excelRange.lastRow(), excelRange.lastColumn(), false, false);
    return first.formatAsString() + ":" + last.formatAsString();
  }
}
