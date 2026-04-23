package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;
import org.apache.poi.ss.util.CellReference;

/** Immutable workbook-core cell or rectangular-range target of a defined name. */
public record ExcelNamedRangeTarget(String sheetName, String range, String formula) {
  /** Creates a sheet-local cell or rectangular range target. */
  public ExcelNamedRangeTarget(String sheetName, String range) {
    this(sheetName, range, null);
  }

  /** Creates a formula-defined target that is stored exactly as authored. */
  public ExcelNamedRangeTarget(String formula) {
    this(null, null, formula);
  }

  public ExcelNamedRangeTarget {
    if (formula != null) {
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      if (sheetName != null || range != null) {
        throw new IllegalArgumentException(
            "formula-defined named-range targets must not also set sheetName or range");
      }
    } else {
      ExcelSheetNames.requireValid(sheetName, "sheetName");
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
      range = normalizeRange(range);
    }
  }

  /** Returns this target as a sheet-qualified absolute Excel formula string. */
  public String refersToFormula() {
    if (formula != null) {
      return formula;
    }
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
