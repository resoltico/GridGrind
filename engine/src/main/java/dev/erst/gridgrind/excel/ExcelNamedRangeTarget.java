package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.util.CellReference;

/** Immutable workbook-core cell or rectangular-range target of a defined name. */
public record ExcelNamedRangeTarget(String sheetName, String range) {
  public ExcelNamedRangeTarget {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(range, "range must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    if (sheetName.length() > 31) {
      throw new IllegalArgumentException("sheetName must not exceed 31 characters: " + sheetName);
    }
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    range = normalizeRange(range);
  }

  /** Returns this target as a sheet-qualified absolute Excel formula string. */
  public String refersToFormula() {
    ExcelRange excelRange = ExcelRange.parse(range);
    CellReference first =
        new CellReference(sheetName, excelRange.firstRow(), excelRange.firstColumn(), true, true);
    CellReference last =
        new CellReference(sheetName, excelRange.lastRow(), excelRange.lastColumn(), true, true);
    if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
      return first.formatAsString();
    }
    return first.formatAsString() + ":" + last.formatAsString();
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
