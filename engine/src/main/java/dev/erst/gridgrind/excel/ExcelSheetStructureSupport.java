package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Shared sheet-structure helpers for range-backed Excel objects such as autofilters and tables. */
final class ExcelSheetStructureSupport {
  private ExcelSheetStructureSupport() {}

  /** Returns whether the supplied range's first row lacks any nonblank header cell. */
  static boolean headerRowMissing(XSSFSheet sheet, ExcelRange range) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");

    Row headerRow = sheet.getRow(range.firstRow());
    if (headerRow == null) {
      return true;
    }
    for (int columnIndex = range.firstColumn(); columnIndex <= range.lastColumn(); columnIndex++) {
      if (hasHeaderValue(headerRow.getCell(columnIndex))) {
        return false;
      }
    }
    return true;
  }

  /** Parses one A1-style range string and returns null when it cannot be parsed. */
  static ExcelRange parseRangeOrNull(String rawRange) {
    Objects.requireNonNull(rawRange, "rawRange must not be null");
    try {
      return ExcelRange.parse(rawRange);
    } catch (RuntimeException exception) {
      return null;
    }
  }

  /** Returns whether two rectangular ranges intersect at any cell. */
  static boolean intersects(ExcelRange first, ExcelRange second) {
    Objects.requireNonNull(first, "first must not be null");
    Objects.requireNonNull(second, "second must not be null");
    boolean rowsOverlap =
        Math.max(first.firstRow(), second.firstRow())
            <= Math.min(first.lastRow(), second.lastRow());
    boolean columnsOverlap =
        Math.max(first.firstColumn(), second.firstColumn())
            <= Math.min(first.lastColumn(), second.lastColumn());
    return rowsOverlap && columnsOverlap;
  }

  /** Formats one parsed range as a canonical A1-style string. */
  static String formatRange(ExcelRange range) {
    Objects.requireNonNull(range, "range must not be null");
    return toCellRangeAddress(range).formatAsString();
  }

  /** Converts one parsed range into the POI cell-range address type. */
  static CellRangeAddress toCellRangeAddress(ExcelRange range) {
    Objects.requireNonNull(range, "range must not be null");
    return new CellRangeAddress(
        range.firstRow(), range.lastRow(), range.firstColumn(), range.lastColumn());
  }

  /** Returns whether one candidate header cell contributes a nonblank header value. */
  static boolean hasHeaderValue(Cell cell) {
    if (cell == null) {
      return false;
    }
    return switch (cell.getCellType()) {
      case BLANK -> false;
      case _NONE -> false;
      case STRING -> !cell.getStringCellValue().isBlank();
      case FORMULA -> !cell.getCellFormula().isBlank();
      default -> true;
    };
  }
}
