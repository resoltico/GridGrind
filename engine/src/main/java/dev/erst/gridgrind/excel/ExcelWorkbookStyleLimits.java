package dev.erst.gridgrind.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Centralized workbook style-capacity checks for POI-authored cell styles. */
final class ExcelWorkbookStyleLimits {
  static final int MAX_CELL_STYLES = SpreadsheetVersion.EXCEL2007.getMaxCellStyles(); // LIM-011

  private ExcelWorkbookStyleLimits() {}

  static void requireCellStyleCapacity(XSSFWorkbook workbook) {
    requireCellStyleCapacity(workbook.getNumCellStyles());
  }

  static void requireCellStyleCapacity(int existingStyleCount) {
    if (existingStyleCount >= MAX_CELL_STYLES) {
      throw new IllegalArgumentException(
          "workbook cannot create more than "
              + MAX_CELL_STYLES
              + " cell styles (Excel/POI workbook style limit)");
    }
  }
}
