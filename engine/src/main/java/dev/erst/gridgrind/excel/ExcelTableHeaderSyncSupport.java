package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Keeps persisted table header metadata converged with live sheet-cell header values. */
final class ExcelTableHeaderSyncSupport {
  private ExcelTableHeaderSyncSupport() {}

  /** Synchronizes table header metadata for every table present in the workbook. */
  static void syncAllHeaders(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      syncAllHeaders(workbook.getSheetAt(sheetIndex));
    }
  }

  /** Synchronizes table header metadata for sheet tables intersecting the changed cell range. */
  static void syncAffectedHeaders(XSSFSheet sheet, ExcelRange changedRange) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(changedRange, "changedRange must not be null");

    for (XSSFTable table : sheet.getTables()) {
      ExcelRange headerRange = headerRangeOrNull(table);
      if (headerRange != null && ExcelSheetStructureSupport.intersects(headerRange, changedRange)) {
        syncHeader(table);
      }
    }
  }

  private static void syncAllHeaders(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");

    for (XSSFTable table : sheet.getTables()) {
      syncHeader(table);
    }
  }

  private static void syncHeader(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    if (headerRangeOrNull(table) != null) {
      table.updateHeaders();
    }
  }

  private static ExcelRange headerRangeOrNull(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    if (table.getHeaderRowCount() < 1) {
      return null;
    }
    ExcelRange tableRange =
        ExcelSheetStructureSupport.parseRangeOrNull(
            Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
    if (tableRange == null) {
      return null;
    }
    return new ExcelRange(
        tableRange.firstRow(),
        Math.min(tableRange.lastRow(), tableRange.firstRow() + table.getHeaderRowCount() - 1),
        tableRange.firstColumn(),
        tableRange.lastColumn());
  }
}
