package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;
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
      Optional<ExcelRange> headerRange = headerRange(table);
      if (headerRange.isPresent()
          && ExcelSheetStructureSupport.intersects(headerRange.orElseThrow(), changedRange)) {
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

    if (headerRange(table).isPresent()) {
      table.updateHeaders();
    }
  }

  private static Optional<ExcelRange> headerRange(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    if (table.getHeaderRowCount() < 1) {
      return Optional.empty();
    }
    Optional<ExcelRange> tableRange =
        ExcelSheetStructureSupport.parseOptionalRange(
            Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
    if (tableRange.isEmpty()) {
      return Optional.empty();
    }
    ExcelRange resolved = tableRange.orElseThrow();
    return Optional.of(
        new ExcelRange(
            resolved.firstRow(),
            Math.min(resolved.lastRow(), resolved.firstRow() + table.getHeaderRowCount() - 1),
            resolved.firstColumn(),
            resolved.lastColumn()));
  }
}
