package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Recreates copied table structures while preserving sheet-local naming and metadata rules. */
final class ExcelSheetCopyTableSupport {
  private ExcelSheetCopyTableSupport() {}

  static void replaceTables(
      ExcelWorkbook workbook, String targetSheetName, List<ExcelTableSnapshot> tables) {
    List<String> existingTableNames =
        workbook.sheet(targetSheetName).xssfSheet().getTables().stream()
            .map(table -> table.getName())
            .toList();
    for (String existingTableName : existingTableNames) {
      workbook.tables().deleteTable(existingTableName, targetSheetName);
    }
    copyTables(workbook, targetSheetName, tables);
  }

  static void requireNoTables(XSSFSheet sheet, String sheetName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
  }

  static List<ExcelTableSnapshot> tablesOnSheet(XSSFSheet sourcePoiSheet) {
    List<ExcelTableSnapshot> tables = new ArrayList<>();
    for (var table : sourcePoiSheet.getTables()) {
      tables.add(ExcelTableCatalogSupport.toSnapshot(sourcePoiSheet.getSheetName(), table));
    }
    return List.copyOf(tables);
  }

  private static void copyTables(
      ExcelWorkbook workbook, String targetSheetName, List<ExcelTableSnapshot> tables) {
    for (ExcelTableSnapshot table : tables) {
      workbook.tables().setTable(copiedTableDefinition(workbook, targetSheetName, table));
    }
  }

  private static ExcelTableDefinition copiedTableDefinition(
      ExcelWorkbook workbook, String targetSheetName, ExcelTableSnapshot table) {
    return new ExcelTableDefinition(
        uniqueCopiedTableName(workbook, table.name()),
        targetSheetName,
        table.range(),
        table.totalsRowCount() > 0,
        table.hasAutofilter(),
        switch (table.style()) {
          case ExcelTableStyleSnapshot.None _ -> new ExcelTableStyle.None();
          case ExcelTableStyleSnapshot.Named named ->
              new ExcelTableStyle.Named(
                  named.name(),
                  named.showFirstColumn(),
                  named.showLastColumn(),
                  named.showRowStripes(),
                  named.showColumnStripes());
        },
        table.comment(),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns().stream()
            .map(
                column ->
                    new ExcelTableColumnDefinition(
                        Math.toIntExact(column.id() - 1L),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
            .toList());
  }

  private static String uniqueCopiedTableName(ExcelWorkbook workbook, String baseName) {
    String candidate = baseName;
    int suffix = 2;
    while (tableNameExists(workbook, candidate)) {
      candidate = baseName + "_Copy" + suffix;
      suffix++;
    }
    return candidate;
  }

  private static boolean tableNameExists(ExcelWorkbook workbook, String candidate) {
    for (String sheetName : workbook.sheets().sheetNames()) {
      XSSFSheet sheet = ExcelWorkbookSheetSupport.requiredSheet(workbook.xssfWorkbook(), sheetName);
      for (var table : sheet.getTables()) {
        if (Objects.requireNonNullElse(table.getName(), "").equalsIgnoreCase(candidate)) {
          return true;
        }
      }
    }
    return false;
  }
}
