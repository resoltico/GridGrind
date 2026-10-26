package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Recreates copied table structures while preserving sheet-local naming and metadata rules. */
final class ExcelSheetCopyTableSupport {
  private ExcelSheetCopyTableSupport() {}

  static void replaceTables(
      ExcelWorkbook workbook, String targetSheetName, List<TableEntryReport> tables) {
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

  static List<TableEntryReport> tablesOnSheet(XSSFSheet sourcePoiSheet) {
    List<TableEntryReport> tables = new ArrayList<>();
    for (var table : sourcePoiSheet.getTables()) {
      tables.add(ExcelTableCatalogSupport.toSnapshot(sourcePoiSheet.getSheetName(), table));
    }
    return List.copyOf(tables);
  }

  private static void copyTables(
      ExcelWorkbook workbook, String targetSheetName, List<TableEntryReport> tables) {
    for (TableEntryReport table : tables) {
      workbook.tables().setTable(copiedTableDefinition(workbook, targetSheetName, table));
    }
  }

  private static ResolvedTable copiedTableDefinition(
      ExcelWorkbook workbook, String targetSheetName, TableEntryReport table) {
    return new ResolvedTable(
        new TableName(uniqueCopiedTableName(workbook, table.name().value())),
        new SheetName(targetSheetName),
        new RangeRef(table.range().value()),
        table.structure().totalsRowCount() > 0,
        table.behavior().hasAutofilter(),
        switch (table.style()) {
          case TableStyleReport.None _ -> new TableStyleSpec.None();
          case TableStyleReport.Named named ->
              new TableStyleSpec.Named(
                  named.name(),
                  named.showFirstColumn(),
                  named.showLastColumn(),
                  named.showRowStripes(),
                  named.showColumnStripes());
        },
        table.presentation().comment().orElse(""),
        table.behavior().published(),
        table.behavior().insertRow(),
        table.behavior().insertRowShift(),
        table.presentation().headerRowCellStyle().orElse(""),
        table.presentation().dataCellStyle().orElse(""),
        table.presentation().totalsRowCellStyle().orElse(""),
        table.structure().columns().stream()
            .map(
                column ->
                    new TableColumnSpec(
                        Math.toIntExact(column.id() - 1L),
                        column.uniqueName().orElse(""),
                        column.totalsRowLabel().orElse(""),
                        column.totalsRowFunction().orElse(""),
                        column.calculatedColumnFormula().orElse("")))
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
