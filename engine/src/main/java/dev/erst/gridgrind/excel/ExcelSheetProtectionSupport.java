package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Reads and writes the supported sheet-protection flag subset. */
final class ExcelSheetProtectionSupport {
  private ExcelSheetProtectionSupport() {}

  /** Returns the public protection snapshot for one sheet. */
  static WorkbookSheetResult.SheetProtection snapshot(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return settings(sheet)
        .<WorkbookSheetResult.SheetProtection>map(
            WorkbookSheetResult.SheetProtection.Protected::new)
        .orElseGet(WorkbookSheetResult.SheetProtection.Unprotected::new);
  }

  /** Returns the supported protection settings when sheet protection is enabled. */
  static Optional<ExcelSheetProtectionSettings> settings(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getProtect()) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelSheetProtectionSettings(
            sheet.isAutoFilterLocked(),
            sheet.isDeleteColumnsLocked(),
            sheet.isDeleteRowsLocked(),
            sheet.isFormatCellsLocked(),
            sheet.isFormatColumnsLocked(),
            sheet.isFormatRowsLocked(),
            sheet.isInsertColumnsLocked(),
            sheet.isInsertHyperlinksLocked(),
            sheet.isInsertRowsLocked(),
            sheet.isObjectsLocked(),
            sheet.isPivotTablesLocked(),
            sheet.isScenariosLocked(),
            sheet.isSelectLockedCellsLocked(),
            sheet.isSelectUnlockedCellsLocked(),
            sheet.isSortLocked()));
  }

  /** Applies the supported protection flags after protection has been enabled on the sheet. */
  static void apply(XSSFSheet sheet, ExcelSheetProtectionSettings protection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
    sheet.lockAutoFilter(protection.autoFilterLocked());
    sheet.lockDeleteColumns(protection.deleteColumnsLocked());
    sheet.lockDeleteRows(protection.deleteRowsLocked());
    sheet.lockFormatCells(protection.formatCellsLocked());
    sheet.lockFormatColumns(protection.formatColumnsLocked());
    sheet.lockFormatRows(protection.formatRowsLocked());
    sheet.lockInsertColumns(protection.insertColumnsLocked());
    sheet.lockInsertHyperlinks(protection.insertHyperlinksLocked());
    sheet.lockInsertRows(protection.insertRowsLocked());
    sheet.lockObjects(protection.objectsLocked());
    sheet.lockPivotTables(protection.pivotTablesLocked());
    sheet.lockScenarios(protection.scenariosLocked());
    sheet.lockSelectLockedCells(protection.selectLockedCellsLocked());
    sheet.lockSelectUnlockedCells(protection.selectUnlockedCellsLocked());
    sheet.lockSort(protection.sortLocked());
  }

  /** Clears sheet protection when it is present, leaving unprotected sheets unchanged. */
  static void clear(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getProtect()) {
      return;
    }
    sheet.protectSheet(null);
  }
}
