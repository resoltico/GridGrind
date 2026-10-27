package dev.erst.gridgrind.engine.protection;

import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.engine.backend.xssf.SheetCollectionAdapter;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookProtection;

/** Owns workbook/sheet lock flags, passwords and factual protection inspection. */
@SuppressWarnings("PMD.CommentRequired")
public final class ProtectionAdapter {
  private final XSSFWorkbook workbook;

  public ProtectionAdapter(XSSFWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  public void setSheetProtection(String sheetName, SheetProtectionSettings protection) {
    setSheetProtection(sheetName, protection, Optional.empty());
  }

  /** Returns workbook-level protection facts including password-hash presence. */
  public WorkbookProtectionReport snapshotWorkbook() {

    Optional<CTWorkbookProtection> protection =
        workbook.getCTWorkbook().isSetWorkbookProtection()
            ? Optional.of(workbook.getCTWorkbook().getWorkbookProtection())
            : Optional.empty();
    return new WorkbookProtectionReport(
        workbook.isStructureLocked(),
        workbook.isWindowsLocked(),
        workbook.isRevisionLocked(),
        workbookPasswordHashPresent(protection),
        revisionsPasswordHashPresent(protection));
  }

  /** Enables sheet protection with the exact supported lock flags. */
  public void setSheetProtection(
      String sheetName, SheetProtectionSettings protection, Optional<String> password) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");
    Objects.requireNonNull(protection, "protection must not be null");
    Objects.requireNonNull(password, "password must not be null");

    XSSFSheet sheet = SheetCollectionAdapter.requiredSheet(workbook, sheetName);
    sheet.protectSheet(password.orElse(""));
    if (password.isPresent()) {
      sheet.setSheetPassword(password.orElseThrow(), HashAlgorithm.sha512);
    }
    apply(sheet, protection);
  }

  /** Disables sheet protection entirely. */
  public void clearSheetProtection(String sheetName) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");

    XSSFSheet sheet = SheetCollectionAdapter.requiredSheet(workbook, sheetName);
    clear(sheet);
  }

  /** Applies workbook-level protection and password hashes authoritatively. */
  public void setWorkbookProtection(ResolvedWorkbookProtection protection) {
    Objects.requireNonNull(protection, "protection must not be null");

    workbook.unLock();
    workbook.getCTWorkbook().addNewWorkbookProtection();
    if (protection.structureLocked()) {
      workbook.lockStructure();
    }
    if (protection.windowsLocked()) {
      workbook.lockWindows();
    }
    if (protection.revisionsLocked()) {
      workbook.lockRevision();
    }
    if (protection.workbookPassword().isPresent()) {
      workbook.setWorkbookPassword(
          protection.workbookPassword().orElseThrow(), HashAlgorithm.sha512);
    }
    if (protection.revisionsPassword().isPresent()) {
      workbook.setRevisionsPassword(
          protection.revisionsPassword().orElseThrow(), HashAlgorithm.sha512);
    }
    normalizeWorkbookProtectionNode();
  }

  /** Clears workbook-level protection and password hashes entirely. */
  public void clearWorkbookProtection() {
    workbook.unLock();
  }

  private void normalizeWorkbookProtectionNode() {
    Optional<CTWorkbookProtection> protection =
        Optional.ofNullable(workbook.getCTWorkbook().getWorkbookProtection());
    boolean hasLocks =
        workbook.isStructureLocked() || workbook.isWindowsLocked() || workbook.isRevisionLocked();
    boolean hasPasswords =
        workbookPasswordHashPresent(protection) || revisionsPasswordHashPresent(protection);
    if (!hasLocks && !hasPasswords) {
      workbook.getCTWorkbook().unsetWorkbookProtection();
    }
  }

  private static boolean workbookPasswordHashPresent(Optional<CTWorkbookProtection> protection) {
    return protection.isPresent()
        && (protection.orElseThrow().isSetWorkbookPassword()
            || protection.orElseThrow().isSetWorkbookHashValue()
            || protection.orElseThrow().isSetWorkbookSaltValue()
            || protection.orElseThrow().isSetWorkbookSpinCount()
            || protection.orElseThrow().isSetWorkbookAlgorithmName());
  }

  private static boolean revisionsPasswordHashPresent(Optional<CTWorkbookProtection> protection) {
    return protection.isPresent()
        && (protection.orElseThrow().isSetRevisionsPassword()
            || protection.orElseThrow().isSetRevisionsHashValue()
            || protection.orElseThrow().isSetRevisionsSaltValue()
            || protection.orElseThrow().isSetRevisionsSpinCount()
            || protection.orElseThrow().isSetRevisionsAlgorithmName());
  }

  /** Returns the public protection snapshot for one sheet. */
  public static SheetProtectionReport snapshotSheet(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return settings(sheet)
        .<SheetProtectionReport>map(SheetProtectionReport.Protected::new)
        .orElseGet(SheetProtectionReport.Unprotected::new);
  }

  /** Returns the supported protection settings when sheet protection is enabled. */
  public static Optional<SheetProtectionSettings> settings(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getProtect()) {
      return Optional.empty();
    }
    return Optional.of(
        new SheetProtectionSettings(
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
  public static void apply(XSSFSheet sheet, SheetProtectionSettings protection) {
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
  public static void clear(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getProtect()) {
      return;
    }
    sheet.protectSheet(null);
  }
}
