package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Sheet lifecycle, selection, visibility, and protection operations for one workbook. */
public final class ExcelWorkbookSheets {
  private final ExcelWorkbook workbook;

  ExcelWorkbookSheets(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Renames an existing sheet to a new destination name. */
  public ExcelWorkbookSheets renameSheet(String sheetName, String newSheetName) {
    new ExcelSheetStateController().renameSheet(workbook, sheetName, newSheetName);
    return this;
  }

  /** Deletes an existing sheet from the workbook. A workbook must retain at least one sheet. */
  public ExcelWorkbookSheets deleteSheet(String sheetName) {
    new ExcelSheetStateController().deleteSheet(workbook, sheetName);
    return this;
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  public ExcelWorkbookSheets moveSheet(String sheetName, int targetIndex) {
    new ExcelSheetStateController().moveSheet(workbook, sheetName, targetIndex);
    return this;
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  public ExcelWorkbookSheets copySheet(
      String sourceSheetName, String newSheetName, ExcelSheetCopyPosition position) {
    new ExcelSheetCopyController().copySheet(workbook, sourceSheetName, newSheetName, position);
    return this;
  }

  /** Sets the active sheet and ensures it is selected. */
  public ExcelWorkbookSheets setActiveSheet(String sheetName) {
    new ExcelSheetStateController().setActiveSheet(workbook, sheetName);
    return this;
  }

  /** Sets the selected visible sheet set and normalizes the active tab into that selection. */
  public ExcelWorkbookSheets setSelectedSheets(List<String> sheetNames) {
    new ExcelSheetStateController().setSelectedSheets(workbook, sheetNames);
    return this;
  }

  /** Sets one sheet visibility while preserving a visible active selected sheet. */
  public ExcelWorkbookSheets setSheetVisibility(String sheetName, ExcelSheetVisibility visibility) {
    new ExcelSheetStateController().setSheetVisibility(workbook, sheetName, visibility);
    return this;
  }

  /** Enables sheet protection with the exact supported lock flags. */
  public ExcelWorkbookSheets setSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection) {
    return setSheetProtection(sheetName, protection, Optional.empty());
  }

  /** Enables sheet protection with the exact supported lock flags and optional password. */
  public ExcelWorkbookSheets setSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection, Optional<String> password) {
    new ExcelSheetStateController().setSheetProtection(workbook, sheetName, protection, password);
    return this;
  }

  /** Disables sheet protection entirely. */
  public ExcelWorkbookSheets clearSheetProtection(String sheetName) {
    new ExcelSheetStateController().clearSheetProtection(workbook, sheetName);
    return this;
  }

  /** Returns the number of sheets in the workbook. */
  public int sheetCount() {
    return workbook.xssfWorkbook().getNumberOfSheets();
  }

  /** Returns an ordered list of all sheet names in the workbook. */
  public List<String> sheetNames() {
    return ExcelWorkbookSheetAccessSupport.sheetNames(workbook);
  }
}
