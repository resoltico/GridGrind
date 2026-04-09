package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Owns workbook-level active, selected, visibility, protection, and summary state. */
final class ExcelSheetStateController {
  /** Returns workbook-level summary facts including active and selected sheet state. */
  WorkbookReadResult.WorkbookSummary summarizeWorkbook(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    if (workbook.sheetCount() == 0) {
      return new WorkbookReadResult.WorkbookSummary.Empty(
          0,
          List.of(),
          workbook.namedRangeCount(),
          workbook.forceFormulaRecalculationOnOpenEnabled());
    }
    return new WorkbookReadResult.WorkbookSummary.WithSheets(
        workbook.sheetCount(),
        workbook.sheetNames(),
        ExcelWorkbookSheetSupport.activeSheetName(workbook.xssfWorkbook()),
        ExcelWorkbookSheetSupport.selectedSheetNames(workbook.xssfWorkbook()),
        workbook.namedRangeCount(),
        workbook.forceFormulaRecalculationOnOpenEnabled());
  }

  /** Returns structural and state facts for one sheet. */
  WorkbookReadResult.SheetSummary summarizeSheet(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");

    XSSFSheet sheet = ExcelWorkbookSheetSupport.requiredSheet(workbook.xssfWorkbook(), sheetName);
    ExcelSheet excelSheet = workbook.sheet(sheetName);
    return new WorkbookReadResult.SheetSummary(
        sheetName,
        ExcelSheetVisibility.fromPoi(
            workbook
                .xssfWorkbook()
                .getSheetVisibility(workbook.xssfWorkbook().getSheetIndex(sheet))),
        ExcelSheetProtectionSupport.snapshot(sheet),
        excelSheet.physicalRowCount(),
        excelSheet.lastRowIndex(),
        excelSheet.lastColumnIndex());
  }

  /** Renames an existing sheet to a new destination name. */
  ExcelWorkbook renameSheet(ExcelWorkbook workbook, String sheetName, String newSheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    ExcelWorkbookSheetSupport.requireSheetName(newSheetName, "newSheetName");

    int sheetIndex =
        ExcelWorkbookSheetSupport.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    ExcelWorkbookSheetSupport.requireSheetNameAvailable(
        workbook.xssfWorkbook(), newSheetName, sheetIndex);
    if (sheetName.equals(newSheetName)) {
      return workbook;
    }
    workbook.xssfWorkbook().setSheetName(sheetIndex, newSheetName);
    return workbook;
  }

  /**
   * Deletes an existing sheet while preserving a valid visible active and selected workbook state.
   */
  ExcelWorkbook deleteSheet(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");

    int sheetIndex =
        ExcelWorkbookSheetSupport.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    if (workbook.xssfWorkbook().getNumberOfSheets() == 1) {
      throw new IllegalArgumentException(
          "cannot delete sheet '" + sheetName + "': a workbook must contain at least one sheet");
    }
    ExcelWorkbookSheetSupport.requireNotLastVisibleSheet(
        workbook.xssfWorkbook(),
        sheetIndex,
        "cannot delete the last visible sheet '" + sheetName + "'");
    workbook.xssfWorkbook().removeSheetAt(sheetIndex);
    ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return workbook;
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  ExcelWorkbook moveSheet(ExcelWorkbook workbook, String sheetName, int targetIndex) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");

    ExcelWorkbookSheetSupport.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    ExcelWorkbookSheetSupport.requireTargetIndex(workbook.xssfWorkbook(), targetIndex);
    workbook.xssfWorkbook().setSheetOrder(sheetName, targetIndex);
    ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return workbook;
  }

  /** Sets the active sheet and ensures it is also selected. */
  ExcelWorkbook setActiveSheet(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");

    int sheetIndex =
        ExcelWorkbookSheetSupport.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    ExcelWorkbookSheetSupport.requireVisibleSheet(
        workbook.xssfWorkbook(), sheetIndex, "cannot activate hidden sheet '" + sheetName + "'");

    workbook.xssfWorkbook().setActiveSheet(sheetIndex);
    workbook.xssfWorkbook().getSheetAt(sheetIndex).setSelected(true);
    ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return workbook;
  }

  /**
   * Sets the selected visible sheet set and normalizes the active tab into that selection if
   * needed.
   */
  ExcelWorkbook setSelectedSheets(ExcelWorkbook workbook, List<String> sheetNames) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    List<String> normalizedSheetNames =
        ExcelWorkbookSheetSupport.copySelectedSheetNames(sheetNames);

    Set<String> selectedNames = new LinkedHashSet<>(normalizedSheetNames);
    for (String sheetName : selectedNames) {
      int sheetIndex =
          ExcelWorkbookSheetSupport.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
      ExcelWorkbookSheetSupport.requireVisibleSheet(
          workbook.xssfWorkbook(), sheetIndex, "cannot select hidden sheet '" + sheetName + "'");
    }

    for (int sheetIndex = 0;
        sheetIndex < workbook.xssfWorkbook().getNumberOfSheets();
        sheetIndex++) {
      XSSFSheet sheet = workbook.xssfWorkbook().getSheetAt(sheetIndex);
      sheet.setSelected(selectedNames.contains(sheet.getSheetName()));
    }

    int activeSheetIndex = workbook.xssfWorkbook().getActiveSheetIndex();
    if (activeSheetIndex < 0
        || activeSheetIndex >= workbook.xssfWorkbook().getNumberOfSheets()
        || !selectedNames.contains(workbook.xssfWorkbook().getSheetName(activeSheetIndex))) {
      workbook
          .xssfWorkbook()
          .setActiveSheet(
              ExcelWorkbookSheetSupport.requiredSelectedSheetIndex(
                  workbook.xssfWorkbook(), normalizedSheetNames.getFirst()));
    }

    ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return workbook;
  }

  /** Sets one sheet visibility while preserving at least one visible active selected sheet. */
  ExcelWorkbook setSheetVisibility(
      ExcelWorkbook workbook, String sheetName, ExcelSheetVisibility visibility) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    Objects.requireNonNull(visibility, "visibility must not be null");

    int sheetIndex =
        ExcelWorkbookSheetSupport.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    ExcelSheetVisibility currentVisibility =
        ExcelSheetVisibility.fromPoi(workbook.xssfWorkbook().getSheetVisibility(sheetIndex));
    if (visibility != ExcelSheetVisibility.VISIBLE
        && currentVisibility == ExcelSheetVisibility.VISIBLE) {
      ExcelWorkbookSheetSupport.requireNotLastVisibleSheet(
          workbook.xssfWorkbook(),
          sheetIndex,
          "cannot hide the last visible sheet '" + sheetName + "'");
    }
    workbook.xssfWorkbook().setSheetVisibility(sheetIndex, visibility.toPoi());
    ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return workbook;
  }

  /** Enables sheet protection with the exact supported lock flags. */
  ExcelWorkbook setSheetProtection(
      ExcelWorkbook workbook, String sheetName, ExcelSheetProtectionSettings protection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    Objects.requireNonNull(protection, "protection must not be null");

    XSSFSheet sheet = ExcelWorkbookSheetSupport.requiredSheet(workbook.xssfWorkbook(), sheetName);
    sheet.protectSheet("");
    ExcelSheetProtectionSupport.apply(sheet, protection);
    return workbook;
  }

  /** Disables sheet protection entirely. */
  ExcelWorkbook clearSheetProtection(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");

    XSSFSheet sheet = ExcelWorkbookSheetSupport.requiredSheet(workbook.xssfWorkbook(), sheetName);
    ExcelSheetProtectionSupport.clear(sheet);
    return workbook;
  }
}
