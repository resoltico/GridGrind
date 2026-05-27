package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Owns row editing, visibility, grouping, and snapshot operations for one XSSF sheet. */
final class ExcelRowStructureController {
  void insertRows(XSSFSheet sheet, int rowIndex, int rowCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertRowBounds(sheet, rowIndex, rowCount);
    int lastRowIndex = sheet.getLastRowNum();
    List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation> expectedValidations =
        dev.erst.gridgrind.excel.validation.ExcelDataValidationStructureSupport
            .expectedValidationsAfterInsertRows(sheet, rowIndex, rowCount);
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForInsert(sheet, rowIndex);
    if (rowIndex <= lastRowIndex) {
      sheet.shiftRows(rowIndex, lastRowIndex, rowCount, true, false);
      ExcelInsertedStructureFormattingSupport.copyAdjacentVisualFormattingIntoInsertedRows(
          sheet, rowIndex, rowCount, lastRowIndex);
    }
    dev.erst.gridgrind.excel.validation.ExcelDataValidationStructureSupport.replaceDataValidations(
        sheet, expectedValidations);
  }

  void deleteRows(XSSFSheet sheet, ExcelRowSpan rows) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    int lastRowIndex = sheet.getLastRowNum();
    if (lastRowIndex < 0) {
      throw new IllegalArgumentException("DELETE_ROWS requires at least one existing row");
    }
    if (rows.lastRowIndex() > lastRowIndex) {
      throw new IllegalArgumentException(
          "DELETE_ROWS rows must stay within existing row bounds: last existing row is "
              + ExcelIndexDisplay.rowValue(lastRowIndex)
              + "; requested "
              + ExcelIndexDisplay.describe("lastRowIndex", rows.lastRowIndex()));
    }
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForDelete(sheet, rows);
    if (rows.lastRowIndex() < lastRowIndex) {
      sheet.shiftRows(rows.lastRowIndex() + 1, lastRowIndex, -rows.count(), true, false);
    }
    int clearStart = Math.max(rows.firstRowIndex(), lastRowIndex - rows.count() + 1);
    for (int rowIndex = clearStart; rowIndex <= lastRowIndex; rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      if (row != null) {
        sheet.removeRow(row);
      }
    }
  }

  void shiftRows(XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    requireShiftedRowBounds(rows, delta);
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForShift(sheet, rows, delta);
    sheet.shiftRows(rows.firstRowIndex(), rows.lastRowIndex(), delta, true, false);
  }

  void setRowVisibility(XSSFSheet sheet, ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    ExcelRowColumnOutlineSupport.ensureRowsExist(sheet, rows);
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      sheet.getRow(rowIndex).setZeroHeight(hidden);
    }
  }

  void groupRows(XSSFSheet sheet, ExcelRowSpan rows, boolean collapsed) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    ExcelRowColumnOutlineSupport.ensureRowsExist(sheet, rows);
    sheet.groupRow(rows.firstRowIndex(), rows.lastRowIndex());
    if (collapsed) {
      ExcelRowColumnOutlineSupport.collapseRows(sheet, rows);
      return;
    }
    if (ExcelRowColumnOutlineSupport.isRowGroupCollapsed(sheet, rows)) {
      ExcelRowColumnOutlineSupport.expandRows(sheet, rows);
      return;
    }
    ExcelRowColumnOutlineSupport.clearExpandedGroupControlRow(sheet, rows);
  }

  void ungroupRows(XSSFSheet sheet, ExcelRowSpan rows) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    ExcelRowColumnOutlineSupport.ensureRowsExist(sheet, rows);
    if (ExcelRowColumnOutlineSupport.isRowGroupCollapsed(sheet, rows)) {
      ExcelRowColumnOutlineSupport.expandRows(sheet, rows);
    } else {
      ExcelRowColumnOutlineSupport.clearExpandedGroupControlRow(sheet, rows);
    }
    sheet.ungroupRow(rows.firstRowIndex(), rows.lastRowIndex());
  }

  List<WorkbookSheetResult.RowLayout> rowLayouts(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return ExcelRowColumnOutlineSupport.rowLayouts(sheet);
  }

  private void requireInsertRowBounds(XSSFSheet sheet, int rowIndex, int rowCount) {
    int lastRowIndex = sheet.getLastRowNum();
    if (rowIndex > lastRowIndex + 1) {
      throw new IllegalArgumentException(
          "INSERT_ROWS "
              + ExcelIndexDisplay.describe("rowIndex", rowIndex)
              + " must be less than or equal to last existing row + 1: "
              + ExcelIndexDisplay.rowValue(lastRowIndex + 1));
    }
    if (rowIndex + rowCount - 1 > ExcelRowSpan.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_ROWS would exceed the maximum row index: destination last row would be "
              + ExcelIndexDisplay.rowValue(rowIndex + rowCount - 1)
              + "; maximum is "
              + ExcelIndexDisplay.rowValue(ExcelRowSpan.MAX_ROW_INDEX));
    }
  }

  private void requireShiftedRowBounds(ExcelRowSpan rows, int delta) {
    if (rows.firstRowIndex() + delta < 0) {
      throw new IllegalArgumentException(
          "SHIFT_ROWS would move "
              + ExcelIndexDisplay.describe("firstRowIndex", rows.firstRowIndex())
              + " by delta "
              + delta
              + " before the first worksheet row ("
              + ExcelIndexDisplay.excelRow(0)
              + ")");
    }
    if (rows.lastRowIndex() + delta > ExcelRowSpan.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "SHIFT_ROWS would move "
              + ExcelIndexDisplay.describe("lastRowIndex", rows.lastRowIndex())
              + " by delta "
              + delta
              + " beyond the maximum row "
              + ExcelIndexDisplay.rowValue(ExcelRowSpan.MAX_ROW_INDEX));
    }
  }
}
