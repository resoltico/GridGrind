package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationStructureSupport;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;

/** Owns column editing, visibility, grouping, and snapshot operations for one XSSF sheet. */
final class ExcelColumnStructureController {
  void insertColumns(XSSFSheet sheet, int columnIndex, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertColumnBounds(sheet, columnIndex, columnCount);
    ExcelRowColumnStructureGuardSupport.rejectFormulaBearingWorkbookForColumnEdit(
        sheet.getWorkbook(), "INSERT_COLUMNS");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    int lastColumnIndex = lastColumnIndex(sheet);
    Map<Integer, CTCol> explicitColumns =
        ExcelRowColumnOutlineSupport.snapshotColumnDefinitions(sheet);
    List<CTDataValidation> expectedValidations =
        ExcelDataValidationStructureSupport.expectedValidationsAfterInsertColumns(
            sheet, columnIndex, columnCount);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments =
          commentRepairSupport.expectedCommentsAfterInsertColumns(columnIndex, columnCount);
    }
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForInsert(sheet, columnIndex);
    Map<Integer, CTCol> shiftedExplicitColumns =
        ExcelRowColumnOutlineSupport.mutableColumnDefinitionsCopy(
            ExcelRowColumnOutlineSupport.shiftedForInsert(
                explicitColumns, columnIndex, columnCount));
    if (columnIndex <= lastColumnIndex) {
      sheet.shiftColumns(columnIndex, lastColumnIndex, columnCount);
      ExcelInsertedStructureFormattingSupport.copyAdjacentVisualFormattingIntoInsertedColumns(
          sheet, shiftedExplicitColumns, columnIndex, columnCount, lastColumnIndex);
    }
    ExcelColumnDefinitionSupport.rebuildColumnDefinitions(sheet, shiftedExplicitColumns);
    ExcelDataValidationStructureSupport.replaceDataValidations(sheet, expectedValidations);
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  void deleteColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      throw new IllegalArgumentException("DELETE_COLUMNS requires at least one existing column");
    }
    if (columns.lastColumnIndex() > lastColumnIndex) {
      throw new IllegalArgumentException(
          "DELETE_COLUMNS columns must stay within existing column bounds: last existing column is "
              + ExcelIndexDisplay.columnValue(lastColumnIndex)
              + "; requested "
              + ExcelIndexDisplay.describe("lastColumnIndex", columns.lastColumnIndex()));
    }
    ExcelRowColumnStructureGuardSupport.rejectFormulaBearingWorkbookForColumnEdit(
        sheet.getWorkbook(), "DELETE_COLUMNS");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns =
        ExcelRowColumnOutlineSupport.snapshotColumnDefinitions(sheet);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterDeleteColumns(columns);
    }
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForDelete(sheet, columns);
    if (columns.lastColumnIndex() < lastColumnIndex) {
      sheet.shiftColumns(columns.lastColumnIndex() + 1, lastColumnIndex, -columns.count());
    }
    int clearStart = Math.max(columns.firstColumnIndex(), lastColumnIndex - columns.count() + 1);
    ExcelRowColumnOutlineSupport.clearTrailingCells(sheet, clearStart, lastColumnIndex);
    ExcelColumnDefinitionSupport.rebuildColumnDefinitions(
        sheet, ExcelRowColumnOutlineSupport.shiftedForDelete(explicitColumns, columns));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  void shiftColumns(XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    requireShiftedColumnBounds(columns, delta);
    ExcelRowColumnStructureGuardSupport.rejectFormulaBearingWorkbookForColumnEdit(
        sheet.getWorkbook(), "SHIFT_COLUMNS");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns =
        ExcelRowColumnOutlineSupport.snapshotColumnDefinitions(sheet);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterShiftColumns(columns, delta);
    }
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForShift(
        sheet, columns, delta);
    sheet.shiftColumns(columns.firstColumnIndex(), columns.lastColumnIndex(), delta);
    ExcelColumnDefinitionSupport.rebuildColumnDefinitions(
        sheet, ExcelRowColumnOutlineSupport.shiftedForShift(explicitColumns, columns, delta));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  void setColumnVisibility(XSSFSheet sheet, ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    for (int columnIndex = columns.firstColumnIndex();
        columnIndex <= columns.lastColumnIndex();
        columnIndex++) {
      sheet.setColumnHidden(columnIndex, hidden);
    }
    ExcelRowColumnStructureController.canonicalizeColumnDefinitions(sheet);
  }

  void groupColumns(XSSFSheet sheet, ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    sheet.groupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    if (collapsed) {
      ExcelRowColumnOutlineSupport.collapseColumns(sheet, columns);
      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(sheet);
      return;
    }
    ExcelRowColumnOutlineSupport.clearExpandedGroupControlColumn(sheet, columns);
    ExcelRowColumnStructureController.canonicalizeColumnDefinitions(sheet);
  }

  void ungroupColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    ExcelRowColumnOutlineSupport.prepareColumnsForUngroup(sheet, columns);
    sheet.ungroupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    ExcelRowColumnStructureController.canonicalizeColumnDefinitions(sheet);
  }

  int lastColumnIndex(XSSFSheet sheet) {
    return ExcelRowColumnOutlineSupport.lastColumnIndex(sheet);
  }

  List<WorkbookSheetResult.ColumnLayout> columnLayouts(XSSFSheet sheet) {
    return ExcelRowColumnOutlineSupport.columnLayouts(sheet);
  }

  private void requireInsertColumnBounds(XSSFSheet sheet, int columnIndex, int columnCount) {
    int lastColumnIndex = lastColumnIndex(sheet);
    if (columnIndex > lastColumnIndex + 1) {
      throw new IllegalArgumentException(
          "INSERT_COLUMNS "
              + ExcelIndexDisplay.describe("columnIndex", columnIndex)
              + " must be less than or equal to last existing column + 1: "
              + ExcelIndexDisplay.columnValue(lastColumnIndex + 1));
    }
    if (columnIndex + columnCount - 1 > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_COLUMNS would exceed the maximum column index: destination last column would be "
              + ExcelIndexDisplay.columnValue(columnIndex + columnCount - 1)
              + "; maximum is "
              + ExcelIndexDisplay.columnValue(ExcelColumnSpan.MAX_COLUMN_INDEX));
    }
  }

  private void requireShiftedColumnBounds(ExcelColumnSpan columns, int delta) {
    if (columns.firstColumnIndex() + delta < 0) {
      throw new IllegalArgumentException(
          "SHIFT_COLUMNS would move "
              + ExcelIndexDisplay.describe("firstColumnIndex", columns.firstColumnIndex())
              + " by delta "
              + delta
              + " before the first worksheet column ("
              + ExcelIndexDisplay.excelColumn(0)
              + ")");
    }
    if (columns.lastColumnIndex() + delta > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "SHIFT_COLUMNS would move "
              + ExcelIndexDisplay.describe("lastColumnIndex", columns.lastColumnIndex())
              + " by delta "
              + delta
              + " beyond the maximum column "
              + ExcelIndexDisplay.columnValue(ExcelColumnSpan.MAX_COLUMN_INDEX));
    }
  }
}
