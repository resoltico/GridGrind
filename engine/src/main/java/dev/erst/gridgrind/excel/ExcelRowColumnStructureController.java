package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;

/** Owns structural row and column editing plus layout normalization for one XSSF sheet. */
final class ExcelRowColumnStructureController {
  /** Inserts one or more blank rows before the provided zero-based row index. */
  void insertRows(XSSFSheet sheet, int rowIndex, int rowCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertRowBounds(sheet, rowIndex, rowCount);
    int lastRowIndex = sheet.getLastRowNum();
    List<CTDataValidation> expectedValidations =
        ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(
            sheet, rowIndex, rowCount);
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForInsert(
        sheet, rowIndex); // LIM-016
    if (rowIndex <= lastRowIndex) {
      sheet.shiftRows(rowIndex, lastRowIndex, rowCount, true, false);
      ExcelInsertedStructureFormattingSupport.copyAdjacentVisualFormattingIntoInsertedRows(
          sheet, rowIndex, rowCount, lastRowIndex);
    }
    ExcelDataValidationStructureSupport.replaceDataValidations(sheet, expectedValidations);
  }

  /** Deletes the requested inclusive zero-based row band. */
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
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForDelete(
        sheet, rows); // LIM-016
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

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  void shiftRows(XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    requireShiftedRowBounds(rows, delta);
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForShift(
        sheet, rows, delta); // LIM-016
    sheet.shiftRows(rows.firstRowIndex(), rows.lastRowIndex(), delta, true, false);
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  void insertColumns(XSSFSheet sheet, int columnIndex, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertColumnBounds(sheet, columnIndex, columnCount);
    ExcelRowColumnStructureGuardSupport.rejectFormulaBearingWorkbookForColumnEdit(
        sheet.getWorkbook(), "INSERT_COLUMNS"); // LIM-017
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
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForInsert(
        sheet, columnIndex); // LIM-016
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

  /** Deletes the requested inclusive zero-based column band. */
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
        sheet.getWorkbook(), "DELETE_COLUMNS"); // LIM-017
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns =
        ExcelRowColumnOutlineSupport.snapshotColumnDefinitions(sheet);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterDeleteColumns(columns);
    }
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForDelete(
        sheet, columns); // LIM-016
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

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  void shiftColumns(XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    requireShiftedColumnBounds(columns, delta);
    ExcelRowColumnStructureGuardSupport.rejectFormulaBearingWorkbookForColumnEdit(
        sheet.getWorkbook(), "SHIFT_COLUMNS"); // LIM-017
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
        sheet, columns, delta); // LIM-016
    sheet.shiftColumns(columns.firstColumnIndex(), columns.lastColumnIndex(), delta);
    ExcelColumnDefinitionSupport.rebuildColumnDefinitions(
        sheet, ExcelRowColumnOutlineSupport.shiftedForShift(explicitColumns, columns, delta));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  void setRowVisibility(XSSFSheet sheet, ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    ExcelRowColumnOutlineSupport.ensureRowsExist(sheet, rows);
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      sheet.getRow(rowIndex).setZeroHeight(hidden);
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  void setColumnVisibility(XSSFSheet sheet, ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    for (int columnIndex = columns.firstColumnIndex();
        columnIndex <= columns.lastColumnIndex();
        columnIndex++) {
      sheet.setColumnHidden(columnIndex, hidden);
    }
    canonicalizeColumnDefinitions(sheet);
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
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

  /** Removes outline grouping from the requested inclusive zero-based row band. */
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

  /** Applies one outline group to the requested inclusive zero-based column band. */
  void groupColumns(XSSFSheet sheet, ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    sheet.groupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    if (collapsed) {
      ExcelRowColumnOutlineSupport.collapseColumns(sheet, columns);
      canonicalizeColumnDefinitions(sheet);
      return;
    }
    ExcelRowColumnOutlineSupport.clearExpandedGroupControlColumn(sheet, columns);
    canonicalizeColumnDefinitions(sheet);
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  void ungroupColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    ExcelRowColumnOutlineSupport.normalizeColumnDefinitionContainer(sheet);
    ExcelRowColumnOutlineSupport.prepareColumnsForUngroup(sheet, columns);
    sheet.ungroupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    canonicalizeColumnDefinitions(sheet);
  }

  /** Returns the last column index implied by cells or explicit column metadata. */
  static int lastColumnIndex(XSSFSheet sheet) {
    return ExcelRowColumnOutlineSupport.lastColumnIndex(sheet);
  }

  /** Returns the sheet column layouts including hidden, outline, and collapsed state. */
  List<WorkbookSheetResult.ColumnLayout> columnLayouts(XSSFSheet sheet) {
    return ExcelRowColumnOutlineSupport.columnLayouts(sheet);
  }

  /** Returns the sheet row layouts including hidden, outline, and collapsed state. */
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

  void rejectAffectedRowStructuresForInsert(XSSFSheet sheet, int rowIndex) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForInsert(sheet, rowIndex);
  }

  void rejectAffectedRowStructuresForDelete(XSSFSheet sheet, ExcelRowSpan rows) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForDelete(sheet, rows);
  }

  void rejectAffectedRowStructuresForShift(XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForShift(sheet, rows, delta);
  }

  void rejectAffectedColumnStructuresForInsert(XSSFSheet sheet, int columnIndex) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForInsert(sheet, columnIndex);
  }

  void rejectAffectedColumnStructuresForDelete(XSSFSheet sheet, ExcelColumnSpan columns) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForDelete(sheet, columns);
  }

  void rejectAffectedColumnStructuresForShift(XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForShift(
        sheet, columns, delta);
  }

  void rejectDestructiveNamedRangesForRowDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForRowDelete(
        workbook, sheet, rows);
  }

  void rejectDestructiveNamedRangesForRowShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForRowShift(
        workbook, sheet, rows, delta);
  }

  void rejectDestructiveNamedRangesForColumnDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForColumnDelete(
        workbook, sheet, columns);
  }

  void rejectDestructiveNamedRangesForColumnShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForColumnShift(
        workbook, sheet, columns, delta);
  }

  static boolean workbookContainsFormulaDefinedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    return ExcelRowColumnStructureGuardSupport.workbookContainsFormulaDefinedNames(workbook, names);
  }

  static Optional<ExcelNamedRangeTarget> resolvedRangeBackedTarget(
      XSSFWorkbook workbook, Name name) {
    return ExcelRowColumnStructureGuardSupport.resolvedRangeBackedTarget(workbook, name);
  }

  static List<ResolvedNamedRange> resolvedRangeBackedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    return ExcelRowColumnStructureGuardSupport.resolvedRangeBackedNames(workbook, names).stream()
        .map(result -> new ResolvedNamedRange(result.name(), result.target(), result.range()))
        .toList();
  }

  static boolean affectsRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    return ExcelRowColumnStructureGuardSupport.affectsRows(range, rows, delta);
  }

  static boolean affectsColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    return ExcelRowColumnStructureGuardSupport.affectsColumns(range, columns, delta);
  }

  static boolean shiftWouldCorruptRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    return ExcelRowColumnStructureGuardSupport.shiftWouldCorruptRows(range, rows, delta);
  }

  static boolean shiftWouldCorruptColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    return ExcelRowColumnStructureGuardSupport.shiftWouldCorruptColumns(range, columns, delta);
  }

  static void setColumnCollapsed(XSSFSheet sheet, int columnIndex, boolean collapsed) {
    ExcelRowColumnOutlineSupport.setColumnCollapsed(sheet, columnIndex, collapsed);
  }

  static void canonicalizeColumnDefinitions(XSSFSheet sheet) {
    ExcelRowColumnOutlineSupport.canonicalizeColumnDefinitions(sheet);
  }

  /** Typed resolved view of a range-backed defined name for structural guard evaluation. */
  record ResolvedNamedRange(String name, ExcelNamedRangeTarget target, ExcelRange range) {
    ResolvedNamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(range, "range must not be null");
    }
  }
}
