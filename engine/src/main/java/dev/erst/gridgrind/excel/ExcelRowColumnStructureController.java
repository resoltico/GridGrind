package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Owns structural row and column editing plus layout normalization for one XSSF sheet. */
final class ExcelRowColumnStructureController {
  /** Inserts one or more blank rows before the provided zero-based row index. */
  void insertRows(XSSFSheet sheet, int rowIndex, int rowCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertRowBounds(sheet, rowIndex, rowCount);
    int lastRowIndex = sheet.getLastRowNum();
    rejectAffectedRowStructuresForInsert(sheet, rowIndex); // LIM-016
    if (rowIndex <= lastRowIndex) {
      sheet.shiftRows(rowIndex, lastRowIndex, rowCount, true, false);
    }
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
    rejectAffectedRowStructuresForDelete(sheet, rows); // LIM-016
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
    rejectAffectedRowStructuresForShift(sheet, rows, delta); // LIM-016
    sheet.shiftRows(rows.firstRowIndex(), rows.lastRowIndex(), delta, true, false);
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  void insertColumns(XSSFSheet sheet, int columnIndex, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertColumnBounds(sheet, columnIndex, columnCount);
    rejectFormulaBearingWorkbookForColumnEdit(sheet.getWorkbook(), "INSERT_COLUMNS"); // LIM-017
    normalizeColumnDefinitionContainer(sheet);
    int lastColumnIndex = lastColumnIndex(sheet);
    Map<Integer, CTCol> explicitColumns = snapshotColumnDefinitions(sheet);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments =
          commentRepairSupport.expectedCommentsAfterInsertColumns(columnIndex, columnCount);
    }
    rejectAffectedColumnStructuresForInsert(sheet, columnIndex); // LIM-016
    if (columnIndex <= lastColumnIndex) {
      sheet.shiftColumns(columnIndex, lastColumnIndex, columnCount);
    }
    rebuildColumnDefinitions(sheet, shiftedForInsert(explicitColumns, columnIndex, columnCount));
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
    rejectFormulaBearingWorkbookForColumnEdit(sheet.getWorkbook(), "DELETE_COLUMNS"); // LIM-017
    normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns = snapshotColumnDefinitions(sheet);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterDeleteColumns(columns);
    }
    rejectAffectedColumnStructuresForDelete(sheet, columns); // LIM-016
    if (columns.lastColumnIndex() < lastColumnIndex) {
      sheet.shiftColumns(columns.lastColumnIndex() + 1, lastColumnIndex, -columns.count());
    }
    int clearStart = Math.max(columns.firstColumnIndex(), lastColumnIndex - columns.count() + 1);
    clearTrailingCells(sheet, clearStart, lastColumnIndex);
    rebuildColumnDefinitions(sheet, shiftedForDelete(explicitColumns, columns));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  void shiftColumns(XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    requireShiftedColumnBounds(columns, delta);
    rejectFormulaBearingWorkbookForColumnEdit(sheet.getWorkbook(), "SHIFT_COLUMNS"); // LIM-017
    normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns = snapshotColumnDefinitions(sheet);
    ExcelSheetCommentRepairSupport commentRepairSupport = new ExcelSheetCommentRepairSupport(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterShiftColumns(columns, delta);
    }
    rejectAffectedColumnStructuresForShift(sheet, columns, delta); // LIM-016
    sheet.shiftColumns(columns.firstColumnIndex(), columns.lastColumnIndex(), delta);
    rebuildColumnDefinitions(sheet, shiftedForShift(explicitColumns, columns, delta));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  void setRowVisibility(XSSFSheet sheet, ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    ensureRowsExist(sheet, rows);
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      sheet.getRow(rowIndex).setZeroHeight(hidden);
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  void setColumnVisibility(XSSFSheet sheet, ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    normalizeColumnDefinitionContainer(sheet);
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
    ensureRowsExist(sheet, rows);
    sheet.groupRow(rows.firstRowIndex(), rows.lastRowIndex());
    if (collapsed) {
      collapseRows(sheet, rows);
      return;
    }
    if (isRowGroupCollapsed(sheet, rows)) {
      expandRows(sheet, rows);
      return;
    }
    clearExpandedGroupControlRow(sheet, rows);
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  void ungroupRows(XSSFSheet sheet, ExcelRowSpan rows) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    ensureRowsExist(sheet, rows);
    if (isRowGroupCollapsed(sheet, rows)) {
      expandRows(sheet, rows);
    } else {
      clearExpandedGroupControlRow(sheet, rows);
    }
    sheet.ungroupRow(rows.firstRowIndex(), rows.lastRowIndex());
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  void groupColumns(XSSFSheet sheet, ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    normalizeColumnDefinitionContainer(sheet);
    sheet.groupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    if (collapsed) {
      collapseColumns(sheet, columns);
      canonicalizeColumnDefinitions(sheet);
      return;
    }
    clearExpandedGroupControlColumn(sheet, columns);
    canonicalizeColumnDefinitions(sheet);
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  void ungroupColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    normalizeColumnDefinitionContainer(sheet);
    prepareColumnsForUngroup(sheet, columns);
    sheet.ungroupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    canonicalizeColumnDefinitions(sheet);
  }

  /** Returns the last column index implied by cells or explicit column metadata. */
  static int lastColumnIndex(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastColumnIndex = -1;
    for (Row row : sheet) {
      lastColumnIndex = Math.max(lastColumnIndex, row.getLastCellNum() - 1);
    }
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        lastColumnIndex = Math.max(lastColumnIndex, (int) col.getMax() - 1);
      }
    }
    return lastColumnIndex;
  }

  /** Returns the sheet column layouts including hidden, outline, and collapsed state. */
  List<WorkbookReadResult.ColumnLayout> columnLayouts(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      return List.of();
    }
    Map<Integer, CTCol> effectiveColumns = effectiveColumnDefinitions(sheet);
    List<WorkbookReadResult.ColumnLayout> columns = new ArrayList<>(lastColumnIndex + 1);
    for (int columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex++) {
      CTCol columnDefinition = effectiveColumns.get(columnIndex);
      columns.add(
          new WorkbookReadResult.ColumnLayout(
              columnIndex,
              sheet.getColumnWidth(columnIndex) / 256.0d,
              columnDefinition != null && columnDefinition.getHidden(),
              columnDefinition == null ? 0 : (int) columnDefinition.getOutlineLevel(),
              columnDefinition != null && columnDefinition.getCollapsed()));
    }
    return List.copyOf(columns);
  }

  /** Returns the sheet row layouts including hidden, outline, and collapsed state. */
  List<WorkbookReadResult.RowLayout> rowLayouts(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastRowIndex = sheet.getLastRowNum();
    if (lastRowIndex < 0) {
      return List.of();
    }
    List<WorkbookReadResult.RowLayout> rows = new ArrayList<>(lastRowIndex + 1);
    for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      double heightPoints =
          row == null ? sheet.getDefaultRowHeight() / 20.0d : row.getHeight() / 20.0d;
      rows.add(
          new WorkbookReadResult.RowLayout(
              rowIndex,
              heightPoints,
              row instanceof XSSFRow xssfRow && xssfRow.getZeroHeight(),
              row == null ? 0 : Math.max(0, row.getOutlineLevel()),
              row instanceof XSSFRow xssfRow && xssfRow.getCTRow().getCollapsed()));
    }
    return List.copyOf(rows);
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

  void rejectAffectedRowStructuresForInsert(XSSFSheet sheet, int rowIndex) { // LIM-016
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastRow() >= rowIndex) {
        throw unsupportedStructure(
            "INSERT_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastRow() >= rowIndex) {
      throw unsupportedStructure(
          "INSERT_ROWS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "row structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastRow() >= rowIndex) {
        throw unsupportedStructure(
            "INSERT_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move data validations are not supported");
      }
    }
  }

  void rejectAffectedRowStructuresForDelete(XSSFSheet sheet, ExcelRowSpan rows) { // LIM-016
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastRow() >= rows.firstRowIndex()) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move or truncate tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastRow() >= rows.firstRowIndex()) {
      throw unsupportedStructure(
          "DELETE_ROWS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "row structural edits that would move or truncate sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastRow() >= rows.firstRowIndex()) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move or truncate data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForRowDelete(sheet.getWorkbook(), sheet, rows); // LIM-018
  }

  void rejectAffectedRowStructuresForShift(
      XSSFSheet sheet, ExcelRowSpan rows, int delta) { // LIM-016
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (affectsRows(range, rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && affectsRows(autofilter, rows, delta)) {
      throw unsupportedStructure(
          "SHIFT_ROWS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "row structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (affectsRows(range, rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForRowShift(sheet.getWorkbook(), sheet, rows, delta); // LIM-018
  }

  void rejectAffectedColumnStructuresForInsert(XSSFSheet sheet, int columnIndex) { // LIM-016
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastColumn() >= columnIndex) {
        throw unsupportedStructure(
            "INSERT_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastColumn() >= columnIndex) {
      throw unsupportedStructure(
          "INSERT_COLUMNS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "column structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastColumn() >= columnIndex) {
        throw unsupportedStructure(
            "INSERT_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move data validations are not supported");
      }
    }
  }

  void rejectAffectedColumnStructuresForDelete(
      XSSFSheet sheet, ExcelColumnSpan columns) { // LIM-016
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastColumn() >= columns.firstColumnIndex()) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move or truncate tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastColumn() >= columns.firstColumnIndex()) {
      throw unsupportedStructure(
          "DELETE_COLUMNS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "column structural edits that would move or truncate sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastColumn() >= columns.firstColumnIndex()) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move or truncate data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForColumnDelete(sheet.getWorkbook(), sheet, columns); // LIM-018
  }

  void rejectAffectedColumnStructuresForShift(
      XSSFSheet sheet, ExcelColumnSpan columns, int delta) { // LIM-016
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (affectsColumns(range, columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && affectsColumns(autofilter, columns, delta)) {
      throw unsupportedStructure(
          "SHIFT_COLUMNS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "column structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (affectsColumns(range, columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForColumnShift(
        sheet.getWorkbook(), sheet, columns, delta); // LIM-018
  }

  void rejectDestructiveNamedRangesForRowDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (overlapsRows(namedRange.range(), rows)) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            namedRange.label(),
            "row structural edits that would truncate range-backed named ranges are not supported");
      }
    }
  }

  void rejectDestructiveNamedRangesForRowShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (shiftWouldCorruptRows(namedRange.range(), rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            namedRange.label(),
            "row structural edits that would overwrite or partially move range-backed named ranges"
                + " are not supported");
      }
    }
  }

  void rejectDestructiveNamedRangesForColumnDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (overlapsColumns(namedRange.range(), columns)) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            namedRange.label(),
            "column structural edits that would truncate range-backed named ranges are not"
                + " supported");
      }
    }
  }

  void rejectDestructiveNamedRangesForColumnShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (shiftWouldCorruptColumns(namedRange.range(), columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            namedRange.label(),
            "column structural edits that would overwrite or partially move range-backed named"
                + " ranges are not supported");
      }
    }
  }

  private void rejectFormulaBearingWorkbookForColumnEdit(
      XSSFWorkbook workbook, String operationType) { // LIM-017
    if (workbookContainsFormulas(workbook)) {
      throw new IllegalArgumentException(
          operationType
              + " cannot run while workbook formulas are present; Apache POI leaves some column"
              + " references stale during column structural edits");
    }
    if (workbookContainsFormulaDefinedNames(workbook)) {
      throw new IllegalArgumentException(
          operationType
              + " cannot run while formula-defined names are present; Apache POI leaves some"
              + " column references stale during column structural edits");
    }
  }

  private static boolean workbookContainsFormulas(XSSFWorkbook workbook) {
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      for (Row row : workbook.getSheetAt(sheetIndex)) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            return true;
          }
        }
      }
    }
    return false;
  }

  private static boolean workbookContainsFormulaDefinedNames(XSSFWorkbook workbook) {
    return workbookContainsFormulaDefinedNames(workbook, workbook.getAllNames());
  }

  /** Returns whether any provided defined name resolves to a non-range formula target. */
  static boolean workbookContainsFormulaDefinedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    for (Name name : names) {
      String refersToFormula = name.getRefersToFormula();
      if (refersToFormula == null) {
        continue;
      }
      if (refersToFormula.isBlank()) {
        continue;
      }
      ExcelNamedRangeScope scope = scopeForName(workbook, name);
      if (ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope).isEmpty()) {
        return true;
      }
    }
    return false;
  }

  /**
   * Returns the typed target of a range-backed defined name or empty when the name is blank, unset,
   * or not backed by one contiguous range.
   */
  static Optional<ExcelNamedRangeTarget> resolvedRangeBackedTarget(
      XSSFWorkbook workbook, Name name) {
    String refersToFormula = name.getRefersToFormula();
    if (refersToFormula == null) {
      return Optional.empty();
    }
    if (refersToFormula.isBlank()) {
      return Optional.empty();
    }
    return ExcelNamedRangeTargets.resolveTarget(refersToFormula, scopeForName(workbook, name));
  }

  /** Returns only the defined names that resolve to one contiguous range target. */
  static List<ResolvedNamedRange> resolvedRangeBackedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    List<ResolvedNamedRange> resolved = new ArrayList<>();
    for (Name name : names) {
      resolvedRangeBackedTarget(workbook, name)
          .ifPresent(
              target ->
                  resolved.add(
                      new ResolvedNamedRange(
                          name.getNameName(), target, ExcelRange.parse(target.range()))));
    }
    return List.copyOf(resolved);
  }

  private static ExcelNamedRangeScope scopeForName(XSSFWorkbook workbook, Name name) {
    if (name.getSheetIndex() < 0) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    return new ExcelNamedRangeScope.SheetScope(workbook.getSheetName(name.getSheetIndex()));
  }

  private static IllegalArgumentException unsupportedStructure(
      String operationType, XSSFSheet sheet, String structureLabel, String reason) {
    return new IllegalArgumentException(
        operationType
            + " cannot move "
            + structureLabel
            + " on sheet '"
            + sheet.getSheetName()
            + "'; "
            + reason);
  }

  /** Returns whether shifting the requested row span would overlap the provided range. */
  static boolean affectsRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    int affectedFirstRow = Math.min(rows.firstRowIndex(), rows.firstRowIndex() + delta);
    int affectedLastRow = Math.max(rows.lastRowIndex(), rows.lastRowIndex() + delta);
    return range.firstRow() <= affectedLastRow && range.lastRow() >= affectedFirstRow;
  }

  /** Returns whether shifting the requested column span would overlap the provided range. */
  static boolean affectsColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    int affectedFirstColumn =
        Math.min(columns.firstColumnIndex(), columns.firstColumnIndex() + delta);
    int affectedLastColumn = Math.max(columns.lastColumnIndex(), columns.lastColumnIndex() + delta);
    return range.firstColumn() <= affectedLastColumn && range.lastColumn() >= affectedFirstColumn;
  }

  private static boolean overlapsRows(ExcelRange range, ExcelRowSpan rows) {
    return range.firstRow() <= rows.lastRowIndex() && range.lastRow() >= rows.firstRowIndex();
  }

  private static boolean overlapsColumns(ExcelRange range, ExcelColumnSpan columns) {
    return range.firstColumn() <= columns.lastColumnIndex()
        && range.lastColumn() >= columns.firstColumnIndex();
  }

  private static boolean fullyWithinRows(ExcelRange range, ExcelRowSpan rows) {
    return range.firstRow() >= rows.firstRowIndex() && range.lastRow() <= rows.lastRowIndex();
  }

  private static boolean fullyWithinColumns(ExcelRange range, ExcelColumnSpan columns) {
    return range.firstColumn() >= columns.firstColumnIndex()
        && range.lastColumn() <= columns.lastColumnIndex();
  }

  /** Returns whether moving the requested row span would overwrite or partially move the range. */
  static boolean shiftWouldCorruptRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    if (fullyWithinRows(range, rows)) {
      return false;
    }
    ExcelRowSpan destination =
        new ExcelRowSpan(rows.firstRowIndex() + delta, rows.lastRowIndex() + delta);
    return overlapsRows(range, rows) || overlapsRows(range, destination);
  }

  /**
   * Returns whether moving the requested column span would overwrite or partially move the range.
   */
  static boolean shiftWouldCorruptColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    if (fullyWithinColumns(range, columns)) {
      return false;
    }
    ExcelColumnSpan destination =
        new ExcelColumnSpan(columns.firstColumnIndex() + delta, columns.lastColumnIndex() + delta);
    return overlapsColumns(range, columns) || overlapsColumns(range, destination);
  }

  /** Typed resolved view of a range-backed defined name for structural guard evaluation. */
  record ResolvedNamedRange(String name, ExcelNamedRangeTarget target, ExcelRange range) {
    ResolvedNamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(range, "range must not be null");
    }

    private boolean targetsSheet(String sheetName) {
      return target.sheetName().equals(sheetName);
    }

    private String label() {
      return "named range '" + name + "'";
    }
  }

  private static ExcelRange sheetAutofilterRange(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return null;
    }
    return parsedRange(
        sheet.getCTWorksheet().getAutoFilter().getRef(),
        "sheet autofilter",
        sheet.getCTWorksheet().getAutoFilter().getRef());
  }

  private static List<String> dataValidationRanges(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    List<String> ranges = new ArrayList<>();
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      ranges.addAll(ExcelSqrefSupport.normalizedSqref(validation.getSqref()));
    }
    return List.copyOf(ranges);
  }

  private static ExcelRange parsedRange(String rawRange, String structureType, String detail) {
    ExcelRange range = ExcelSheetStructureSupport.parseRangeOrNull(rawRange);
    if (range == null) {
      throw new IllegalArgumentException(
          "Stored " + structureType + " range is invalid and cannot be normalized: " + detail);
    }
    return range;
  }

  private static void ensureRowsExist(XSSFSheet sheet, ExcelRowSpan rows) {
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      if (sheet.getRow(rowIndex) == null) {
        sheet.createRow(rowIndex);
      }
    }
  }

  private static void expandRows(XSSFSheet sheet, ExcelRowSpan rows) {
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      row.setZeroHeight(false);
    }
    clearExpandedGroupControlRow(sheet, rows);
  }

  private static boolean isRowGroupCollapsed(XSSFSheet sheet, ExcelRowSpan rows) {
    if (rows.lastRowIndex() >= ExcelRowSpan.MAX_ROW_INDEX) {
      return false;
    }
    Row controlRow = sheet.getRow(rows.lastRowIndex() + 1);
    return controlRow instanceof XSSFRow xssfRow && xssfRow.getCTRow().getCollapsed();
  }

  private static void clearExpandedGroupControlRow(XSSFSheet sheet, ExcelRowSpan rows) {
    if (rows.lastRowIndex() < ExcelRowSpan.MAX_ROW_INDEX) {
      Row controlRow = sheet.getRow(rows.lastRowIndex() + 1);
      if (controlRow instanceof XSSFRow xssfRow) {
        xssfRow.getCTRow().setCollapsed(false);
      }
    }
  }

  private static void collapseRows(XSSFSheet sheet, ExcelRowSpan rows) {
    XSSFRow controlRow = null;
    if (rows.lastRowIndex() < ExcelRowSpan.MAX_ROW_INDEX) {
      Row existingControlRow = sheet.getRow(rows.lastRowIndex() + 1);
      controlRow =
          existingControlRow instanceof XSSFRow xssfRow
              ? xssfRow
              : (XSSFRow) sheet.createRow(rows.lastRowIndex() + 1);
      controlRow.setZeroHeight(false);
    }
    sheet.setRowGroupCollapsed(rows.firstRowIndex(), true);
    if (controlRow != null) {
      controlRow.getCTRow().setCollapsed(true);
    }
  }

  private static void prepareColumnsForUngroup(XSSFSheet sheet, ExcelColumnSpan columns) {
    for (int columnIndex = columns.firstColumnIndex();
        columnIndex <= columns.lastColumnIndex();
        columnIndex++) {
      sheet.setColumnHidden(columnIndex, false);
    }
    clearExpandedGroupControlColumn(sheet, columns);
  }

  private static void collapseColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    sheet.setColumnGroupCollapsed(columns.firstColumnIndex(), true);
  }

  /**
   * Clears only the control-column collapsed marker for one explicit group boundary.
   *
   * <p>Grouping columns with {@code collapsed=false} creates a new expanded outline level; it is
   * not the same operation as expanding an already-collapsed group. Delegating that case to POI's
   * {@code setColumnGroupCollapsed(..., false)} can walk and mutate overlapping descendant groups,
   * which is both semantically broader than GridGrind's request and vulnerable to POI/XMLBeans
   * crashes on split {@code CTCol} ranges. GridGrind therefore clears only the target group's own
   * control marker here and leaves existing descendant collapsed state intact.
   */
  private static void clearExpandedGroupControlColumn(XSSFSheet sheet, ExcelColumnSpan columns) {
    if (columns.lastColumnIndex() < ExcelColumnSpan.MAX_COLUMN_INDEX) {
      setColumnCollapsed(sheet, columns.lastColumnIndex() + 1, false);
    }
  }

  /** Applies a collapsed marker to the explicit column definition that owns the target index. */
  static void setColumnCollapsed(XSSFSheet sheet, int columnIndex, boolean collapsed) {
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 >= col.getMin() && columnIndex + 1 <= col.getMax()) {
          col.setCollapsed(collapsed);
        }
      }
    }
  }

  private static void clearTrailingCells(
      XSSFSheet sheet, int firstColumnIndex, int lastColumnIndex) {
    for (Row row : sheet) {
      for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
          continue;
        }
        cell.removeHyperlink();
        ExcelSheetAnnotationSupport.clearCellComment(cell);
        row.removeCell(cell);
      }
    }
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> snapshotColumnDefinitions(XSSFSheet sheet) {
    Map<Integer, CTCol> explicitColumns = new LinkedHashMap<>();
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        for (int columnIndex = (int) col.getMin() - 1;
            columnIndex <= (int) col.getMax() - 1;
            columnIndex++) {
          CTCol definition = copyOf(col);
          definition.setMin(columnIndex + 1L);
          definition.setMax(columnIndex + 1L);
          explicitColumns.put(columnIndex, definition);
        }
      }
    }
    return Map.copyOf(explicitColumns);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> shiftedForInsert(
      Map<Integer, CTCol> explicitColumns, int columnIndex, int columnCount) {
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) ->
            shifted.put(
                existingColumnIndex >= columnIndex
                    ? existingColumnIndex + columnCount
                    : existingColumnIndex,
                definition));
    return Map.copyOf(shifted);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> shiftedForDelete(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns) {
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          if (existingColumnIndex < columns.firstColumnIndex()) {
            shifted.put(existingColumnIndex, definition);
            return;
          }
          if (existingColumnIndex > columns.lastColumnIndex()) {
            shifted.put(existingColumnIndex - columns.count(), definition);
          }
        });
    return Map.copyOf(shifted);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> shiftedForShift(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns, int delta) {
    int destinationFirstColumn = columns.firstColumnIndex() + delta;
    int destinationLastColumn = columns.lastColumnIndex() + delta;
    int overwrittenFirstColumn = Math.min(destinationFirstColumn, destinationLastColumn);
    int overwrittenLastColumn = Math.max(destinationFirstColumn, destinationLastColumn);
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          boolean inSource =
              existingColumnIndex >= columns.firstColumnIndex()
                  && existingColumnIndex <= columns.lastColumnIndex();
          boolean overwrittenDestination =
              existingColumnIndex >= overwrittenFirstColumn
                  && existingColumnIndex <= overwrittenLastColumn;
          if (!inSource && !overwrittenDestination) {
            shifted.put(existingColumnIndex, definition);
          }
        });
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          if (existingColumnIndex >= columns.firstColumnIndex()
              && existingColumnIndex <= columns.lastColumnIndex()) {
            shifted.put(existingColumnIndex + delta, definition);
          }
        });
    return Map.copyOf(shifted);
  }

  private static void normalizeColumnDefinitionContainer(XSSFSheet sheet) {
    if (!requiresColumnDefinitionCanonicalization(sheet)) {
      return;
    }
    canonicalizeColumnDefinitions(sheet);
  }

  static void canonicalizeColumnDefinitions(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    rebuildColumnDefinitions(sheet, effectiveColumnDefinitions(sheet));
  }

  private static void rebuildColumnDefinitions(
      XSSFSheet sheet, Map<Integer, CTCol> explicitColumns) {
    CTCols cols = CTCols.Factory.newInstance();
    explicitColumns.entrySet().stream()
        .sorted(Map.Entry.comparingByKey())
        .forEach(
            entry -> {
              CTCol definition = copyOf(entry.getValue());
              definition.setMin(entry.getKey() + 1L);
              definition.setMax(entry.getKey() + 1L);
              cols.addNewCol().set(definition);
            });
    sheet.getCTWorksheet().setColsArray(new CTCols[] {cols});
  }

  private static boolean requiresColumnDefinitionCanonicalization(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().sizeOfColsArray() != 1) {
      return true;
    }
    boolean[] seenColumns = new boolean[ExcelColumnSpan.MAX_COLUMN_INDEX + 1];
    for (CTCol col : sheet.getCTWorksheet().getColsArray(0).getColList()) {
      if (col.getMin() != col.getMax()) {
        return true;
      }
      if (isSemanticallyEmptyColumnDefinition(col)) {
        return true;
      }
      for (int columnIndex = (int) col.getMin() - 1;
          columnIndex <= (int) col.getMax() - 1;
          columnIndex++) {
        if (seenColumns[columnIndex]) {
          return true;
        }
        seenColumns[columnIndex] = true;
      }
    }
    return false;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> effectiveColumnDefinitions(XSSFSheet sheet) {
    Map<Integer, CTCol> explicitColumns = snapshotColumnDefinitions(sheet);
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      return Map.of();
    }
    Map<Integer, CTCol> effectiveColumns = new LinkedHashMap<>();
    for (int columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex++) {
      CTCol effectiveDefinition =
          effectiveColumnDefinition(sheet, columnIndex, explicitColumns.get(columnIndex));
      if (effectiveDefinition != null) {
        effectiveColumns.put(columnIndex, effectiveDefinition);
      }
    }
    return Map.copyOf(effectiveColumns);
  }

  private static CTCol effectiveColumnDefinition(
      XSSFSheet sheet, int columnIndex, CTCol baseDefinition) {
    boolean hidden = false;
    long outlineLevel = 0L;
    boolean collapsed = false;
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 < col.getMin() || columnIndex + 1 > col.getMax()) {
          continue;
        }
        hidden |= col.getHidden();
        outlineLevel = Math.max(outlineLevel, col.getOutlineLevel());
        collapsed |= col.getCollapsed();
      }
    }
    if (!hasMeaningfulColumnDefinition(baseDefinition, hidden, outlineLevel, collapsed)) {
      return null;
    }
    CTCol effectiveDefinition = copyOf(baseDefinition);
    effectiveDefinition.setMin(columnIndex + 1L);
    effectiveDefinition.setMax(columnIndex + 1L);
    effectiveDefinition.setHidden(hidden);
    effectiveDefinition.setOutlineLevel((short) outlineLevel);
    effectiveDefinition.setCollapsed(collapsed);
    return effectiveDefinition;
  }

  private static boolean hasMeaningfulColumnDefinition(
      CTCol baseDefinition, boolean hidden, long outlineLevel, boolean collapsed) {
    return hidden
        || outlineLevel > 0L
        || collapsed
        || (baseDefinition != null && !isSemanticallyEmptyColumnDefinition(baseDefinition));
  }

  private static boolean isSemanticallyEmptyColumnDefinition(CTCol definition) {
    return !definition.getHidden()
        && definition.getOutlineLevel() == 0
        && !definition.getCollapsed()
        && !definition.getCustomWidth()
        && !definition.getBestFit()
        && !definition.getPhonetic()
        && (!definition.isSetStyle() || definition.getStyle() == 0L);
  }

  private static CTCol copyOf(CTCol original) {
    CTCol copy = CTCol.Factory.newInstance();
    copy.set(original);
    return copy;
  }
}
