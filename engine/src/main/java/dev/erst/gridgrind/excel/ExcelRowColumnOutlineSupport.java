package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/** Owns row and column layout, outline, and column-definition normalization helpers. */
final class ExcelRowColumnOutlineSupport {
  private ExcelRowColumnOutlineSupport() {}

  static int lastColumnIndex(XSSFSheet sheet) {
    return ExcelColumnDefinitionSupport.lastColumnIndex(sheet);
  }

  static List<WorkbookSheetResult.ColumnLayout> columnLayouts(XSSFSheet sheet) {
    return ExcelColumnDefinitionSupport.columnLayouts(sheet);
  }

  static List<WorkbookSheetResult.RowLayout> rowLayouts(XSSFSheet sheet) {
    int lastRowIndex = sheet.getLastRowNum();
    if (lastRowIndex < 0) {
      return List.of();
    }
    List<WorkbookSheetResult.RowLayout> rows = new ArrayList<>(lastRowIndex + 1);
    for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      double heightPoints =
          row == null ? sheet.getDefaultRowHeight() / 20.0d : row.getHeight() / 20.0d;
      rows.add(
          new WorkbookSheetResult.RowLayout(
              rowIndex,
              heightPoints,
              row instanceof XSSFRow xssfRow && xssfRow.getZeroHeight(),
              row == null ? 0 : Math.max(0, row.getOutlineLevel()),
              row instanceof XSSFRow xssfRow && xssfRow.getCTRow().getCollapsed()));
    }
    return List.copyOf(rows);
  }

  static void ensureRowsExist(XSSFSheet sheet, ExcelRowSpan rows) {
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      if (sheet.getRow(rowIndex) == null) {
        sheet.createRow(rowIndex);
      }
    }
  }

  static void expandRows(XSSFSheet sheet, ExcelRowSpan rows) {
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      row.setZeroHeight(false);
    }
    clearExpandedGroupControlRow(sheet, rows);
  }

  static boolean isRowGroupCollapsed(XSSFSheet sheet, ExcelRowSpan rows) {
    if (rows.lastRowIndex() >= ExcelRowSpan.MAX_ROW_INDEX) {
      return false;
    }
    Row controlRow = sheet.getRow(rows.lastRowIndex() + 1);
    return controlRow instanceof XSSFRow xssfRow && xssfRow.getCTRow().getCollapsed();
  }

  static void clearExpandedGroupControlRow(XSSFSheet sheet, ExcelRowSpan rows) {
    if (rows.lastRowIndex() < ExcelRowSpan.MAX_ROW_INDEX) {
      Row controlRow = sheet.getRow(rows.lastRowIndex() + 1);
      if (controlRow instanceof XSSFRow xssfRow) {
        xssfRow.getCTRow().setCollapsed(false);
      }
    }
  }

  static void collapseRows(XSSFSheet sheet, ExcelRowSpan rows) {
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

  static void prepareColumnsForUngroup(XSSFSheet sheet, ExcelColumnSpan columns) {
    for (int columnIndex = columns.firstColumnIndex();
        columnIndex <= columns.lastColumnIndex();
        columnIndex++) {
      sheet.setColumnHidden(columnIndex, false);
    }
    clearExpandedGroupControlColumn(sheet, columns);
  }

  static void collapseColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    sheet.setColumnGroupCollapsed(columns.firstColumnIndex(), true);
  }

  static void clearExpandedGroupControlColumn(XSSFSheet sheet, ExcelColumnSpan columns) {
    if (columns.lastColumnIndex() < ExcelColumnSpan.MAX_COLUMN_INDEX) {
      setColumnCollapsed(sheet, columns.lastColumnIndex() + 1, false);
    }
  }

  static void setColumnCollapsed(XSSFSheet sheet, int columnIndex, boolean collapsed) {
    ExcelColumnDefinitionSupport.setColumnCollapsed(sheet, columnIndex, collapsed);
  }

  static void clearTrailingCells(XSSFSheet sheet, int firstColumnIndex, int lastColumnIndex) {
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

  static Map<Integer, CTCol> snapshotColumnDefinitions(XSSFSheet sheet) {
    return ExcelColumnDefinitionSupport.snapshotColumnDefinitions(sheet);
  }

  static Map<Integer, CTCol> shiftedForInsert(
      Map<Integer, CTCol> explicitColumns, int columnIndex, int columnCount) {
    return ExcelColumnDefinitionSupport.shiftedForInsert(explicitColumns, columnIndex, columnCount);
  }

  static Map<Integer, CTCol> shiftedForDelete(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns) {
    return ExcelColumnDefinitionSupport.shiftedForDelete(explicitColumns, columns);
  }

  static Map<Integer, CTCol> shiftedForShift(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns, int delta) {
    return ExcelColumnDefinitionSupport.shiftedForShift(explicitColumns, columns, delta);
  }

  static Map<Integer, CTCol> mutableColumnDefinitionsCopy(Map<Integer, CTCol> explicitColumns) {
    return new LinkedHashMap<>(explicitColumns);
  }

  static void normalizeColumnDefinitionContainer(XSSFSheet sheet) {
    ExcelColumnDefinitionSupport.normalizeColumnDefinitionContainer(sheet);
  }

  static void canonicalizeColumnDefinitions(XSSFSheet sheet) {
    ExcelColumnDefinitionSupport.canonicalizeColumnDefinitions(sheet);
  }
}
