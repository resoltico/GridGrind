package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import java.util.BitSet;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/** Owns adjacent visual-format propagation for inserted rows and columns. */
final class ExcelInsertedStructureFormattingSupport {
  private ExcelInsertedStructureFormattingSupport() {}

  static void copyAdjacentVisualFormattingIntoInsertedRows(
      XSSFSheet sheet, int rowIndex, int rowCount, int lastRowIndexBeforeShift) {
    Row templateRow = visualTemplateRowForInsert(sheet, rowIndex, lastRowIndexBeforeShift);
    for (int insertedRowIndex = rowIndex;
        insertedRowIndex < rowIndex + rowCount;
        insertedRowIndex++) {
      XSSFRow insertedRow = sheet.createRow(insertedRowIndex);
      copyRowVisualFormatting(templateRow, insertedRow);
    }
  }

  static void copyAdjacentVisualFormattingIntoInsertedColumns(
      XSSFSheet sheet,
      Map<Integer, CTCol> explicitColumnsAfterShift,
      int columnIndex,
      int columnCount,
      int lastColumnIndexBeforeShift) {
    int templateColumnIndex =
        visualTemplateColumnForInsert(
            sheet, explicitColumnsAfterShift, columnIndex, columnCount, lastColumnIndexBeforeShift);
    CTCol templateColumnDefinition = explicitColumnsAfterShift.get(templateColumnIndex);
    if (templateColumnDefinition != null) {
      for (int insertedColumnIndex = columnIndex;
          insertedColumnIndex < columnIndex + columnCount;
          insertedColumnIndex++) {
        CTCol insertedColumnDefinition = CTCol.Factory.newInstance();
        insertedColumnDefinition.set(templateColumnDefinition);
        explicitColumnsAfterShift.put(insertedColumnIndex, insertedColumnDefinition);
      }
    }
    for (Row row : sheet) {
      copyColumnCellVisualFormatting(row, templateColumnIndex, columnIndex, columnCount);
    }
  }

  private static Row visualTemplateRowForInsert(
      XSSFSheet sheet, int rowIndex, int lastRowIndexBeforeShift) {
    Row nearestAbove = null;
    Row nearestBelow = null;
    for (Row candidateRow : sheet) {
      if (candidateRow.getRowNum() < rowIndex) {
        nearestAbove = candidateRow;
      } else if (nearestBelow == null) {
        nearestBelow = candidateRow;
      }
    }
    return Objects.requireNonNull(
        nearestAbove != null ? nearestAbove : nearestBelow,
        "INSERT_ROWS expected a physical template row after shifting row "
            + ExcelIndexDisplay.rowValue(rowIndex)
            + " inside non-empty sheet bounds ending at "
            + ExcelIndexDisplay.rowValue(lastRowIndexBeforeShift));
  }

  private static void copyRowVisualFormatting(Row templateRow, XSSFRow insertedRow) {
    insertedRow.setHeight(templateRow.getHeight());
    if (templateRow.isFormatted()) {
      insertedRow.setRowStyle(templateRow.getRowStyle());
    }
    for (Cell templateCell : templateRow) {
      Cell insertedCell = insertedRow.createCell(templateCell.getColumnIndex(), CellType.BLANK);
      insertedCell.setCellStyle(templateCell.getCellStyle());
    }
  }

  private static int visualTemplateColumnForInsert(
      XSSFSheet sheet,
      Map<Integer, CTCol> explicitColumnsAfterShift,
      int columnIndex,
      int columnCount,
      int lastColumnIndexBeforeShift) {
    BitSet visualColumns = new BitSet(lastColumnIndexBeforeShift + columnCount + 1);
    explicitColumnsAfterShift.keySet().forEach(visualColumns::set);
    for (Row row : sheet) {
      for (Cell cell : row) {
        visualColumns.set(cell.getColumnIndex());
      }
    }
    int nearestLeft = visualColumns.previousSetBit(columnIndex - 1);
    int templateColumnIndex =
        nearestLeft >= 0 ? nearestLeft : visualColumns.nextSetBit(columnIndex);
    return Objects.checkIndex(templateColumnIndex, lastColumnIndexBeforeShift + columnCount + 1);
  }

  private static void copyColumnCellVisualFormatting(
      Row row, int templateColumnIndex, int columnIndex, int columnCount) {
    Cell templateCell = row.getCell(templateColumnIndex);
    if (templateCell == null) {
      return;
    }
    for (int insertedColumnIndex = columnIndex;
        insertedColumnIndex < columnIndex + columnCount;
        insertedColumnIndex++) {
      Cell insertedCell = row.createCell(insertedColumnIndex, CellType.BLANK);
      insertedCell.setCellStyle(templateCell.getCellStyle());
    }
  }
}
