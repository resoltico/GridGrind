package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Row operations for one sheet wrapper. */
final class ExcelSheetRowSupport {
  private final Sheet sheet;
  private final ExcelRowStructureController rowStructureController;

  ExcelSheetRowSupport(Sheet sheet, ExcelRowStructureController rowStructureController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.rowStructureController =
        Objects.requireNonNull(rowStructureController, "rowStructureController must not be null");
  }

  void setRowHeight(int firstRowIndex, int lastRowIndex, double heightPoints) {
    requireNonNegative(firstRowIndex, "firstRowIndex");
    requireNonNegative(lastRowIndex, "lastRowIndex");
    requireOrderedSpan(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
    float heightPointsValue = ExcelSheetStructureSupport.toRowHeightPoints(heightPoints);
    for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
      getOrCreateRow(rowIndex).setHeightInPoints(heightPointsValue);
    }
  }

  void insertRows(int rowIndex, int rowCount) {
    rowStructureController.insertRows(xssfSheet(), rowIndex, rowCount);
  }

  void deleteRows(ExcelRowSpan rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowStructureController.deleteRows(xssfSheet(), rows);
  }

  void shiftRows(ExcelRowSpan rows, int delta) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowStructureController.shiftRows(xssfSheet(), rows, delta);
  }

  void setRowVisibility(ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowStructureController.setRowVisibility(xssfSheet(), rows, hidden);
  }

  void groupRows(ExcelRowSpan rows, boolean collapsed) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowStructureController.groupRows(xssfSheet(), rows, collapsed);
  }

  void ungroupRows(ExcelRowSpan rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowStructureController.ungroupRows(xssfSheet(), rows);
  }

  int physicalRowCount() {
    return sheet.getPhysicalNumberOfRows();
  }

  int lastRowIndex() {
    return sheet.getLastRowNum();
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(ExcelIndexDisplay.mustNotBeNegative(fieldName, value));
    }
  }

  private static void requireOrderedSpan(
      int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
    if (lastValue < firstValue) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeLessThan(
              lastFieldName, lastValue, firstFieldName, firstValue));
    }
  }
}
