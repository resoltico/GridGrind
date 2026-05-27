package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import java.util.Objects;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Column operations for one sheet wrapper. */
final class ExcelSheetColumnSupport {
  private final Sheet sheet;
  private final ExcelFormulaRuntime formulaRuntime;
  private final DataFormatter dataFormatter;
  private final ExcelColumnStructureController columnStructureController;

  ExcelSheetColumnSupport(
      Sheet sheet,
      ExcelFormulaRuntime formulaRuntime,
      DataFormatter dataFormatter,
      ExcelColumnStructureController columnStructureController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.dataFormatter = Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
    this.columnStructureController =
        Objects.requireNonNull(
            columnStructureController, "columnStructureController must not be null");
  }

  void setColumnWidth(int firstColumnIndex, int lastColumnIndex, double widthCharacters) {
    requireNonNegative(firstColumnIndex, "firstColumnIndex");
    requireNonNegative(lastColumnIndex, "lastColumnIndex");
    requireOrderedSpan(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");
    int widthUnits = ExcelSheetStructureSupport.toColumnWidthUnits(widthCharacters);
    for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
      sheet.setColumnWidth(columnIndex, widthUnits);
    }
  }

  void insertColumns(int columnIndex, int columnCount) {
    columnStructureController.insertColumns(xssfSheet(), columnIndex, columnCount);
  }

  void deleteColumns(ExcelColumnSpan columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    columnStructureController.deleteColumns(xssfSheet(), columns);
  }

  void shiftColumns(ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(columns, "columns must not be null");
    columnStructureController.shiftColumns(xssfSheet(), columns, delta);
  }

  void setColumnVisibility(ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(columns, "columns must not be null");
    columnStructureController.setColumnVisibility(xssfSheet(), columns, hidden);
  }

  void groupColumns(ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(columns, "columns must not be null");
    columnStructureController.groupColumns(xssfSheet(), columns, collapsed);
  }

  void ungroupColumns(ExcelColumnSpan columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    columnStructureController.ungroupColumns(xssfSheet(), columns);
  }

  int lastColumnIndex() {
    return columnStructureController.lastColumnIndex(xssfSheet());
  }

  void autoSizeColumns(String sheetName) {
    DeterministicColumnSizer.autoSize(sheet, sheetName, dataFormatter, formulaRuntime);
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
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
