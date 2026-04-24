package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Cell-value mutation, style application, and append helpers for one sheet wrapper. */
final class ExcelSheetCellMutationSupport {
  private final Sheet sheet;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final ExcelDrawingController drawingController;

  ExcelSheetCellMutationSupport(
      Sheet sheet,
      WorkbookStyleRegistry styleRegistry,
      ExcelFormulaRuntime formulaRuntime,
      ExcelDrawingController drawingController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.styleRegistry = Objects.requireNonNull(styleRegistry, "styleRegistry must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.drawingController =
        Objects.requireNonNull(drawingController, "drawingController must not be null");
  }

  void setCell(String address, ExcelCellValue value) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(value, "value must not be null");

    CellReference cellReference = parseCellReference(address);
    writeCellValue(cellReference.getRow(), cellReference.getCol(), value);
    syncTableHeaders(
        new ExcelRange(
            cellReference.getRow(),
            cellReference.getRow(),
            cellReference.getCol(),
            cellReference.getCol()));
  }

  void setRange(String range, List<List<ExcelCellValue>> rows) {
    ExcelSheet.requireNonBlank(range, "range");
    Objects.requireNonNull(rows, "rows must not be null");
    if (rows.isEmpty()) {
      throw new IllegalArgumentException("rows must not be empty");
    }

    List<List<ExcelCellValue>> copiedRows = copyRows(rows);
    ExcelRange excelRange = ExcelRange.parse(range);
    if (copiedRows.size() != excelRange.rowCount()
        || copiedRows.getFirst().size() != excelRange.columnCount()) {
      throw new IllegalArgumentException(
          "range dimensions do not match provided values: "
              + range
              + " expects "
              + excelRange.rowCount()
              + "x"
              + excelRange.columnCount()
              + " but received "
              + copiedRows.size()
              + "x"
              + copiedRows.getFirst().size());
    }

    for (int rowOffset = 0; rowOffset < copiedRows.size(); rowOffset++) {
      List<ExcelCellValue> rowValues = copiedRows.get(rowOffset);
      for (int columnOffset = 0; columnOffset < rowValues.size(); columnOffset++) {
        writeCellValue(
            excelRange.firstRow() + rowOffset,
            excelRange.firstColumn() + columnOffset,
            rowValues.get(columnOffset));
      }
    }
    drawingController.cleanupEmptyDrawingPatriarch(xssfSheet());
    syncTableHeaders(excelRange);
  }

  void clearRange(String range) {
    ExcelSheet.requireNonBlank(range, "range");
    ExcelRange excelRange = ExcelRange.parse(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      if (row == null) {
        continue;
      }
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
          continue;
        }
        resetToDefaultStyle(cell);
        cell.removeHyperlink();
        ExcelSheetAnnotationSupport.clearCellComment(cell);
        cell.setBlank();
      }
    }
    syncTableHeaders(excelRange);
  }

  void setArrayFormula(String range, ExcelArrayFormulaDefinition formula) {
    ExcelSheet.requireNonBlank(range, "range");
    Objects.requireNonNull(formula, "formula must not be null");

    ExcelRange excelRange = ExcelRange.parse(range);
    CellRangeAddress poiRange =
        new CellRangeAddress(
            excelRange.firstRow(),
            excelRange.lastRow(),
            excelRange.firstColumn(),
            excelRange.lastColumn());
    try {
      xssfSheet().setArrayFormula(formula.formula(), poiRange);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          formulaRuntime,
          sheet.getSheetName(),
          new CellReference(excelRange.firstRow(), excelRange.firstColumn()).formatAsString(),
          formula.formula(),
          exception);
    }
    syncTableHeaders(excelRange);
  }

  void clearArrayFormula(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (!cell.isPartOfArrayFormulaGroup()) {
      throw new IllegalArgumentException("Cell " + address + " is not part of an array formula.");
    }
    CellRangeAddress removed = cell.getArrayFormulaRange();
    xssfSheet().removeArrayFormula(cell);
    syncTableHeaders(
        new ExcelRange(
            removed.getFirstRow(),
            removed.getLastRow(),
            removed.getFirstColumn(),
            removed.getLastColumn()));
    drawingController.cleanupEmptyDrawingPatriarch(xssfSheet());
  }

  void applyStyle(String range, ExcelCellStyle style) {
    ExcelSheet.requireNonBlank(range, "range");
    Objects.requireNonNull(style, "style must not be null");

    ExcelRange excelRange = ExcelRange.parse(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = getOrCreateCell(rowIndex, columnIndex);
        cell.setCellStyle(styleRegistry.mergedStyle(cell, style));
      }
    }
    syncTableHeaders(excelRange);
  }

  void appendRow(ExcelCellValue... values) {
    Objects.requireNonNull(values, "values must not be null");
    for (ExcelCellValue value : values) {
      Objects.requireNonNull(value, "values must not contain nulls");
    }

    int rowIndex = nextAppendRowIndex();
    for (int columnIndex = 0; columnIndex < values.length; columnIndex++) {
      writeCellValue(rowIndex, columnIndex, values[columnIndex]);
    }
    if (values.length > 0) {
      syncTableHeaders(new ExcelRange(rowIndex, rowIndex, 0, values.length - 1));
    }
  }

  private void writeCellValue(int rowIndex, int columnIndex, ExcelCellValue value) {
    Cell cell = getOrCreateCell(rowIndex, columnIndex);

    switch (value) {
      case ExcelCellValue.BlankValue _ -> {
        clearExistingFormula(cell);
        cell.setBlank();
      }
      case ExcelCellValue.TextValue textValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(textValue.value());
      }
      case ExcelCellValue.RichTextValue richTextValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(
            ExcelRichTextSupport.toPoiRichText(xssfSheet().getWorkbook(), richTextValue.value()));
      }
      case ExcelCellValue.NumberValue numberValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(numberValue.value());
      }
      case ExcelCellValue.BooleanValue booleanValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(booleanValue.value());
      }
      case ExcelCellValue.DateValue dateValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(dateValue.value());
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case ExcelCellValue.DateTimeValue dateTimeValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(dateTimeValue.value());
        cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      }
      case ExcelCellValue.FormulaValue formulaValue ->
          ExcelFormulaWriteSupport.setAuthoredFormula(
              cell,
              formulaValue.expression(),
              formulaRuntime,
              sheet.getSheetName(),
              new CellReference(rowIndex, columnIndex).formatAsString());
    }
  }

  private static void clearExistingFormula(Cell cell) {
    if (cell.getCellType() == CellType.FORMULA) {
      cell.removeFormula();
    }
  }

  private void syncTableHeaders(ExcelRange changedRange) {
    ExcelTableHeaderSyncSupport.syncAffectedHeaders(xssfSheet(), changedRange);
  }

  private Cell getOrCreateCell(int rowIndex, int columnIndex) {
    return getOrCreateRow(rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  private Cell requiredCell(String address) {
    CellReference cellReference = parseCellReference(address);
    requireValidCellReference(address, cellReference);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }
    Cell cell = row.getCell(cellReference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }
    return cell;
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private int nextAppendRowIndex() {
    int lastValueBearingRowIndex = -1;
    for (Row row : sheet) {
      if (rowHasValueBearingCell(row)) {
        lastValueBearingRowIndex = row.getRowNum();
      }
    }
    return lastValueBearingRowIndex + 1;
  }

  private static boolean rowHasValueBearingCell(Row row) {
    for (Cell cell : row) {
      if (cell.getCellType() != CellType.BLANK) {
        return true;
      }
    }
    return false;
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private void resetToDefaultStyle(Cell cell) {
    cell.setCellStyle(styleRegistry.defaultStyle());
  }

  private static CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  static void requireValidCellReference(String address, CellReference cellReference) {
    int row = cellReference.getRow();
    int col = cellReference.getCol();
    if (row < 0
        || col < 0
        || row > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || col > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new InvalidCellAddressException(
          address, new IllegalArgumentException("not a valid A1-style cell address: " + address));
    }
  }

  private static List<List<ExcelCellValue>> copyRows(List<List<ExcelCellValue>> rows) {
    List<List<ExcelCellValue>> copiedRows = new ArrayList<>(rows.size());
    int expectedWidth = -1;
    for (List<ExcelCellValue> row : rows) {
      Objects.requireNonNull(row, "rows must not contain nulls");
      List<ExcelCellValue> copiedRow = List.copyOf(row);
      if (copiedRow.isEmpty()) {
        throw new IllegalArgumentException("rows must not contain empty rows");
      }
      if (expectedWidth < 0) {
        expectedWidth = copiedRow.size();
      } else if (copiedRow.size() != expectedWidth) {
        throw new IllegalArgumentException("rows must describe a rectangular matrix");
      }
      for (ExcelCellValue value : copiedRow) {
        Objects.requireNonNull(value, "rows must not contain null cell values");
      }
      copiedRows.add(copiedRow);
    }
    return List.copyOf(copiedRows);
  }
}
