package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Owns workbook-level sheet lookup, creation, and cell-address access helpers. */
final class ExcelWorkbookSheetAccessSupport {
  private ExcelWorkbookSheetAccessSupport() {}

  static ExcelSheet getOrCreateSheet(ExcelWorkbook workbook, String sheetName) {
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");

    Sheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      sheet = workbook.xssfWorkbook().createSheet(sheetName);
    }
    return new ExcelSheet(
        sheet, workbook.context().styleRegistry(), workbook.context().formulaRuntime());
  }

  static ExcelSheet sheet(ExcelWorkbook workbook, String sheetName) {
    return new ExcelSheet(
        requiredSheet(workbook, sheetName),
        workbook.context().styleRegistry(),
        workbook.context().formulaRuntime());
  }

  static java.util.List<String> sheetNames(ExcelWorkbook workbook) {
    java.util.List<String> sheetNames =
        new ArrayList<>(workbook.xssfWorkbook().getNumberOfSheets());
    for (int sheetIndex = 0;
        sheetIndex < workbook.xssfWorkbook().getNumberOfSheets();
        sheetIndex++) {
      sheetNames.add(workbook.xssfWorkbook().getSheetName(sheetIndex));
    }
    return java.util.List.copyOf(sheetNames);
  }

  static Cell requiredCell(ExcelWorkbook workbook, String sheetName, String address) {
    CellReference reference = parseCellReference(address);
    Sheet sheet = requiredSheet(workbook, sheetName);
    Row row = sheet.getRow(reference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }
    Cell cell = row.getCell(reference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }
    return cell;
  }

  static Sheet requiredSheet(ExcelWorkbook workbook, String sheetName) {
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    Sheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  static int requiredSheetIndex(ExcelWorkbook workbook, String sheetName) {
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    int sheetIndex = workbook.xssfWorkbook().getSheetIndex(sheetName);
    if (sheetIndex < 0) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheetIndex;
  }

  private static CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }
}
