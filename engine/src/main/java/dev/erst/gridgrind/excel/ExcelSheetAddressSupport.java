package dev.erst.gridgrind.excel;

import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Cell-address parsing and cell lookup helpers shared by sheet annotation helpers. */
final class ExcelSheetAddressSupport {
  private ExcelSheetAddressSupport() {}

  static CellReference parseCellReference(String address) {
    try {
      CellReference reference = new CellReference(address);
      requireValidCellReference(address, reference);
      return reference;
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  static Optional<Cell> optionalCell(Sheet sheet, String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      return Optional.empty();
    }
    return Optional.ofNullable(row.getCell(cellReference.getCol()));
  }

  static Optional<Cell> cellOrNull(Sheet sheet, String address) {
    CellReference reference = parseCellReference(address);
    Row row = sheet.getRow(reference.getRow());
    return row == null ? Optional.empty() : Optional.ofNullable(row.getCell(reference.getCol()));
  }

  static Cell getOrCreateCell(Sheet sheet, int rowIndex, int columnIndex) {
    return getOrCreateRow(sheet, rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  static Row getOrCreateRow(Sheet sheet, int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private static void requireValidCellReference(String address, CellReference cellReference) {
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
}
