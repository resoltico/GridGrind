package dev.erst.gridgrind.engine.cells;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/** Canonical cell lookup and native row creation shared by sheet annotation helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class CellAccess {
  private CellAccess() {}

  public static Optional<Cell> optionalCell(Sheet sheet, CellAddress address) {
    Row row = sheet.getRow(address.rowIndex());
    return row == null ? Optional.empty() : Optional.ofNullable(row.getCell(address.columnIndex()));
  }

  public static Cell getOrCreateCell(Sheet sheet, int rowIndex, int columnIndex) {
    return getOrCreateRow(sheet, rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }
}
