package dev.erst.gridgrind.engine.cells;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import dev.erst.gridgrind.engine.backend.xssf.InvalidCellAddressException;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Cell-address parsing and cell lookup helpers shared by sheet annotation helpers. */
@SuppressWarnings("PMD.CommentRequired")
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_ROWS_PER_WORKSHEET,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_COLUMNS_PER_WORKSHEET
})
public final class CellAccess {
  private CellAccess() {}

  public static CellReference parseCellReference(String address) {
    try {
      CellReference reference = new CellReference(address);
      requireValidCellReference(address, reference);
      return reference;
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  public static Optional<Cell> optionalCell(Sheet sheet, String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      return Optional.empty();
    }
    return Optional.ofNullable(row.getCell(cellReference.getCol()));
  }

  public static Optional<Cell> cellOrNull(Sheet sheet, String address) {
    CellReference reference = parseCellReference(address);
    Row row = sheet.getRow(reference.getRow());
    return row == null ? Optional.empty() : Optional.ofNullable(row.getCell(reference.getCol()));
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

  private static void requireValidCellReference(String address, CellReference cellReference) {
    int row = cellReference.getRow();
    int col = cellReference.getCol();
    if (row < 0 || col < 0 || row > Limit.MAX_ROW_INDEX || col > Limit.MAX_COLUMN_INDEX) {
      throw new InvalidCellAddressException(
          address, new IllegalArgumentException("not a valid A1-style cell address: " + address));
    }
  }

  /** Preserves the bound-cell operation's explicit nonblank address precondition. */
  public static void requireAddress(String address) {
    TextValidation.requireNonBlank(address, "address");
  }

  /** Returns the native A1 rendering used for persisted cell coordinates. */
  public static String address(int rowIndex, int columnIndex) {
    return new CellReference(rowIndex, columnIndex).formatAsString();
  }

  /** Parses a mutation range with the same native rules used by formula-origin tracking. */
  public static dev.erst.gridgrind.engine.poi.CellBounds rangeBounds(String range) {
    var bounds = org.apache.poi.ss.util.CellRangeAddress.valueOf(range);
    return new dev.erst.gridgrind.engine.poi.CellBounds(
        bounds.getFirstRow(), bounds.getLastRow(), bounds.getFirstColumn(), bounds.getLastColumn());
  }
}
