package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.Objects;

/** Row insertion, movement, visibility, and grouping operations for one sheet. */
public final class ExcelSheetRows {
  private final ExcelSheetRowSupport rowSupport;

  ExcelSheetRows(ExcelSheet sheet, ExcelSheetRowSupport rowSupport) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    this.rowSupport = Objects.requireNonNull(rowSupport, "rowSupport must not be null");
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  public ExcelSheetRows insertRows(int rowIndex, int rowCount) {
    rowSupport.insertRows(rowIndex, rowCount);
    return this;
  }

  /** Deletes the requested inclusive zero-based row band. */
  public ExcelSheetRows deleteRows(ExcelRowSpan rows) {
    rowSupport.deleteRows(rows);
    return this;
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  public ExcelSheetRows shiftRows(ExcelRowSpan rows, int delta) {
    rowSupport.shiftRows(rows, delta);
    return this;
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  public ExcelSheetRows setVisibility(ExcelRowSpan rows, boolean hidden) {
    rowSupport.setRowVisibility(rows, hidden);
    return this;
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
  public ExcelSheetRows group(ExcelRowSpan rows, boolean collapsed) {
    rowSupport.groupRows(rows, collapsed);
    return this;
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  public ExcelSheetRows ungroup(ExcelRowSpan rows) {
    rowSupport.ungroupRows(rows);
    return this;
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  public ExcelSheetRows setHeight(int firstRowIndex, int lastRowIndex, double heightPoints) {
    rowSupport.setRowHeight(firstRowIndex, lastRowIndex, heightPoints);
    return this;
  }

  /** Returns the number of physically stored rows in the sheet. */
  public int physicalCount() {
    return rowSupport.physicalRowCount();
  }

  /** Returns the 0-based index of the last row, or -1 if no rows exist. */
  public int lastIndex() {
    return rowSupport.lastRowIndex();
  }
}
