package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import java.util.Objects;

/** Column insertion, movement, visibility, grouping, and sizing operations for one sheet. */
public final class ExcelSheetColumns {
  private final ExcelSheet sheet;
  private final ExcelSheetColumnSupport columnSupport;

  ExcelSheetColumns(ExcelSheet sheet, ExcelSheetColumnSupport columnSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.columnSupport = Objects.requireNonNull(columnSupport, "columnSupport must not be null");
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  public ExcelSheetColumns setWidth(
      int firstColumnIndex, int lastColumnIndex, double widthCharacters) {
    columnSupport.setColumnWidth(firstColumnIndex, lastColumnIndex, widthCharacters);
    return this;
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  public ExcelSheetColumns insert(int columnIndex, int columnCount) {
    columnSupport.insertColumns(columnIndex, columnCount);
    return this;
  }

  /** Deletes the requested inclusive zero-based column band. */
  public ExcelSheetColumns delete(ExcelColumnSpan columns) {
    columnSupport.deleteColumns(columns);
    return this;
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  public ExcelSheetColumns shift(ExcelColumnSpan columns, int delta) {
    columnSupport.shiftColumns(columns, delta);
    return this;
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  public ExcelSheetColumns setVisibility(ExcelColumnSpan columns, boolean hidden) {
    columnSupport.setColumnVisibility(columns, hidden);
    return this;
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  public ExcelSheetColumns group(ExcelColumnSpan columns, boolean collapsed) {
    columnSupport.groupColumns(columns, collapsed);
    return this;
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  public ExcelSheetColumns ungroup(ExcelColumnSpan columns) {
    columnSupport.ungroupColumns(columns);
    return this;
  }

  /** Auto-sizes all populated columns on this sheet to fit their content. */
  public ExcelSheetColumns autoSize() {
    columnSupport.autoSizeColumns(sheet.name());
    return this;
  }

  /**
   * Returns the 0-based index of the widest column across all rows, or -1 if the sheet is empty.
   */
  public int lastIndex() {
    return columnSupport.lastColumnIndex();
  }
}
