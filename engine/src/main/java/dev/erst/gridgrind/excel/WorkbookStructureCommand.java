package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import java.util.Objects;

/** Structural sheet commands that change geometry, grouping, and visibility bands. */
public sealed interface WorkbookStructureCommand extends WorkbookCommand
    permits WorkbookStructureCommand.MergeCells,
        WorkbookStructureCommand.UnmergeCells,
        WorkbookStructureCommand.SetColumnWidth,
        WorkbookStructureCommand.SetRowHeight,
        WorkbookStructureCommand.InsertRows,
        WorkbookStructureCommand.DeleteRows,
        WorkbookStructureCommand.ShiftRows,
        WorkbookStructureCommand.InsertColumns,
        WorkbookStructureCommand.DeleteColumns,
        WorkbookStructureCommand.ShiftColumns,
        WorkbookStructureCommand.SetRowVisibility,
        WorkbookStructureCommand.SetColumnVisibility,
        WorkbookStructureCommand.GroupRows,
        WorkbookStructureCommand.UngroupRows,
        WorkbookStructureCommand.GroupColumns,
        WorkbookStructureCommand.UngroupColumns {

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells(String sheetName, String range) implements WorkbookStructureCommand {
    public MergeCells {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  record UnmergeCells(String sheetName, String range) implements WorkbookStructureCommand {
    public UnmergeCells {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  record SetColumnWidth(
      String sheetName, int firstColumnIndex, int lastColumnIndex, double widthCharacters)
      implements WorkbookStructureCommand {
    public SetColumnWidth {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (firstColumnIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("firstColumnIndex", firstColumnIndex));
      }
      if (lastColumnIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("lastColumnIndex", lastColumnIndex));
      }
      if (lastColumnIndex < firstColumnIndex) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeLessThan(
                "lastColumnIndex", lastColumnIndex, "firstColumnIndex", firstColumnIndex));
      }
      requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  record SetRowHeight(String sheetName, int firstRowIndex, int lastRowIndex, double heightPoints)
      implements WorkbookStructureCommand {
    public SetRowHeight {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (firstRowIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("firstRowIndex", firstRowIndex));
      }
      if (lastRowIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("lastRowIndex", lastRowIndex));
      }
      if (lastRowIndex < firstRowIndex) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeLessThan(
                "lastRowIndex", lastRowIndex, "firstRowIndex", firstRowIndex));
      }
      requireRowHeightPoints(heightPoints);
    }
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  record InsertRows(String sheetName, int rowIndex, int rowCount)
      implements WorkbookStructureCommand {
    public InsertRows {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (rowIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("rowIndex", rowIndex));
      }
      if (rowIndex > ExcelRowSpan.MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotExceed("rowIndex", rowIndex, ExcelRowSpan.MAX_ROW_INDEX));
      }
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
    }
  }

  /** Deletes the requested inclusive zero-based row band. */
  record DeleteRows(String sheetName, ExcelRowSpan rows) implements WorkbookStructureCommand {
    public DeleteRows {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  record ShiftRows(String sheetName, ExcelRowSpan rows, int delta)
      implements WorkbookStructureCommand {
    public ShiftRows {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (delta == 0) {
        throw new IllegalArgumentException("delta must not be 0");
      }
    }
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  record InsertColumns(String sheetName, int columnIndex, int columnCount)
      implements WorkbookStructureCommand {
    public InsertColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (columnIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("columnIndex", columnIndex));
      }
      if (columnIndex > ExcelColumnSpan.MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotExceed(
                "columnIndex", columnIndex, ExcelColumnSpan.MAX_COLUMN_INDEX));
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
    }
  }

  /** Deletes the requested inclusive zero-based column band. */
  record DeleteColumns(String sheetName, ExcelColumnSpan columns)
      implements WorkbookStructureCommand {
    public DeleteColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(columns, "columns must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  record ShiftColumns(String sheetName, ExcelColumnSpan columns, int delta)
      implements WorkbookStructureCommand {
    public ShiftColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(columns, "columns must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (delta == 0) {
        throw new IllegalArgumentException("delta must not be 0");
      }
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  record SetRowVisibility(String sheetName, ExcelRowSpan rows, boolean hidden)
      implements WorkbookStructureCommand {
    public SetRowVisibility {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  record SetColumnVisibility(String sheetName, ExcelColumnSpan columns, boolean hidden)
      implements WorkbookStructureCommand {
    public SetColumnVisibility {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(columns, "columns must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
  record GroupRows(String sheetName, ExcelRowSpan rows, boolean collapsed)
      implements WorkbookStructureCommand {
    public GroupRows {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  record UngroupRows(String sheetName, ExcelRowSpan rows) implements WorkbookStructureCommand {
    public UngroupRows {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  record GroupColumns(String sheetName, ExcelColumnSpan columns, boolean collapsed)
      implements WorkbookStructureCommand {
    public GroupColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(columns, "columns must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  record UngroupColumns(String sheetName, ExcelColumnSpan columns)
      implements WorkbookStructureCommand {
    public UngroupColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(columns, "columns must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  private static void requireColumnWidthCharacters(double widthCharacters) {
    ExcelSheetLayoutLimits.requireColumnWidthCharacters(widthCharacters, "widthCharacters");
  }

  private static void requireRowHeightPoints(double heightPoints) {
    ExcelSheetLayoutLimits.requireRowHeightPoints(heightPoints, "heightPoints");
  }
}
