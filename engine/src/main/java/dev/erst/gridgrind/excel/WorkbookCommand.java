package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/**
 * Workbook-core commands that can be executed deterministically against an {@link ExcelWorkbook}.
 */
public sealed interface WorkbookCommand
    permits WorkbookCommand.CreateSheet,
        WorkbookCommand.RenameSheet,
        WorkbookCommand.DeleteSheet,
        WorkbookCommand.MoveSheet,
        WorkbookCommand.MergeCells,
        WorkbookCommand.UnmergeCells,
        WorkbookCommand.SetColumnWidth,
        WorkbookCommand.SetRowHeight,
        WorkbookCommand.FreezePanes,
        WorkbookCommand.SetCell,
        WorkbookCommand.SetRange,
        WorkbookCommand.ClearRange,
        WorkbookCommand.SetHyperlink,
        WorkbookCommand.ClearHyperlink,
        WorkbookCommand.SetComment,
        WorkbookCommand.ClearComment,
        WorkbookCommand.ApplyStyle,
        WorkbookCommand.SetNamedRange,
        WorkbookCommand.DeleteNamedRange,
        WorkbookCommand.AppendRow,
        WorkbookCommand.AutoSizeColumns,
        WorkbookCommand.EvaluateAllFormulas,
        WorkbookCommand.ForceFormulaRecalculationOnOpen {

  record CreateSheet(String sheetName) implements WorkbookCommand {
    public CreateSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String sheetName, String newSheetName) implements WorkbookCommand {
    public RenameSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(newSheetName, "newSheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (newSheetName.isBlank()) {
        throw new IllegalArgumentException("newSheetName must not be blank");
      }
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet(String sheetName) implements WorkbookCommand {
    public DeleteSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(String sheetName, int targetIndex) implements WorkbookCommand {
    public MoveSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (targetIndex < 0) {
        throw new IllegalArgumentException("targetIndex must not be negative");
      }
    }
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells(String sheetName, String range) implements WorkbookCommand {
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
  record UnmergeCells(String sheetName, String range) implements WorkbookCommand {
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
      implements WorkbookCommand {
    public SetColumnWidth {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (firstColumnIndex < 0) {
        throw new IllegalArgumentException("firstColumnIndex must not be negative");
      }
      if (lastColumnIndex < 0) {
        throw new IllegalArgumentException("lastColumnIndex must not be negative");
      }
      if (lastColumnIndex < firstColumnIndex) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not be less than firstColumnIndex");
      }
      requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  record SetRowHeight(String sheetName, int firstRowIndex, int lastRowIndex, double heightPoints)
      implements WorkbookCommand {
    public SetRowHeight {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (firstRowIndex < 0) {
        throw new IllegalArgumentException("firstRowIndex must not be negative");
      }
      if (lastRowIndex < 0) {
        throw new IllegalArgumentException("lastRowIndex must not be negative");
      }
      if (lastRowIndex < firstRowIndex) {
        throw new IllegalArgumentException("lastRowIndex must not be less than firstRowIndex");
      }
      requireRowHeightPoints(heightPoints);
    }
  }

  /** Freezes panes using explicit split and visible-origin coordinates. */
  record FreezePanes(
      String sheetName, int splitColumn, int splitRow, int leftmostColumn, int topRow)
      implements WorkbookCommand {
    public FreezePanes {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (splitColumn < 0) {
        throw new IllegalArgumentException("splitColumn must not be negative");
      }
      if (splitRow < 0) {
        throw new IllegalArgumentException("splitRow must not be negative");
      }
      if (leftmostColumn < 0) {
        throw new IllegalArgumentException("leftmostColumn must not be negative");
      }
      if (topRow < 0) {
        throw new IllegalArgumentException("topRow must not be negative");
      }
      requireFreezePaneCoordinates(splitColumn, splitRow, leftmostColumn, topRow);
    }
  }

  record SetCell(String sheetName, String address, ExcelCellValue value)
      implements WorkbookCommand {
    public SetCell {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(value, "value must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  record SetRange(String sheetName, String range, List<List<ExcelCellValue>> rows)
      implements WorkbookCommand {
    public SetRange {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
      rows = List.copyOf(rows);
      if (rows.isEmpty()) {
        throw new IllegalArgumentException("rows must not be empty");
      }
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
      }
    }
  }

  record ClearRange(String sheetName, String range) implements WorkbookCommand {
    public ClearRange {
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

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(String sheetName, String address, ExcelHyperlink target)
      implements WorkbookCommand {
    public SetHyperlink {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(target, "target must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink(String sheetName, String address) implements WorkbookCommand {
    public ClearHyperlink {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(String sheetName, String address, ExcelComment comment)
      implements WorkbookCommand {
    public SetComment {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(comment, "comment must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment(String sheetName, String address) implements WorkbookCommand {
    public ClearComment {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  record ApplyStyle(String sheetName, String range, ExcelCellStyle style)
      implements WorkbookCommand {
    public ApplyStyle {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(style, "style must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Creates or replaces one named range in workbook or sheet scope. */
  record SetNamedRange(ExcelNamedRangeDefinition definition) implements WorkbookCommand {
    public SetNamedRange {
      Objects.requireNonNull(definition, "definition must not be null");
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  record DeleteNamedRange(String name, ExcelNamedRangeScope scope) implements WorkbookCommand {
    public DeleteNamedRange {
      name = ExcelNamedRangeDefinition.validateName(name);
      Objects.requireNonNull(scope, "scope must not be null");
    }
  }

  record AppendRow(String sheetName, List<ExcelCellValue> values) implements WorkbookCommand {
    public AppendRow {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(values, "values must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      values = List.copyOf(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (ExcelCellValue value : values) {
        Objects.requireNonNull(value, "value must not be null");
      }
    }
  }

  /** Auto-sizes all populated columns on the named sheet. */
  record AutoSizeColumns(String sheetName) implements WorkbookCommand {
    public AutoSizeColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  record EvaluateAllFormulas() implements WorkbookCommand {}

  record ForceFormulaRecalculationOnOpen() implements WorkbookCommand {}

  private static void requireColumnWidthCharacters(double widthCharacters) {
    requireFinitePositive(widthCharacters, "widthCharacters");
    if (widthCharacters > 255.0d) {
      throw new IllegalArgumentException(
          "widthCharacters must be less than or equal to 255.0: " + widthCharacters);
    }
    if (Math.round(widthCharacters * 256.0d) <= 0) {
      throw new IllegalArgumentException(
          "widthCharacters is too small to produce a visible Excel column width: "
              + widthCharacters);
    }
  }

  private static void requireRowHeightPoints(double heightPoints) {
    requireFinitePositive(heightPoints, "heightPoints");
    if (heightPoints > Short.MAX_VALUE / 20.0d) {
      throw new IllegalArgumentException(
          "heightPoints is too large for Excel row height storage: " + heightPoints);
    }
    if ((long) (heightPoints * 20.0d) <= 0) {
      throw new IllegalArgumentException(
          "heightPoints is too small to produce a visible Excel row height: " + heightPoints);
    }
  }

  private static void requireFreezePaneCoordinates(
      int splitColumn, int splitRow, int leftmostColumn, int topRow) {
    if (splitColumn == 0 && splitRow == 0) {
      throw new IllegalArgumentException("splitColumn and splitRow must not both be 0");
    }
    if (splitColumn == 0 && leftmostColumn != 0) {
      throw new IllegalArgumentException(
          "leftmostColumn must be 0 when splitColumn is 0: " + leftmostColumn);
    }
    if (splitRow == 0 && topRow != 0) {
      throw new IllegalArgumentException("topRow must be 0 when splitRow is 0: " + topRow);
    }
    if (splitColumn > 0 && leftmostColumn < splitColumn) {
      throw new IllegalArgumentException(
          "leftmostColumn must be greater than or equal to splitColumn");
    }
    if (splitRow > 0 && topRow < splitRow) {
      throw new IllegalArgumentException("topRow must be greater than or equal to splitRow");
    }
  }

  private static void requireFinitePositive(double value, String fieldName) {
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    if (value <= 0.0d) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
  }
}
