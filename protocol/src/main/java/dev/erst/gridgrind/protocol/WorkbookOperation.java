package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import java.util.List;
import java.util.Objects;

/** One validated workbook operation expressed in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WorkbookOperation.EnsureSheet.class, name = "ENSURE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.RenameSheet.class, name = "RENAME_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteSheet.class, name = "DELETE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.MoveSheet.class, name = "MOVE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.MergeCells.class, name = "MERGE_CELLS"),
  @JsonSubTypes.Type(value = WorkbookOperation.UnmergeCells.class, name = "UNMERGE_CELLS"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetColumnWidth.class, name = "SET_COLUMN_WIDTH"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetRowHeight.class, name = "SET_ROW_HEIGHT"),
  @JsonSubTypes.Type(value = WorkbookOperation.FreezePanes.class, name = "FREEZE_PANES"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetCell.class, name = "SET_CELL"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetRange.class, name = "SET_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearRange.class, name = "CLEAR_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetHyperlink.class, name = "SET_HYPERLINK"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearHyperlink.class, name = "CLEAR_HYPERLINK"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetComment.class, name = "SET_COMMENT"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearComment.class, name = "CLEAR_COMMENT"),
  @JsonSubTypes.Type(value = WorkbookOperation.ApplyStyle.class, name = "APPLY_STYLE"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetNamedRange.class, name = "SET_NAMED_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteNamedRange.class, name = "DELETE_NAMED_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.AppendRow.class, name = "APPEND_ROW"),
  @JsonSubTypes.Type(value = WorkbookOperation.AutoSizeColumns.class, name = "AUTO_SIZE_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.EvaluateFormulas.class, name = "EVALUATE_FORMULAS"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ForceFormulaRecalculationOnOpen.class,
      name = "FORCE_FORMULA_RECALCULATION_ON_OPEN")
})
public sealed interface WorkbookOperation {

  /** Ensures a sheet with the given name exists, creating it if absent. */
  record EnsureSheet(String sheetName) implements WorkbookOperation {
    public EnsureSheet {
      Validation.requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String sheetName, String newSheetName) implements WorkbookOperation {
    public RenameSheet {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(newSheetName, "newSheetName");
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet(String sheetName) implements WorkbookOperation {
    public DeleteSheet {
      Validation.requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(String sheetName, Integer targetIndex) implements WorkbookOperation {
    public MoveSheet {
      Validation.requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(targetIndex, "targetIndex must not be null");
      Validation.requireNonNegative(targetIndex, "targetIndex");
    }
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells(String sheetName, String range) implements WorkbookOperation {
    public MergeCells {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  record UnmergeCells(String sheetName, String range) implements WorkbookOperation {
    public UnmergeCells {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  record SetColumnWidth(
      String sheetName, Integer firstColumnIndex, Integer lastColumnIndex, Double widthCharacters)
      implements WorkbookOperation {
    public SetColumnWidth {
      Validation.requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(firstColumnIndex, "firstColumnIndex must not be null");
      Objects.requireNonNull(lastColumnIndex, "lastColumnIndex must not be null");
      Objects.requireNonNull(widthCharacters, "widthCharacters must not be null");
      Validation.requireNonNegative(firstColumnIndex, "firstColumnIndex");
      Validation.requireNonNegative(lastColumnIndex, "lastColumnIndex");
      Validation.requireOrderedSpan(
          firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");
      Validation.requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  record SetRowHeight(
      String sheetName, Integer firstRowIndex, Integer lastRowIndex, Double heightPoints)
      implements WorkbookOperation {
    public SetRowHeight {
      Validation.requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(firstRowIndex, "firstRowIndex must not be null");
      Objects.requireNonNull(lastRowIndex, "lastRowIndex must not be null");
      Objects.requireNonNull(heightPoints, "heightPoints must not be null");
      Validation.requireNonNegative(firstRowIndex, "firstRowIndex");
      Validation.requireNonNegative(lastRowIndex, "lastRowIndex");
      Validation.requireOrderedSpan(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
      Validation.requireRowHeightPoints(heightPoints);
    }
  }

  /** Freezes panes using explicit split and visible-origin coordinates. */
  record FreezePanes(
      String sheetName,
      Integer splitColumn,
      Integer splitRow,
      Integer leftmostColumn,
      Integer topRow)
      implements WorkbookOperation {
    public FreezePanes {
      Validation.requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(splitColumn, "splitColumn must not be null");
      Objects.requireNonNull(splitRow, "splitRow must not be null");
      Objects.requireNonNull(leftmostColumn, "leftmostColumn must not be null");
      Objects.requireNonNull(topRow, "topRow must not be null");
      Validation.requireNonNegative(splitColumn, "splitColumn");
      Validation.requireNonNegative(splitRow, "splitRow");
      Validation.requireNonNegative(leftmostColumn, "leftmostColumn");
      Validation.requireNonNegative(topRow, "topRow");
      Validation.requireFreezePaneCoordinates(splitColumn, splitRow, leftmostColumn, topRow);
    }
  }

  /** Sets a single cell to the given value. */
  record SetCell(String sheetName, String address, CellInput value) implements WorkbookOperation {
    public SetCell {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Sets a rectangular region of cells from a row-major grid of values. */
  record SetRange(String sheetName, String range, List<List<CellInput>> rows)
      implements WorkbookOperation {
    public SetRange {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      rows = Validation.copyRows(rows);
      Validation.requireRectangularRows(rows);
    }
  }

  /** Clears all cell values and styles within the specified range. */
  record ClearRange(String sheetName, String range) implements WorkbookOperation {
    public ClearRange {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(String sheetName, String address, HyperlinkTarget target)
      implements WorkbookOperation {
    public SetHyperlink {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink(String sheetName, String address) implements WorkbookOperation {
    public ClearHyperlink {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
    }
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(String sheetName, String address, CommentInput comment)
      implements WorkbookOperation {
    public SetComment {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment(String sheetName, String address) implements WorkbookOperation {
    public ClearComment {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
    }
  }

  /** Applies a style patch to every cell in the specified range. */
  record ApplyStyle(String sheetName, String range, CellStyleInput style)
      implements WorkbookOperation {
    public ApplyStyle {
      Validation.requireNonBlank(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Creates or replaces one typed named range in workbook or sheet scope. */
  record SetNamedRange(String name, NamedRangeScope scope, NamedRangeTarget target)
      implements WorkbookOperation {
    public SetNamedRange {
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Validation.requireNamedRangeName(name);
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  record DeleteNamedRange(String name, NamedRangeScope scope) implements WorkbookOperation {
    public DeleteNamedRange {
      Objects.requireNonNull(scope, "scope must not be null");
      Validation.requireNamedRangeName(name);
    }
  }

  /** Appends a new row of values after the last occupied row on the sheet. */
  record AppendRow(String sheetName, List<CellInput> values) implements WorkbookOperation {
    public AppendRow {
      Validation.requireNonBlank(sheetName, "sheetName");
      values = values == null ? List.of() : List.copyOf(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (CellInput item : values) {
        Objects.requireNonNull(item, "values must not contain nulls");
      }
    }
  }

  /** Auto-sizes all populated columns on the sheet to fit their content. */
  record AutoSizeColumns(String sheetName) implements WorkbookOperation {
    public AutoSizeColumns {
      Validation.requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Evaluates all formulas in the workbook at operation time. */
  record EvaluateFormulas() implements WorkbookOperation {}

  /** Marks the workbook so that Excel recalculates all formulas on next open. */
  record ForceFormulaRecalculationOnOpen() implements WorkbookOperation {}

  /** Returns the SCREAMING_SNAKE_CASE type name of this operation as used in the wire protocol. */
  default String operationType() {
    return switch (this) {
      case EnsureSheet _ -> "ENSURE_SHEET";
      case RenameSheet _ -> "RENAME_SHEET";
      case DeleteSheet _ -> "DELETE_SHEET";
      case MoveSheet _ -> "MOVE_SHEET";
      case MergeCells _ -> "MERGE_CELLS";
      case UnmergeCells _ -> "UNMERGE_CELLS";
      case SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case SetRowHeight _ -> "SET_ROW_HEIGHT";
      case FreezePanes _ -> "FREEZE_PANES";
      case SetCell _ -> "SET_CELL";
      case SetRange _ -> "SET_RANGE";
      case ClearRange _ -> "CLEAR_RANGE";
      case SetHyperlink _ -> "SET_HYPERLINK";
      case ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case SetComment _ -> "SET_COMMENT";
      case ClearComment _ -> "CLEAR_COMMENT";
      case ApplyStyle _ -> "APPLY_STYLE";
      case SetNamedRange _ -> "SET_NAMED_RANGE";
      case DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case AppendRow _ -> "APPEND_ROW";
      case AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case EvaluateFormulas _ -> "EVALUATE_FORMULAS";
      case ForceFormulaRecalculationOnOpen _ -> "FORCE_FORMULA_RECALCULATION_ON_OPEN";
    };
  }

  /** Shared validation helpers for WorkbookOperation compact constructors. */
  final class Validation {
    private Validation() {}

    static void requireNonBlank(String value, String fieldName) {
      Objects.requireNonNull(value, fieldName + " must not be null");
      if (value.isBlank()) {
        throw new IllegalArgumentException(fieldName + " must not be blank");
      }
    }

    static void requireNonNegative(int value, String fieldName) {
      if (value < 0) {
        throw new IllegalArgumentException(fieldName + " must not be negative");
      }
    }

    static void requireOrderedSpan(
        int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
      if (lastValue < firstValue) {
        throw new IllegalArgumentException(
            lastFieldName + " must not be less than " + firstFieldName);
      }
    }

    static void requireColumnWidthCharacters(double widthCharacters) {
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

    static void requireRowHeightPoints(double heightPoints) {
      requireFinitePositive(heightPoints, "heightPoints");
      if (Math.round(heightPoints * 20.0d) > Short.MAX_VALUE) {
        throw new IllegalArgumentException(
            "heightPoints is too large for Excel row height storage: " + heightPoints);
      }
      if (Math.round(heightPoints * 20.0d) <= 0) {
        throw new IllegalArgumentException(
            "heightPoints is too small to produce a visible Excel row height: " + heightPoints);
      }
    }

    static void requireFreezePaneCoordinates(
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

    static void requireNamedRangeName(String name) {
      ExcelNamedRangeDefinition.validateName(name);
    }

    static void requireFinitePositive(double value, String fieldName) {
      if (!Double.isFinite(value)) {
        throw new IllegalArgumentException(fieldName + " must be finite");
      }
      if (value <= 0.0d) {
        throw new IllegalArgumentException(fieldName + " must be greater than 0");
      }
    }

    static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
      if (rows == null) {
        return List.of();
      }
      return rows.stream().map(row -> row == null ? null : List.copyOf(row)).toList();
    }

    static void requireRectangularRows(List<List<CellInput>> rows) {
      if (rows.isEmpty()) {
        throw new IllegalArgumentException("rows must not be empty");
      }
      int expectedWidth = -1;
      for (List<CellInput> row : rows) {
        Objects.requireNonNull(row, "rows must not contain null rows");
        if (row.isEmpty()) {
          throw new IllegalArgumentException("rows must not contain empty rows");
        }
        if (expectedWidth < 0) {
          expectedWidth = row.size();
        } else if (row.size() != expectedWidth) {
          throw new IllegalArgumentException("rows must describe a rectangular matrix");
        }
        for (CellInput value : row) {
          Objects.requireNonNull(value, "rows must not contain null cell values");
        }
      }
    }
  }
}
