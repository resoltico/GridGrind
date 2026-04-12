package dev.erst.gridgrind.protocol.operation;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelColumnSpan;
import dev.erst.gridgrind.excel.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelSheetNames;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.dto.ProtocolDefinedNameValidation;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** One validated workbook operation expressed in protocol form. */
@SuppressWarnings("PMD.ExcessivePublicCount")
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WorkbookOperation.EnsureSheet.class, name = "ENSURE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.RenameSheet.class, name = "RENAME_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteSheet.class, name = "DELETE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.MoveSheet.class, name = "MOVE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.CopySheet.class, name = "COPY_SHEET"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetActiveSheet.class, name = "SET_ACTIVE_SHEET"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetSelectedSheets.class,
      name = "SET_SELECTED_SHEETS"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetSheetVisibility.class,
      name = "SET_SHEET_VISIBILITY"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetSheetProtection.class,
      name = "SET_SHEET_PROTECTION"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ClearSheetProtection.class,
      name = "CLEAR_SHEET_PROTECTION"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetWorkbookProtection.class,
      name = "SET_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ClearWorkbookProtection.class,
      name = "CLEAR_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(value = WorkbookOperation.MergeCells.class, name = "MERGE_CELLS"),
  @JsonSubTypes.Type(value = WorkbookOperation.UnmergeCells.class, name = "UNMERGE_CELLS"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetColumnWidth.class, name = "SET_COLUMN_WIDTH"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetRowHeight.class, name = "SET_ROW_HEIGHT"),
  @JsonSubTypes.Type(value = WorkbookOperation.InsertRows.class, name = "INSERT_ROWS"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteRows.class, name = "DELETE_ROWS"),
  @JsonSubTypes.Type(value = WorkbookOperation.ShiftRows.class, name = "SHIFT_ROWS"),
  @JsonSubTypes.Type(value = WorkbookOperation.InsertColumns.class, name = "INSERT_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteColumns.class, name = "DELETE_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.ShiftColumns.class, name = "SHIFT_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetRowVisibility.class, name = "SET_ROW_VISIBILITY"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetColumnVisibility.class,
      name = "SET_COLUMN_VISIBILITY"),
  @JsonSubTypes.Type(value = WorkbookOperation.GroupRows.class, name = "GROUP_ROWS"),
  @JsonSubTypes.Type(value = WorkbookOperation.UngroupRows.class, name = "UNGROUP_ROWS"),
  @JsonSubTypes.Type(value = WorkbookOperation.GroupColumns.class, name = "GROUP_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.UngroupColumns.class, name = "UNGROUP_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetSheetPane.class, name = "SET_SHEET_PANE"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetSheetZoom.class, name = "SET_SHEET_ZOOM"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetPrintLayout.class, name = "SET_PRINT_LAYOUT"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearPrintLayout.class, name = "CLEAR_PRINT_LAYOUT"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetCell.class, name = "SET_CELL"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetRange.class, name = "SET_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearRange.class, name = "CLEAR_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetHyperlink.class, name = "SET_HYPERLINK"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearHyperlink.class, name = "CLEAR_HYPERLINK"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetComment.class, name = "SET_COMMENT"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearComment.class, name = "CLEAR_COMMENT"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetPicture.class, name = "SET_PICTURE"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetChart.class, name = "SET_CHART"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetShape.class, name = "SET_SHAPE"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetEmbeddedObject.class,
      name = "SET_EMBEDDED_OBJECT"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetDrawingObjectAnchor.class,
      name = "SET_DRAWING_OBJECT_ANCHOR"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.DeleteDrawingObject.class,
      name = "DELETE_DRAWING_OBJECT"),
  @JsonSubTypes.Type(value = WorkbookOperation.ApplyStyle.class, name = "APPLY_STYLE"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetDataValidation.class,
      name = "SET_DATA_VALIDATION"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ClearDataValidations.class,
      name = "CLEAR_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.SetConditionalFormatting.class,
      name = "SET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ClearConditionalFormatting.class,
      name = "CLEAR_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetAutofilter.class, name = "SET_AUTOFILTER"),
  @JsonSubTypes.Type(value = WorkbookOperation.ClearAutofilter.class, name = "CLEAR_AUTOFILTER"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetTable.class, name = "SET_TABLE"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteTable.class, name = "DELETE_TABLE"),
  @JsonSubTypes.Type(value = WorkbookOperation.SetNamedRange.class, name = "SET_NAMED_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.DeleteNamedRange.class, name = "DELETE_NAMED_RANGE"),
  @JsonSubTypes.Type(value = WorkbookOperation.AppendRow.class, name = "APPEND_ROW"),
  @JsonSubTypes.Type(value = WorkbookOperation.AutoSizeColumns.class, name = "AUTO_SIZE_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookOperation.EvaluateFormulas.class, name = "EVALUATE_FORMULAS"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.EvaluateFormulaCells.class,
      name = "EVALUATE_FORMULA_CELLS"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ClearFormulaCaches.class,
      name = "CLEAR_FORMULA_CACHES"),
  @JsonSubTypes.Type(
      value = WorkbookOperation.ForceFormulaRecalculationOnOpen.class,
      name = "FORCE_FORMULA_RECALCULATION_ON_OPEN")
})
public sealed interface WorkbookOperation {

  /** Ensures a sheet with the given name exists, creating it if absent. */
  record EnsureSheet(String sheetName) implements WorkbookOperation {
    public EnsureSheet {
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String sheetName, String newSheetName) implements WorkbookOperation {
    public RenameSheet {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireSheetName(newSheetName, "newSheetName");
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet(String sheetName) implements WorkbookOperation {
    public DeleteSheet {
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(String sheetName, Integer targetIndex) implements WorkbookOperation {
    public MoveSheet {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(targetIndex, "targetIndex must not be null");
      Validation.requireNonNegative(targetIndex, "targetIndex");
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  record CopySheet(String sourceSheetName, String newSheetName, SheetCopyPosition position)
      implements WorkbookOperation {
    public CopySheet {
      Validation.requireSheetName(sourceSheetName, "sourceSheetName");
      Validation.requireSheetName(newSheetName, "newSheetName");
      position = position == null ? new SheetCopyPosition.AppendAtEnd() : position;
    }
  }

  /** Sets the active sheet and ensures it is selected. */
  record SetActiveSheet(String sheetName) implements WorkbookOperation {
    public SetActiveSheet {
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Sets the selected visible sheet set. */
  record SetSelectedSheets(List<String> sheetNames) implements WorkbookOperation {
    public SetSelectedSheets {
      sheetNames = Validation.copySheetNames(sheetNames, "sheetNames");
      if (sheetNames.isEmpty()) {
        throw new IllegalArgumentException("sheetNames must not be empty");
      }
      Validation.requireDistinct(sheetNames, "sheetNames");
    }
  }

  /** Sets one sheet visibility. */
  record SetSheetVisibility(String sheetName, ExcelSheetVisibility visibility)
      implements WorkbookOperation {
    public SetSheetVisibility {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(visibility, "visibility must not be null");
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  record SetSheetProtection(String sheetName, SheetProtectionSettings protection, String password)
      implements WorkbookOperation {
    /** Enables sheet protection without applying a password hash. */
    public SetSheetProtection(String sheetName, SheetProtectionSettings protection) {
      this(sheetName, protection, null);
    }

    public SetSheetProtection {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(protection, "protection must not be null");
      if (password != null && password.isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }
  }

  /** Disables sheet protection entirely. */
  record ClearSheetProtection(String sheetName) implements WorkbookOperation {
    public ClearSheetProtection {
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  record SetWorkbookProtection(WorkbookProtectionInput protection) implements WorkbookOperation {
    public SetWorkbookProtection {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Clears workbook-level protection and password hashes entirely. */
  record ClearWorkbookProtection() implements WorkbookOperation {}

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells(String sheetName, String range) implements WorkbookOperation {
    public MergeCells {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  record UnmergeCells(String sheetName, String range) implements WorkbookOperation {
    public UnmergeCells {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  record SetColumnWidth(
      String sheetName, Integer firstColumnIndex, Integer lastColumnIndex, Double widthCharacters)
      implements WorkbookOperation {
    public SetColumnWidth {
      Validation.requireSheetName(sheetName, "sheetName");
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
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(firstRowIndex, "firstRowIndex must not be null");
      Objects.requireNonNull(lastRowIndex, "lastRowIndex must not be null");
      Objects.requireNonNull(heightPoints, "heightPoints must not be null");
      Validation.requireNonNegative(firstRowIndex, "firstRowIndex");
      Validation.requireNonNegative(lastRowIndex, "lastRowIndex");
      Validation.requireOrderedSpan(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
      Validation.requireRowHeightPoints(heightPoints);
    }
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  record InsertRows(String sheetName, Integer rowIndex, Integer rowCount)
      implements WorkbookOperation {
    public InsertRows {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(rowIndex, "rowIndex must not be null");
      Objects.requireNonNull(rowCount, "rowCount must not be null");
      Validation.requireRowIndex(rowIndex, "rowIndex");
      Validation.requirePositive(rowCount, "rowCount");
    }
  }

  /** Deletes the requested inclusive zero-based row band. */
  record DeleteRows(String sheetName, RowSpanInput rows) implements WorkbookOperation {
    public DeleteRows {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(rows, "rows must not be null");
    }
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  record ShiftRows(String sheetName, RowSpanInput rows, Integer delta)
      implements WorkbookOperation {
    public ShiftRows {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(rows, "rows must not be null");
      Objects.requireNonNull(delta, "delta must not be null");
      Validation.requireNonZero(delta, "delta");
    }
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  record InsertColumns(String sheetName, Integer columnIndex, Integer columnCount)
      implements WorkbookOperation {
    public InsertColumns {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(columnIndex, "columnIndex must not be null");
      Objects.requireNonNull(columnCount, "columnCount must not be null");
      Validation.requireColumnIndex(columnIndex, "columnIndex");
      Validation.requirePositive(columnCount, "columnCount");
    }
  }

  /** Deletes the requested inclusive zero-based column band. */
  record DeleteColumns(String sheetName, ColumnSpanInput columns) implements WorkbookOperation {
    public DeleteColumns {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(columns, "columns must not be null");
    }
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  record ShiftColumns(String sheetName, ColumnSpanInput columns, Integer delta)
      implements WorkbookOperation {
    public ShiftColumns {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(columns, "columns must not be null");
      Objects.requireNonNull(delta, "delta must not be null");
      Validation.requireNonZero(delta, "delta");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  record SetRowVisibility(String sheetName, RowSpanInput rows, Boolean hidden)
      implements WorkbookOperation {
    public SetRowVisibility {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(rows, "rows must not be null");
      Objects.requireNonNull(hidden, "hidden must not be null");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  record SetColumnVisibility(String sheetName, ColumnSpanInput columns, Boolean hidden)
      implements WorkbookOperation {
    public SetColumnVisibility {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(columns, "columns must not be null");
      Objects.requireNonNull(hidden, "hidden must not be null");
    }
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
  record GroupRows(String sheetName, RowSpanInput rows, Boolean collapsed)
      implements WorkbookOperation {
    public GroupRows {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(rows, "rows must not be null");
      collapsed = collapsed == null ? Boolean.FALSE : collapsed;
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  record UngroupRows(String sheetName, RowSpanInput rows) implements WorkbookOperation {
    public UngroupRows {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(rows, "rows must not be null");
    }
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  record GroupColumns(String sheetName, ColumnSpanInput columns, Boolean collapsed)
      implements WorkbookOperation {
    public GroupColumns {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(columns, "columns must not be null");
      collapsed = collapsed == null ? Boolean.FALSE : collapsed;
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  record UngroupColumns(String sheetName, ColumnSpanInput columns) implements WorkbookOperation {
    public UngroupColumns {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(columns, "columns must not be null");
    }
  }

  /** Applies one explicit pane state to a sheet. */
  record SetSheetPane(String sheetName, PaneInput pane) implements WorkbookOperation {
    public SetSheetPane {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(pane, "pane must not be null");
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  record SetSheetZoom(String sheetName, Integer zoomPercent) implements WorkbookOperation {
    public SetSheetZoom {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(zoomPercent, "zoomPercent must not be null");
      Validation.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  record SetPrintLayout(String sheetName, PrintLayoutInput printLayout)
      implements WorkbookOperation {
    public SetPrintLayout {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  record ClearPrintLayout(String sheetName) implements WorkbookOperation {
    public ClearPrintLayout {
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Sets a single cell to the given value. */
  record SetCell(String sheetName, String address, CellInput value) implements WorkbookOperation {
    public SetCell {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Sets a rectangular region of cells from a row-major grid of values. */
  record SetRange(String sheetName, String range, List<List<CellInput>> rows)
      implements WorkbookOperation {
    public SetRange {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      rows = Validation.copyRows(rows);
      Validation.requireRectangularRows(rows);
    }
  }

  /** Clears all cell values and styles within the specified range. */
  record ClearRange(String sheetName, String range) implements WorkbookOperation {
    public ClearRange {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
    }
  }

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(String sheetName, String address, HyperlinkTarget target)
      implements WorkbookOperation {
    public SetHyperlink {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink(String sheetName, String address) implements WorkbookOperation {
    public ClearHyperlink {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
    }
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(String sheetName, String address, CommentInput comment)
      implements WorkbookOperation {
    public SetComment {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment(String sheetName, String address) implements WorkbookOperation {
    public ClearComment {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(address, "address");
    }
  }

  /** Creates or replaces one picture-backed drawing object on one sheet. */
  record SetPicture(String sheetName, PictureInput picture) implements WorkbookOperation {
    public SetPicture {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(picture, "picture must not be null");
    }
  }

  /** Creates or mutates one supported simple chart on one sheet. */
  record SetChart(String sheetName, ChartInput chart) implements WorkbookOperation {
    public SetChart {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(chart, "chart must not be null");
    }
  }

  /** Creates or replaces one simple-shape or connector drawing object on one sheet. */
  record SetShape(String sheetName, ShapeInput shape) implements WorkbookOperation {
    public SetShape {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(shape, "shape must not be null");
    }
  }

  /** Creates or replaces one embedded-object drawing object on one sheet. */
  record SetEmbeddedObject(String sheetName, EmbeddedObjectInput embeddedObject)
      implements WorkbookOperation {
    public SetEmbeddedObject {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(embeddedObject, "embeddedObject must not be null");
    }
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  record SetDrawingObjectAnchor(String sheetName, String objectName, DrawingAnchorInput anchor)
      implements WorkbookOperation {
    public SetDrawingObjectAnchor {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(objectName, "objectName");
      Objects.requireNonNull(anchor, "anchor must not be null");
    }
  }

  /** Deletes one existing drawing object by sheet-local name. */
  record DeleteDrawingObject(String sheetName, String objectName) implements WorkbookOperation {
    public DeleteDrawingObject {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(objectName, "objectName");
    }
  }

  /** Applies a style patch to every cell in the specified range. */
  record ApplyStyle(String sheetName, String range, CellStyleInput style)
      implements WorkbookOperation {
    public ApplyStyle {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  record SetDataValidation(String sheetName, String range, DataValidationInput validation)
      implements WorkbookOperation {
    public SetDataValidation {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  record ClearDataValidations(String sheetName, RangeSelection selection)
      implements WorkbookOperation {
    public ClearDataValidations {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Creates or replaces one logical conditional-formatting block over the requested ranges. */
  record SetConditionalFormatting(
      String sheetName, ConditionalFormattingBlockInput conditionalFormatting)
      implements WorkbookOperation {
    public SetConditionalFormatting {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(conditionalFormatting, "conditionalFormatting must not be null");
    }
  }

  /** Removes conditional-formatting blocks on the sheet that match the provided range selection. */
  record ClearConditionalFormatting(String sheetName, RangeSelection selection)
      implements WorkbookOperation {
    public ClearConditionalFormatting {
      Validation.requireSheetName(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Creates or replaces one sheet-level autofilter range. */
  record SetAutofilter(
      String sheetName,
      String range,
      List<AutofilterFilterColumnInput> criteria,
      AutofilterSortStateInput sortState)
      implements WorkbookOperation {
    /** Creates a plain sheet-level autofilter without criteria or explicit sort state. */
    public SetAutofilter(String sheetName, String range) {
      this(sheetName, range, List.of(), null);
    }

    public SetAutofilter {
      Validation.requireSheetName(sheetName, "sheetName");
      Validation.requireNonBlank(range, "range");
      criteria = criteria == null ? List.of() : List.copyOf(criteria);
      for (AutofilterFilterColumnInput criterion : criteria) {
        Objects.requireNonNull(criterion, "criteria must not contain null values");
      }
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  record ClearAutofilter(String sheetName) implements WorkbookOperation {
    public ClearAutofilter {
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Creates or replaces one workbook-global table definition. */
  record SetTable(TableInput table) implements WorkbookOperation {
    public SetTable {
      Objects.requireNonNull(table, "table must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  record DeleteTable(String name, String sheetName) implements WorkbookOperation {
    public DeleteTable {
      Validation.requireTableName(name);
      Validation.requireSheetName(sheetName, "sheetName");
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
      Validation.requireSheetName(sheetName, "sheetName");
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
      Validation.requireSheetName(sheetName, "sheetName");
    }
  }

  /** Evaluates all formulas in the workbook at operation time. */
  record EvaluateFormulas() implements WorkbookOperation {}

  /** Evaluates one or more concrete formula cells and stores their cached results. */
  record EvaluateFormulaCells(List<FormulaCellTargetInput> cells) implements WorkbookOperation {
    public EvaluateFormulaCells {
      Objects.requireNonNull(cells, "cells must not be null");
      cells = List.copyOf(cells);
      if (cells.isEmpty()) {
        throw new IllegalArgumentException("cells must not be empty");
      }
      Validation.requireDistinctFormulaCellTargets(cells);
    }
  }

  /** Clears all evaluator caches so later formula reads recompute from workbook state. */
  record ClearFormulaCaches() implements WorkbookOperation {}

  /** Marks the workbook so that Excel recalculates all formulas on next open. */
  record ForceFormulaRecalculationOnOpen() implements WorkbookOperation {}

  /** Returns the SCREAMING_SNAKE_CASE type name of this operation as used in the wire protocol. */
  default String operationType() {
    return switch (this) {
      case EnsureSheet _ -> "ENSURE_SHEET";
      case RenameSheet _ -> "RENAME_SHEET";
      case DeleteSheet _ -> "DELETE_SHEET";
      case MoveSheet _ -> "MOVE_SHEET";
      case CopySheet _ -> "COPY_SHEET";
      case SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case MergeCells _ -> "MERGE_CELLS";
      case UnmergeCells _ -> "UNMERGE_CELLS";
      case SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case SetRowHeight _ -> "SET_ROW_HEIGHT";
      case InsertRows _ -> "INSERT_ROWS";
      case DeleteRows _ -> "DELETE_ROWS";
      case ShiftRows _ -> "SHIFT_ROWS";
      case InsertColumns _ -> "INSERT_COLUMNS";
      case DeleteColumns _ -> "DELETE_COLUMNS";
      case ShiftColumns _ -> "SHIFT_COLUMNS";
      case SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case GroupRows _ -> "GROUP_ROWS";
      case UngroupRows _ -> "UNGROUP_ROWS";
      case GroupColumns _ -> "GROUP_COLUMNS";
      case UngroupColumns _ -> "UNGROUP_COLUMNS";
      case SetSheetPane _ -> "SET_SHEET_PANE";
      case SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case SetCell _ -> "SET_CELL";
      case SetRange _ -> "SET_RANGE";
      case ClearRange _ -> "CLEAR_RANGE";
      case SetHyperlink _ -> "SET_HYPERLINK";
      case ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case SetComment _ -> "SET_COMMENT";
      case ClearComment _ -> "CLEAR_COMMENT";
      case SetPicture _ -> "SET_PICTURE";
      case SetChart _ -> "SET_CHART";
      case SetShape _ -> "SET_SHAPE";
      case SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case ApplyStyle _ -> "APPLY_STYLE";
      case SetDataValidation _ -> "SET_DATA_VALIDATION";
      case ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case SetAutofilter _ -> "SET_AUTOFILTER";
      case ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case SetTable _ -> "SET_TABLE";
      case DeleteTable _ -> "DELETE_TABLE";
      case SetNamedRange _ -> "SET_NAMED_RANGE";
      case DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case AppendRow _ -> "APPEND_ROW";
      case AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case EvaluateFormulas _ -> "EVALUATE_FORMULAS";
      case EvaluateFormulaCells _ -> "EVALUATE_FORMULA_CELLS";
      case ClearFormulaCaches _ -> "CLEAR_FORMULA_CACHES";
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

    static void requireSheetName(String value, String fieldName) { // LIM-003
      ExcelSheetNames.requireValid(value, fieldName);
    }

    static void requireNonNegative(int value, String fieldName) {
      if (value < 0) {
        throw new IllegalArgumentException(fieldName + " must not be negative");
      }
    }

    static void requirePositive(int value, String fieldName) {
      if (value <= 0) {
        throw new IllegalArgumentException(fieldName + " must be greater than 0");
      }
    }

    static void requireNonZero(int value, String fieldName) {
      if (value == 0) {
        throw new IllegalArgumentException(fieldName + " must not be 0");
      }
    }

    static void requireRowIndex(int value, String fieldName) {
      requireNonNegative(value, fieldName);
      if (value > ExcelRowSpan.MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            fieldName + " must not exceed " + ExcelRowSpan.MAX_ROW_INDEX + " (Excel row limit)");
      }
    }

    static void requireColumnIndex(int value, String fieldName) {
      requireNonNegative(value, fieldName);
      if (value > ExcelColumnSpan.MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            fieldName
                + " must not exceed "
                + ExcelColumnSpan.MAX_COLUMN_INDEX
                + " (Excel column limit)");
      }
    }

    static void requireOrderedSpan(
        int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
      if (lastValue < firstValue) {
        throw new IllegalArgumentException(
            lastFieldName + " must not be less than " + firstFieldName);
      }
    }

    static void requireColumnWidthCharacters(double widthCharacters) { // LIM-004
      requireFinitePositive(widthCharacters, "widthCharacters");
      if (widthCharacters > 255.0d) {
        throw new IllegalArgumentException(
            "widthCharacters must not exceed 255.0 (Excel column width limit): got "
                + widthCharacters);
      }
      if (Math.round(widthCharacters * 256.0d) <= 0) {
        throw new IllegalArgumentException(
            "widthCharacters is too small to produce a visible Excel column width: got "
                + widthCharacters);
      }
    }

    static void requireRowHeightPoints(double heightPoints) { // LIM-005
      requireFinitePositive(heightPoints, "heightPoints");
      if (Math.round(heightPoints * 20.0d) > Short.MAX_VALUE) {
        throw new IllegalArgumentException(
            "heightPoints must not exceed 1638.35 (Excel storage limit: 32767 twips): got "
                + heightPoints);
      }
      if (Math.round(heightPoints * 20.0d) <= 0) {
        throw new IllegalArgumentException(
            "heightPoints is too small to produce a visible Excel row height: " + heightPoints);
      }
    }

    static void requireNamedRangeName(String name) {
      ProtocolDefinedNameValidation.validateName(name);
    }

    static void requireTableName(String name) {
      ProtocolDefinedNameValidation.validateName(name);
    }

    static void requireFinitePositive(double value, String fieldName) {
      if (!Double.isFinite(value)) {
        throw new IllegalArgumentException(fieldName + " must be finite");
      }
      if (value <= 0.0d) {
        throw new IllegalArgumentException(fieldName + " must be greater than 0");
      }
    }

    static void requireZoomPercent(int zoomPercent) {
      if (zoomPercent < 10 || zoomPercent > 400) {
        throw new IllegalArgumentException(
            "zoomPercent must be between 10 and 400 inclusive: " + zoomPercent);
      }
    }

    static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
      if (rows == null) {
        return List.of();
      }
      return rows.stream().map(row -> row == null ? null : List.copyOf(row)).toList();
    }

    static List<String> copySheetNames(List<String> sheetNames, String fieldName) {
      Objects.requireNonNull(sheetNames, fieldName + " must not be null");
      List<String> copy = List.copyOf(sheetNames);
      for (String sheetName : copy) {
        requireSheetName(sheetName, fieldName);
      }
      return copy;
    }

    static void requireDistinct(List<String> values, String fieldName) {
      if (new java.util.LinkedHashSet<>(values).size() != values.size()) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
    }

    static void requireDistinctFormulaCellTargets(List<FormulaCellTargetInput> cells) {
      Set<String> seen = new java.util.LinkedHashSet<>();
      for (FormulaCellTargetInput cell : cells) {
        Objects.requireNonNull(cell, "cells must not contain nulls");
        String key = (cell.sheetName() + "!" + cell.address()).toUpperCase(java.util.Locale.ROOT);
        if (!seen.add(key)) {
          throw new IllegalArgumentException(
              "cells must not contain duplicate sheetName/address targets: "
                  + cell.sheetName()
                  + "!"
                  + cell.address());
        }
      }
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
