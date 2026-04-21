package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.excel.ExcelColumnSpan;
import dev.erst.gridgrind.excel.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelSheetNames;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * One validated mutation action expressed in protocol form.
 *
 * <p>This closed sum type intentionally imports many payload records because it is the canonical
 * contract owner for every mutation family.
 */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.ExcessivePublicCount"})
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = MutationAction.EnsureSheet.class, name = "ENSURE_SHEET"),
  @JsonSubTypes.Type(value = MutationAction.RenameSheet.class, name = "RENAME_SHEET"),
  @JsonSubTypes.Type(value = MutationAction.DeleteSheet.class, name = "DELETE_SHEET"),
  @JsonSubTypes.Type(value = MutationAction.MoveSheet.class, name = "MOVE_SHEET"),
  @JsonSubTypes.Type(value = MutationAction.CopySheet.class, name = "COPY_SHEET"),
  @JsonSubTypes.Type(value = MutationAction.SetActiveSheet.class, name = "SET_ACTIVE_SHEET"),
  @JsonSubTypes.Type(value = MutationAction.SetSelectedSheets.class, name = "SET_SELECTED_SHEETS"),
  @JsonSubTypes.Type(
      value = MutationAction.SetSheetVisibility.class,
      name = "SET_SHEET_VISIBILITY"),
  @JsonSubTypes.Type(
      value = MutationAction.SetSheetProtection.class,
      name = "SET_SHEET_PROTECTION"),
  @JsonSubTypes.Type(
      value = MutationAction.ClearSheetProtection.class,
      name = "CLEAR_SHEET_PROTECTION"),
  @JsonSubTypes.Type(
      value = MutationAction.SetWorkbookProtection.class,
      name = "SET_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(
      value = MutationAction.ClearWorkbookProtection.class,
      name = "CLEAR_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(value = MutationAction.MergeCells.class, name = "MERGE_CELLS"),
  @JsonSubTypes.Type(value = MutationAction.UnmergeCells.class, name = "UNMERGE_CELLS"),
  @JsonSubTypes.Type(value = MutationAction.SetColumnWidth.class, name = "SET_COLUMN_WIDTH"),
  @JsonSubTypes.Type(value = MutationAction.SetRowHeight.class, name = "SET_ROW_HEIGHT"),
  @JsonSubTypes.Type(value = MutationAction.InsertRows.class, name = "INSERT_ROWS"),
  @JsonSubTypes.Type(value = MutationAction.DeleteRows.class, name = "DELETE_ROWS"),
  @JsonSubTypes.Type(value = MutationAction.ShiftRows.class, name = "SHIFT_ROWS"),
  @JsonSubTypes.Type(value = MutationAction.InsertColumns.class, name = "INSERT_COLUMNS"),
  @JsonSubTypes.Type(value = MutationAction.DeleteColumns.class, name = "DELETE_COLUMNS"),
  @JsonSubTypes.Type(value = MutationAction.ShiftColumns.class, name = "SHIFT_COLUMNS"),
  @JsonSubTypes.Type(value = MutationAction.SetRowVisibility.class, name = "SET_ROW_VISIBILITY"),
  @JsonSubTypes.Type(
      value = MutationAction.SetColumnVisibility.class,
      name = "SET_COLUMN_VISIBILITY"),
  @JsonSubTypes.Type(value = MutationAction.GroupRows.class, name = "GROUP_ROWS"),
  @JsonSubTypes.Type(value = MutationAction.UngroupRows.class, name = "UNGROUP_ROWS"),
  @JsonSubTypes.Type(value = MutationAction.GroupColumns.class, name = "GROUP_COLUMNS"),
  @JsonSubTypes.Type(value = MutationAction.UngroupColumns.class, name = "UNGROUP_COLUMNS"),
  @JsonSubTypes.Type(value = MutationAction.SetSheetPane.class, name = "SET_SHEET_PANE"),
  @JsonSubTypes.Type(value = MutationAction.SetSheetZoom.class, name = "SET_SHEET_ZOOM"),
  @JsonSubTypes.Type(
      value = MutationAction.SetSheetPresentation.class,
      name = "SET_SHEET_PRESENTATION"),
  @JsonSubTypes.Type(value = MutationAction.SetPrintLayout.class, name = "SET_PRINT_LAYOUT"),
  @JsonSubTypes.Type(value = MutationAction.ClearPrintLayout.class, name = "CLEAR_PRINT_LAYOUT"),
  @JsonSubTypes.Type(value = MutationAction.SetCell.class, name = "SET_CELL"),
  @JsonSubTypes.Type(value = MutationAction.SetRange.class, name = "SET_RANGE"),
  @JsonSubTypes.Type(value = MutationAction.ClearRange.class, name = "CLEAR_RANGE"),
  @JsonSubTypes.Type(value = MutationAction.SetHyperlink.class, name = "SET_HYPERLINK"),
  @JsonSubTypes.Type(value = MutationAction.ClearHyperlink.class, name = "CLEAR_HYPERLINK"),
  @JsonSubTypes.Type(value = MutationAction.SetComment.class, name = "SET_COMMENT"),
  @JsonSubTypes.Type(value = MutationAction.ClearComment.class, name = "CLEAR_COMMENT"),
  @JsonSubTypes.Type(value = MutationAction.SetPicture.class, name = "SET_PICTURE"),
  @JsonSubTypes.Type(value = MutationAction.SetChart.class, name = "SET_CHART"),
  @JsonSubTypes.Type(value = MutationAction.SetPivotTable.class, name = "SET_PIVOT_TABLE"),
  @JsonSubTypes.Type(value = MutationAction.SetShape.class, name = "SET_SHAPE"),
  @JsonSubTypes.Type(value = MutationAction.SetEmbeddedObject.class, name = "SET_EMBEDDED_OBJECT"),
  @JsonSubTypes.Type(
      value = MutationAction.SetDrawingObjectAnchor.class,
      name = "SET_DRAWING_OBJECT_ANCHOR"),
  @JsonSubTypes.Type(
      value = MutationAction.DeleteDrawingObject.class,
      name = "DELETE_DRAWING_OBJECT"),
  @JsonSubTypes.Type(value = MutationAction.ApplyStyle.class, name = "APPLY_STYLE"),
  @JsonSubTypes.Type(value = MutationAction.SetDataValidation.class, name = "SET_DATA_VALIDATION"),
  @JsonSubTypes.Type(
      value = MutationAction.ClearDataValidations.class,
      name = "CLEAR_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = MutationAction.SetConditionalFormatting.class,
      name = "SET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(
      value = MutationAction.ClearConditionalFormatting.class,
      name = "CLEAR_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = MutationAction.SetAutofilter.class, name = "SET_AUTOFILTER"),
  @JsonSubTypes.Type(value = MutationAction.ClearAutofilter.class, name = "CLEAR_AUTOFILTER"),
  @JsonSubTypes.Type(value = MutationAction.SetTable.class, name = "SET_TABLE"),
  @JsonSubTypes.Type(value = MutationAction.DeleteTable.class, name = "DELETE_TABLE"),
  @JsonSubTypes.Type(value = MutationAction.DeletePivotTable.class, name = "DELETE_PIVOT_TABLE"),
  @JsonSubTypes.Type(value = MutationAction.SetNamedRange.class, name = "SET_NAMED_RANGE"),
  @JsonSubTypes.Type(value = MutationAction.DeleteNamedRange.class, name = "DELETE_NAMED_RANGE"),
  @JsonSubTypes.Type(value = MutationAction.AppendRow.class, name = "APPEND_ROW"),
  @JsonSubTypes.Type(value = MutationAction.AutoSizeColumns.class, name = "AUTO_SIZE_COLUMNS")
})
public sealed interface MutationAction {

  /** Ensures a sheet with the given name exists, creating it if absent. */
  record EnsureSheet() implements MutationAction {
    public EnsureSheet {}
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String newSheetName) implements MutationAction {
    public RenameSheet {
      Validation.requireSheetName(newSheetName, "newSheetName");
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet() implements MutationAction {
    public DeleteSheet {}
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(Integer targetIndex) implements MutationAction {
    public MoveSheet {
      Objects.requireNonNull(targetIndex, "targetIndex must not be null");
      Validation.requireNonNegative(targetIndex, "targetIndex");
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  record CopySheet(String newSheetName, SheetCopyPosition position) implements MutationAction {
    public CopySheet {
      Validation.requireSheetName(newSheetName, "newSheetName");
      position = position == null ? new SheetCopyPosition.AppendAtEnd() : position;
    }
  }

  /** Sets the active sheet and ensures it is selected. */
  record SetActiveSheet() implements MutationAction {
    public SetActiveSheet {}
  }

  /** Sets the selected visible sheet set. */
  record SetSelectedSheets() implements MutationAction {
    public SetSelectedSheets {}
  }

  /** Sets one sheet visibility. */
  record SetSheetVisibility(ExcelSheetVisibility visibility) implements MutationAction {
    public SetSheetVisibility {
      Objects.requireNonNull(visibility, "visibility must not be null");
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  record SetSheetProtection(SheetProtectionSettings protection, String password)
      implements MutationAction {
    /** Enables sheet protection without applying a password hash. */
    public SetSheetProtection(SheetProtectionSettings protection) {
      this(protection, null);
    }

    public SetSheetProtection {
      Objects.requireNonNull(protection, "protection must not be null");
      if (password != null && password.isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }
  }

  /** Disables sheet protection entirely. */
  record ClearSheetProtection() implements MutationAction {
    public ClearSheetProtection {}
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  record SetWorkbookProtection(WorkbookProtectionInput protection) implements MutationAction {
    public SetWorkbookProtection {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Clears workbook-level protection and password hashes entirely. */
  record ClearWorkbookProtection() implements MutationAction {
    public ClearWorkbookProtection {}
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells() implements MutationAction {
    public MergeCells {}
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  record UnmergeCells() implements MutationAction {
    public UnmergeCells {}
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  record SetColumnWidth(Double widthCharacters) implements MutationAction {
    public SetColumnWidth {
      Objects.requireNonNull(widthCharacters, "widthCharacters must not be null");
      Validation.requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  record SetRowHeight(Double heightPoints) implements MutationAction {
    public SetRowHeight {
      Objects.requireNonNull(heightPoints, "heightPoints must not be null");
      Validation.requireRowHeightPoints(heightPoints);
    }
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  record InsertRows() implements MutationAction {
    public InsertRows {}
  }

  /** Deletes the requested inclusive zero-based row band. */
  record DeleteRows() implements MutationAction {
    public DeleteRows {}
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  record ShiftRows(Integer delta) implements MutationAction {
    public ShiftRows {
      Objects.requireNonNull(delta, "delta must not be null");
      Validation.requireNonZero(delta, "delta");
    }
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  record InsertColumns() implements MutationAction {
    public InsertColumns {}
  }

  /** Deletes the requested inclusive zero-based column band. */
  record DeleteColumns() implements MutationAction {
    public DeleteColumns {}
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  record ShiftColumns(Integer delta) implements MutationAction {
    public ShiftColumns {
      Objects.requireNonNull(delta, "delta must not be null");
      Validation.requireNonZero(delta, "delta");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  record SetRowVisibility(Boolean hidden) implements MutationAction {
    public SetRowVisibility {
      Objects.requireNonNull(hidden, "hidden must not be null");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  record SetColumnVisibility(Boolean hidden) implements MutationAction {
    public SetColumnVisibility {
      Objects.requireNonNull(hidden, "hidden must not be null");
    }
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
  record GroupRows(Boolean collapsed) implements MutationAction {
    public GroupRows {
      collapsed = collapsed == null ? Boolean.FALSE : collapsed;
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  record UngroupRows() implements MutationAction {
    public UngroupRows {}
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  record GroupColumns(Boolean collapsed) implements MutationAction {
    public GroupColumns {
      collapsed = collapsed == null ? Boolean.FALSE : collapsed;
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  record UngroupColumns() implements MutationAction {
    public UngroupColumns {}
  }

  /** Applies one explicit pane state to a sheet. */
  record SetSheetPane(PaneInput pane) implements MutationAction {
    public SetSheetPane {
      Objects.requireNonNull(pane, "pane must not be null");
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  record SetSheetZoom(Integer zoomPercent) implements MutationAction {
    public SetSheetZoom {
      Objects.requireNonNull(zoomPercent, "zoomPercent must not be null");
      Validation.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  record SetSheetPresentation(SheetPresentationInput presentation) implements MutationAction {
    public SetSheetPresentation {
      Objects.requireNonNull(presentation, "presentation must not be null");
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  record SetPrintLayout(PrintLayoutInput printLayout) implements MutationAction {
    public SetPrintLayout {
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  record ClearPrintLayout() implements MutationAction {
    public ClearPrintLayout {}
  }

  /** Sets a single cell to the given value. */
  record SetCell(CellInput value) implements MutationAction {
    public SetCell {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Sets a rectangular region of cells from a row-major grid of values. */
  record SetRange(List<List<CellInput>> rows) implements MutationAction {
    public SetRange {
      rows = Validation.copyRows(rows);
      Validation.requireRectangularRows(rows);
      rows = Validation.freezeRows(rows);
    }
  }

  /** Clears all cell values and styles within the specified range. */
  record ClearRange() implements MutationAction {
    public ClearRange {}
  }

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(HyperlinkTarget target) implements MutationAction {
    public SetHyperlink {
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink() implements MutationAction {
    public ClearHyperlink {}
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(CommentInput comment) implements MutationAction {
    public SetComment {
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment() implements MutationAction {
    public ClearComment {}
  }

  /** Creates or replaces one picture-backed drawing object on one sheet. */
  record SetPicture(PictureInput picture) implements MutationAction {
    public SetPicture {
      Objects.requireNonNull(picture, "picture must not be null");
    }
  }

  /** Creates or mutates one supported simple chart on one sheet. */
  record SetChart(ChartInput chart) implements MutationAction {
    public SetChart {
      Objects.requireNonNull(chart, "chart must not be null");
    }
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  record SetPivotTable(PivotTableInput pivotTable) implements MutationAction {
    public SetPivotTable {
      Objects.requireNonNull(pivotTable, "pivotTable must not be null");
    }
  }

  /** Creates or replaces one simple-shape or connector drawing object on one sheet. */
  record SetShape(ShapeInput shape) implements MutationAction {
    public SetShape {
      Objects.requireNonNull(shape, "shape must not be null");
    }
  }

  /** Creates or replaces one embedded-object drawing object on one sheet. */
  record SetEmbeddedObject(EmbeddedObjectInput embeddedObject) implements MutationAction {
    public SetEmbeddedObject {
      Objects.requireNonNull(embeddedObject, "embeddedObject must not be null");
    }
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  record SetDrawingObjectAnchor(DrawingAnchorInput anchor) implements MutationAction {
    public SetDrawingObjectAnchor {
      Objects.requireNonNull(anchor, "anchor must not be null");
    }
  }

  /** Deletes one existing drawing object by sheet-local name. */
  record DeleteDrawingObject() implements MutationAction {
    public DeleteDrawingObject {}
  }

  /** Applies a style patch to every cell in the specified range. */
  record ApplyStyle(CellStyleInput style) implements MutationAction {
    public ApplyStyle {
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  record SetDataValidation(DataValidationInput validation) implements MutationAction {
    public SetDataValidation {
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  record ClearDataValidations() implements MutationAction {
    public ClearDataValidations {}
  }

  /** Creates or replaces one logical conditional-formatting block over the requested ranges. */
  record SetConditionalFormatting(ConditionalFormattingBlockInput conditionalFormatting)
      implements MutationAction {
    public SetConditionalFormatting {
      Objects.requireNonNull(conditionalFormatting, "conditionalFormatting must not be null");
    }
  }

  /** Removes conditional-formatting blocks on the sheet that match the provided range selection. */
  record ClearConditionalFormatting() implements MutationAction {
    public ClearConditionalFormatting {}
  }

  /** Creates or replaces one sheet-level autofilter range. */
  record SetAutofilter(
      List<AutofilterFilterColumnInput> criteria, AutofilterSortStateInput sortState)
      implements MutationAction {
    /** Creates a plain sheet-level autofilter without criteria or explicit sort state. */
    public SetAutofilter() {
      this(List.of(), null);
    }

    public SetAutofilter {
      criteria = criteria == null ? List.of() : new ArrayList<>(criteria);
      for (AutofilterFilterColumnInput criterion : criteria) {
        Objects.requireNonNull(criterion, "criteria must not contain null values");
      }
      criteria = List.copyOf(criteria);
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  record ClearAutofilter() implements MutationAction {
    public ClearAutofilter {}
  }

  /** Creates or replaces one workbook-global table definition. */
  record SetTable(TableInput table) implements MutationAction {
    public SetTable {
      Objects.requireNonNull(table, "table must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  record DeleteTable() implements MutationAction {
    public DeleteTable {}
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet. */
  record DeletePivotTable() implements MutationAction {
    public DeletePivotTable {}
  }

  /** Creates or replaces one typed named range in workbook or sheet scope. */
  record SetNamedRange(String name, NamedRangeScope scope, NamedRangeTarget target)
      implements MutationAction {
    public SetNamedRange {
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Validation.requireNamedRangeName(name);
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  record DeleteNamedRange() implements MutationAction {
    public DeleteNamedRange {}
  }

  /** Appends a new row of values after the last occupied row on the sheet. */
  record AppendRow(List<CellInput> values) implements MutationAction {
    public AppendRow {
      values = values == null ? List.of() : new ArrayList<>(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (CellInput item : values) {
        Objects.requireNonNull(item, "values must not contain nulls");
      }
      values = List.copyOf(values);
    }
  }

  /** Auto-sizes all populated columns on the sheet to fit their content. */
  record AutoSizeColumns() implements MutationAction {
    public AutoSizeColumns {}
  }

  /** Returns the SCREAMING_SNAKE_CASE type name of this action as used in the wire protocol. */
  default String actionType() {
    return GridGrindProtocolTypeNames.mutationActionTypeName(
        getClass().asSubclass(MutationAction.class));
  }

  /** Shared validation helpers for MutationAction compact constructors. */
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

    static void requirePivotTableName(String name) {
      ExcelPivotTableNaming.validateName(name);
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
      List<List<CellInput>> copy = new ArrayList<>(rows.size());
      for (List<CellInput> row : rows) {
        copy.add(row == null ? null : new ArrayList<>(row));
      }
      return java.util.Collections.unmodifiableList(copy);
    }

    static List<List<CellInput>> freezeRows(List<List<CellInput>> rows) {
      return rows.stream().map(List::copyOf).toList();
    }

    static List<String> copySheetNames(List<String> sheetNames, String fieldName) {
      Objects.requireNonNull(sheetNames, fieldName + " must not be null");
      List<String> copy = new ArrayList<>(sheetNames);
      for (String sheetName : copy) {
        requireSheetName(sheetName, fieldName);
      }
      return List.copyOf(copy);
    }

    static void requireDistinct(List<String> values, String fieldName) {
      if (new java.util.LinkedHashSet<>(values).size() != values.size()) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
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
