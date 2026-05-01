package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** One validated mutation action expressed in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WorkbookMutationAction.EnsureSheet.class, name = "ENSURE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.RenameSheet.class, name = "RENAME_SHEET"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.DeleteSheet.class, name = "DELETE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.MoveSheet.class, name = "MOVE_SHEET"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.CopySheet.class, name = "COPY_SHEET"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetActiveSheet.class,
      name = "SET_ACTIVE_SHEET"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetSelectedSheets.class,
      name = "SET_SELECTED_SHEETS"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetSheetVisibility.class,
      name = "SET_SHEET_VISIBILITY"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetSheetProtection.class,
      name = "SET_SHEET_PROTECTION"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.ClearSheetProtection.class,
      name = "CLEAR_SHEET_PROTECTION"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetWorkbookProtection.class,
      name = "SET_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.ClearWorkbookProtection.class,
      name = "CLEAR_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.MergeCells.class, name = "MERGE_CELLS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.UnmergeCells.class, name = "UNMERGE_CELLS"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetColumnWidth.class,
      name = "SET_COLUMN_WIDTH"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.SetRowHeight.class, name = "SET_ROW_HEIGHT"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.InsertRows.class, name = "INSERT_ROWS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.DeleteRows.class, name = "DELETE_ROWS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.ShiftRows.class, name = "SHIFT_ROWS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.InsertColumns.class, name = "INSERT_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.DeleteColumns.class, name = "DELETE_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.ShiftColumns.class, name = "SHIFT_COLUMNS"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetRowVisibility.class,
      name = "SET_ROW_VISIBILITY"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetColumnVisibility.class,
      name = "SET_COLUMN_VISIBILITY"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.GroupRows.class, name = "GROUP_ROWS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.UngroupRows.class, name = "UNGROUP_ROWS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.GroupColumns.class, name = "GROUP_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.UngroupColumns.class, name = "UNGROUP_COLUMNS"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.SetSheetPane.class, name = "SET_SHEET_PANE"),
  @JsonSubTypes.Type(value = WorkbookMutationAction.SetSheetZoom.class, name = "SET_SHEET_ZOOM"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetSheetPresentation.class,
      name = "SET_SHEET_PRESENTATION"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.SetPrintLayout.class,
      name = "SET_PRINT_LAYOUT"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.ClearPrintLayout.class,
      name = "CLEAR_PRINT_LAYOUT"),
  @JsonSubTypes.Type(value = CellMutationAction.SetCell.class, name = "SET_CELL"),
  @JsonSubTypes.Type(value = CellMutationAction.SetRange.class, name = "SET_RANGE"),
  @JsonSubTypes.Type(value = CellMutationAction.ClearRange.class, name = "CLEAR_RANGE"),
  @JsonSubTypes.Type(value = CellMutationAction.SetArrayFormula.class, name = "SET_ARRAY_FORMULA"),
  @JsonSubTypes.Type(
      value = CellMutationAction.ClearArrayFormula.class,
      name = "CLEAR_ARRAY_FORMULA"),
  @JsonSubTypes.Type(value = CellMutationAction.SetHyperlink.class, name = "SET_HYPERLINK"),
  @JsonSubTypes.Type(value = CellMutationAction.ClearHyperlink.class, name = "CLEAR_HYPERLINK"),
  @JsonSubTypes.Type(value = CellMutationAction.SetComment.class, name = "SET_COMMENT"),
  @JsonSubTypes.Type(value = CellMutationAction.ClearComment.class, name = "CLEAR_COMMENT"),
  @JsonSubTypes.Type(value = CellMutationAction.ApplyStyle.class, name = "APPLY_STYLE"),
  @JsonSubTypes.Type(value = CellMutationAction.AppendRow.class, name = "APPEND_ROW"),
  @JsonSubTypes.Type(value = DrawingMutationAction.SetPicture.class, name = "SET_PICTURE"),
  @JsonSubTypes.Type(
      value = DrawingMutationAction.SetSignatureLine.class,
      name = "SET_SIGNATURE_LINE"),
  @JsonSubTypes.Type(value = DrawingMutationAction.SetChart.class, name = "SET_CHART"),
  @JsonSubTypes.Type(value = DrawingMutationAction.SetShape.class, name = "SET_SHAPE"),
  @JsonSubTypes.Type(
      value = DrawingMutationAction.SetEmbeddedObject.class,
      name = "SET_EMBEDDED_OBJECT"),
  @JsonSubTypes.Type(
      value = DrawingMutationAction.SetDrawingObjectAnchor.class,
      name = "SET_DRAWING_OBJECT_ANCHOR"),
  @JsonSubTypes.Type(
      value = DrawingMutationAction.DeleteDrawingObject.class,
      name = "DELETE_DRAWING_OBJECT"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.ImportCustomXmlMapping.class,
      name = "IMPORT_CUSTOM_XML_MAPPING"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.SetPivotTable.class,
      name = "SET_PIVOT_TABLE"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.SetDataValidation.class,
      name = "SET_DATA_VALIDATION"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.ClearDataValidations.class,
      name = "CLEAR_DATA_VALIDATIONS"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.SetConditionalFormatting.class,
      name = "SET_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.ClearConditionalFormatting.class,
      name = "CLEAR_CONDITIONAL_FORMATTING"),
  @JsonSubTypes.Type(value = StructuredMutationAction.SetAutofilter.class, name = "SET_AUTOFILTER"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.ClearAutofilter.class,
      name = "CLEAR_AUTOFILTER"),
  @JsonSubTypes.Type(value = StructuredMutationAction.SetTable.class, name = "SET_TABLE"),
  @JsonSubTypes.Type(value = StructuredMutationAction.DeleteTable.class, name = "DELETE_TABLE"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.DeletePivotTable.class,
      name = "DELETE_PIVOT_TABLE"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.SetNamedRange.class,
      name = "SET_NAMED_RANGE"),
  @JsonSubTypes.Type(
      value = StructuredMutationAction.DeleteNamedRange.class,
      name = "DELETE_NAMED_RANGE"),
  @JsonSubTypes.Type(
      value = WorkbookMutationAction.AutoSizeColumns.class,
      name = "AUTO_SIZE_COLUMNS")
})
public sealed interface MutationAction
    permits WorkbookMutationAction,
        CellMutationAction,
        DrawingMutationAction,
        StructuredMutationAction {
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
      // LIM-008
      requireNonNegative(value, fieldName);
      if (value > ExcelRowSpan.MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            fieldName + " must not exceed " + ExcelRowSpan.MAX_ROW_INDEX + " (Excel row limit)");
      }
    }

    static void requireColumnIndex(int value, String fieldName) {
      // LIM-009
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
      ExcelSheetLayoutLimits.requireColumnWidthCharacters(widthCharacters, "widthCharacters");
    }

    static void requireRowHeightPoints(double heightPoints) { // LIM-005
      ExcelSheetLayoutLimits.requireRowHeightPoints(heightPoints, "heightPoints");
    }

    static void requireNamedRangeName(String name) {
      ProtocolDefinedNameValidation.validateName(name);
    }

    static void requirePivotTableName(String name) {
      ExcelPivotTableNaming.validateName(name);
    }

    static void requireZoomPercent(int zoomPercent) { // LIM-022
      ExcelSheetLayoutLimits.requireZoomPercent(zoomPercent, "zoomPercent");
    }

    static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
      Objects.requireNonNull(rows, "rows must not be null");
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
